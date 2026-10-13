import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { normalizeTitle } from "./taskQuery.js";

/**
 * Turns a meal plan into a list of things to check before shopping.
 *
 * Recipe tasks carry a JSON note: `dates_planned` (YYYY-MM-DD), `kød` and
 * `ingredienser`, each a list of `{ vare: [alternatives, preferred first],
 * mængde }`. For every recipe planned inside the window this collects what
 * it needs, drops meat that is in the freezer list and staples on the
 * pantry list, skips what is already open on the shopping list, and writes
 * the rest as tasks in a staging list. Someone ticks off what the household
 * already has; what is left is moved to the shopping list by Home
 * Assistant.
 *
 * Shopping lists keep their completed items so the same item is reopened
 * rather than added again. Their titles are free text ("Flåede tomater – 2
 * dåser"), so matching compares a reduced key: the part before the first
 * dash/comma/parenthesis, without amounts or units, with simple Danish
 * plural endings stripped. When a completed item matches, its id goes into
 * the staging task's note as `shopping_id: <id>` for Home Assistant to
 * reopen.
 */

export interface Ingredient {
  vare: string[] | string;
  mængde?: string;
}

export interface RecipeNote {
  dates_planned?: string[];
  kategori?: string;
  kød?: Ingredient[];
  ingredienser?: Ingredient[];
  ingredienser_usikker?: boolean;
}

const ENTITIES: Record<string, string> = { "&quot;": '"', "&#39;": "'", "&lt;": "<", "&gt;": ">", "&nbsp;": " ", "&amp;": "&" };

/**
 * The note as plain text; notes written through Graph since the HTML-only
 * change come back as HTML. Only parsed as JSON, never rendered, but tags are
 * still stripped until none are left so nested ones ("<<b>script>") can't
 * survive a single pass.
 */
export function noteText(body?: { content: string; contentType: string }): string {
  if (!body?.content) return "";
  if (body.contentType !== "html") return body.content;
  let text = body.content.replace(/<br\s*\/?>|<\/p>|<\/div>/gi, "\n");
  let previous: string;
  do {
    previous = text;
    text = text.replace(/<[^<>]*>/g, "");
  } while (text !== previous);
  return text.replace(/&(quot|#39|lt|gt|nbsp|amp);/g, (m) => ENTITIES[m] ?? m);
}

export function parseRecipeNote(body?: { content: string; contentType: string }): RecipeNote | undefined {
  const text = noteText(body);
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return undefined;
  try {
    const parsed = JSON.parse(text.slice(start, end + 1));
    return parsed && typeof parsed === "object" ? (parsed as RecipeNote) : undefined;
  } catch {
    return undefined;
  }
}

const UNIT_WORDS = new Set(
  "ca stk g gr kg ml cl dl l liter tsk spsk knsp dase daser pakke pakker pk bakke bakker pose poser glas bundt fed skive skiver x evt"
    .split(" ")
);

function stem(word: string): string {
  if (word.length <= 3) return word;
  for (const ending of ["erne", "ene", "er", "e"]) {
    if (word.endsWith(ending) && word.length - ending.length >= 3) return word.slice(0, -ending.length);
  }
  return word;
}

/** "Flåede tomater – 2 dåser" → "flaed tomat". Empty when nothing item-like is left. */
export function itemKey(text: string): string {
  const head = text.split(/\s[–-]\s|[,(:/]/)[0] ?? "";
  return normalizeTitle(head)
    .split(" ")
    .filter((w) => w && !UNIT_WORDS.has(w) && !/\d/.test(w))
    .map(stem)
    .join(" ");
}

/**
 * True when `title` names the same item as one of `alternatives`: equal keys,
 * or the title is the alternative plus at most one extra word ("Champignon"
 * matches "Brune champignon"). Never the other way round, so "Tomater" does
 * not stand in for "Hakkede tomater".
 */
export function matchesItem(alternatives: string[], title: string, opts: { eitherWay?: boolean; exact?: boolean } = {}): boolean {
  const titleWords = itemKey(title).split(" ").filter(Boolean);
  if (!titleWords.length) return false;
  return alternatives.some((alt) => {
    const altWords = itemKey(alt).split(" ").filter(Boolean);
    if (!altWords.length) return false;
    if (opts.exact) return altWords.join(" ") === titleWords.join(" ");
    const covers = (outer: string[], inner: string[]) => inner.every((w) => outer.includes(w)) && outer.length - inner.length <= 1;
    return covers(titleWords, altWords) || (opts.eitherWay === true && covers(altWords, titleWords));
  });
}

/** Freezer titles are "<Placering> – <Vare> – <mængde> – pakke n/m"; the item is the second part. */
function freezerItem(title: string): string {
  const parts = title.split(/\s[–-]\s/);
  return parts.length >= 2 ? parts[1]! : title;
}

const alternativesOf = (i: Ingredient): string[] =>
  (Array.isArray(i.vare) ? i.vare : [i.vare]).map((v) => String(v ?? "").trim()).filter(Boolean);

const WEEKDAYS = ["søn", "man", "tir", "ons", "tor", "fre", "lør"];

function dayLabel(date: string): string {
  const d = new Date(`${date}T12:00:00Z`);
  return `${WEEKDAYS[d.getUTCDay()]} ${d.getUTCDate()}/${d.getUTCMonth() + 1}`;
}

export interface ShoppingCandidateOptions {
  recipeListId: string;
  from: string;
  to: string;
  stagingListId: string;
  shoppingListId: string;
  pantryListId?: string;
  freezerListId?: string;
  dryRun: boolean;
}

interface Use {
  recipe: string;
  date: string;
  mængde?: string;
}

interface Need {
  vare: string;
  alternatives: string[];
  uses: Use[];
}

export type CandidateStatus = "staged" | "preview" | "already_on_shopping" | "already_staged" | "error";

export interface CandidateResult {
  vare: string;
  status: CandidateStatus;
  uses: string[];
  /** Completed shopping task Home Assistant should reopen instead of adding a new one. */
  reuse_shopping_task?: { id: string; title: string };
  matched_title?: string;
  staged_task_id?: string;
}

const useLine = (u: Use) => `${u.mængde ? `${u.mængde} – ` : ""}${u.recipe} (${dayLabel(u.date)})`;

export function stagingNote(need: Need, reuse?: { id: string; title: string }): string {
  const lines = need.uses.map(useLine);
  if (reuse) lines.push(`Genbruger Shopping-varen "${reuse.title}"`, `shopping_id: ${reuse.id}`);
  else lines.push("Ny vare på Shopping");
  return lines.join("\n");
}

export async function buildShoppingCandidates(client: Client, o: ShoppingCandidateOptions) {
  if (o.from > o.to) throw new Error("from must be on or before to.");

  const [recipes, shopping, staging, pantry, freezer] = await Promise.all([
    todo.listTasks(client, o.recipeListId, { maxItems: 2000 }),
    todo.listTasks(client, o.shoppingListId, { compact: true, maxItems: 5000 }),
    todo.listTasks(client, o.stagingListId, { compact: true, maxItems: 1000 }),
    o.pantryListId ? todo.listTasks(client, o.pantryListId, { compact: true, maxItems: 1000 }) : Promise.resolve([]),
    o.freezerListId ? todo.listTasks(client, o.freezerListId, { compact: true, maxItems: 1000 }) : Promise.resolve([]),
  ]);
  const isOpen = (t: { status: string }) => t.status !== "completed";
  const openFreezer = freezer.filter(isOpen);
  const pantryTitles = pantry.filter(isOpen).map((t) => t.title);

  const meals: Array<{ recipe: string; date: string }> = [];
  const missingIngredients: Array<{ recipe: string; date: string; reason: string }> = [];
  const fromFreezer: Array<{ vare: string; recipe: string; date: string; freezer_item: string }> = [];
  const pantrySkipped = new Set<string>();
  const needs = new Map<string, Need>();

  const addNeed = (alternatives: string[], use: Use) => {
    const key = itemKey(alternatives[0]!) || normalizeTitle(alternatives[0]!);
    const existing = needs.get(key);
    if (existing) existing.uses.push(use);
    else needs.set(key, { vare: alternatives[0]!, alternatives, uses: [use] });
  };

  for (const recipe of recipes) {
    const note = parseRecipeNote(recipe.body);
    const dates = (note?.dates_planned ?? []).filter((d) => typeof d === "string" && d >= o.from && d <= o.to);
    for (const date of dates.sort()) {
      meals.push({ recipe: recipe.title, date });
      for (const meat of note?.kød ?? []) {
        const alternatives = alternativesOf(meat);
        if (!alternatives.length) continue;
        const inFreezer = openFreezer.find((f) => matchesItem(alternatives, freezerItem(f.title), { eitherWay: true }));
        if (inFreezer) fromFreezer.push({ vare: alternatives[0]!, recipe: recipe.title, date, freezer_item: inFreezer.title });
        else addNeed(alternatives, { recipe: recipe.title, date, mængde: meat.mængde });
      }
      if (!Array.isArray(note?.ingredienser)) {
        if (note?.kategori !== "rester") missingIngredients.push({ recipe: recipe.title, date, reason: note ? "no ingredienser" : "note is not recipe JSON" });
        continue;
      }
      for (const ingredient of note.ingredienser) {
        const alternatives = alternativesOf(ingredient);
        if (!alternatives.length) continue;
        // Exact only: a pantry "Paprika" (the spice) must not swallow "Rød paprika".
        if (pantryTitles.some((p) => matchesItem(alternatives, p, { exact: true }))) {
          pantrySkipped.add(alternatives[0]!);
          continue;
        }
        addNeed(alternatives, { recipe: recipe.title, date, mængde: ingredient.mængde });
      }
    }
  }

  const openStaging = staging.filter(isOpen);
  const completedShopping = shopping
    .filter((t) => !isOpen(t))
    .sort((a, b) => (b.completedDateTime?.iso ?? "").localeCompare(a.completedDateTime?.iso ?? ""));

  const results: CandidateResult[] = [];
  for (const need of needs.values()) {
    need.uses.sort((a, b) => a.date.localeCompare(b.date));
    const base = { vare: need.vare, uses: need.uses.map(useLine) };
    const onShopping = shopping.find((t) => isOpen(t) && matchesItem(need.alternatives, t.title));
    if (onShopping) {
      results.push({ ...base, status: "already_on_shopping", matched_title: onShopping.title });
      continue;
    }
    const staged = openStaging.find((t) => matchesItem(need.alternatives, t.title));
    if (staged) {
      results.push({ ...base, status: "already_staged", matched_title: staged.title });
      continue;
    }
    const previous = completedShopping.find((t) => matchesItem(need.alternatives, t.title));
    const reuse = previous ? { id: previous.id, title: previous.title } : undefined;
    if (o.dryRun) {
      results.push({ ...base, status: "preview", reuse_shopping_task: reuse });
      continue;
    }
    try {
      const title = need.vare.charAt(0).toUpperCase() + need.vare.slice(1);
      const created = await todo.createTask(client, o.stagingListId, { title, body: stagingNote(need, reuse) });
      results.push({ ...base, status: "staged", reuse_shopping_task: reuse, staged_task_id: created.id });
    } catch {
      // Never surface raw Graph errors (they can contain internal data).
      results.push({ ...base, status: "error", reuse_shopping_task: reuse });
    }
  }

  const count = (s: CandidateStatus) => results.filter((r) => r.status === s).length;
  return {
    dry_run: o.dryRun,
    window: { from: o.from, to: o.to },
    meals: meals.sort((a, b) => a.date.localeCompare(b.date)),
    staged: count("staged"),
    preview: count("preview"),
    reused_from_shopping: results.filter((r) => r.reuse_shopping_task && (r.status === "staged" || r.status === "preview")).length,
    already_on_shopping: count("already_on_shopping"),
    already_staged: count("already_staged"),
    errors: count("error"),
    from_freezer: fromFreezer,
    skipped_pantry: [...pantrySkipped].sort((a, b) => a.localeCompare(b, "da")),
    recipes_missing_ingredients: missingIngredients,
    results,
  };
}
