import { beforeEach, describe, expect, it } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import { buildShoppingCandidates, itemKey, matchesItem, noteText, parseRecipeNote } from "../../src/graph/shoppingCandidates.js";
import * as todo from "../../src/graph/todoApi.js";

type Task = Record<string, unknown> & { id: string; title: string; status?: string };

/** Fake of the list/create endpoints the planner uses. */
function fakeAccount(lists: Record<string, Task[]>) {
  const writes: Array<{ path: string; body: Record<string, unknown> }> = [];
  let seq = 0;
  const client = {
    api(path: string) {
      const listId = decodeURIComponent(/^\/me\/todo\/lists\/([^/]+)\/tasks$/.exec(path)![1]!);
      const chain = {
        top: () => chain, select: () => chain, filter: () => chain,
        async get() { return { value: lists[listId] }; },
        async post(body: Record<string, unknown>) {
          writes.push({ path, body });
          const created = { ...body, id: `n${++seq}` } as Task;
          lists[listId]!.push(created);
          return created;
        },
      };
      return chain;
    },
  } as unknown as Client;
  return { client, lists, writes };
}

const recipe = (id: string, title: string, note: Record<string, unknown>): Task => ({
  id, title, status: "notStarted", body: { content: JSON.stringify(note), contentType: "text" },
});

beforeEach(() => todo._resetTaskPagingForTests());

describe("item matching", () => {
  it("reduces free-text shopping titles to the item", () => {
    expect(itemKey("Flåede tomater – 2 dåser")).toBe(itemKey("flåede tomat"));
    expect(itemKey("Remoulade - kig i kælder depot")).toBe("remoulad");
    expect(itemKey("Jordbær 400 g")).toBe("jordbær");
    expect(itemKey("Æbler")).toBe(itemKey("æble"));
    expect(itemKey("Hakket oksekød 8-12%")).toBe("hakket oksekød");
  });

  it("lets a more specific title stand in for an item, never a more general one", () => {
    expect(matchesItem(["Champignon"], "Champignon, brune danske")).toBe(true);
    expect(matchesItem(["Champignon"], "Brune champignon")).toBe(true);
    expect(matchesItem(["Hakkede tomater"], "Tomater")).toBe(false);
    expect(matchesItem(["Piskefløde", "Madlavningsfløde"], "Madlavningsfløde 9%")).toBe(true);
    expect(matchesItem(["Mælk"], "Kokosmælk")).toBe(false);
    expect(matchesItem(["Rød paprika"], "Paprika", { exact: true })).toBe(false);
    expect(matchesItem(["Salt"], "Salt", { exact: true })).toBe(true);
  });
});

describe("recipe notes", () => {
  it("reads JSON from text and from HTML notes", () => {
    const json = { dates_planned: ["2026-10-17"], ingredienser: [{ vare: ["Æg"], mængde: "6" }] };
    expect(parseRecipeNote({ content: `${JSON.stringify(json)}\r\n`, contentType: "text" })).toEqual(json);
    const html = JSON.stringify(json).replace(/"/g, "&quot;");
    expect(parseRecipeNote({ content: `<p>${html}</p>`, contentType: "html" })).toEqual(json);
    expect(noteText({ content: "a<br>b &amp; c", contentType: "html" })).toBe("a\nb & c");
    expect(parseRecipeNote({ content: "side 112", contentType: "text" })).toBeUndefined();
  });
});

describe("buildShoppingCandidates", () => {
  const account = () =>
    fakeAccount({
      R: [
        recipe("r1", "Æggekage", {
          dates_planned: ["2026-10-17"],
          kød: [{ vare: ["Bacon i tern", "Bacon i skiver"], mængde: "1 pakke" }],
          ingredienser: [
            { vare: ["Æg"], mængde: "6 stk" },
            { vare: ["Rugbrød"], mængde: "1" },
            { vare: ["Salt"] },
            { vare: ["Tomater"], mængde: "3" },
          ],
        }),
        recipe("r2", "Burger", {
          dates_planned: ["2026-10-13", "2026-10-30"],
          kød: [{ vare: ["Hakket oksekød"], mængde: "500 g" }],
          ingredienser: [{ vare: ["Tomater"], mængde: "2" }, { vare: ["Burgerboller"], mængde: "4" }],
        }),
        recipe("r3", "Pariserbøf", { dates_planned: ["2026-10-14"], kød: [{ vare: ["Hakket oksekød"], mængde: "400 g" }] }),
        recipe("r4", "Rester", { dates_planned: ["2026-10-15"], kategori: "rester" }),
        recipe("r5", "Ikke planlagt", { dates_planned: ["2026-11-01"], ingredienser: [{ vare: ["Kaviar"] }] }),
      ],
      S: [
        { id: "s1", title: "Rugbrød", status: "notStarted" },
        { id: "s2", title: "Tomater", status: "completed", completedDateTime: { dateTime: "2026-09-01T10:00:00", timeZone: "UTC" } },
        { id: "s3", title: "Tomater – 6 stk", status: "completed", completedDateTime: { dateTime: "2026-10-01T10:00:00", timeZone: "UTC" } },
        { id: "s4", title: "Æg julefrokost", status: "completed" },
      ],
      A: [{ id: "a1", title: "Burgerboller", status: "notStarted" }],
      P: [{ id: "p1", title: "Salt", status: "notStarted" }],
      F: [{ id: "f1", title: "Skuffe 2 – Hakket oksekød – 500 g – pakke 1/2", status: "notStarted" }],
    });
  const options = { recipeListId: "R", from: "2026-10-13", to: "2026-10-19", stagingListId: "A", shoppingListId: "S", pantryListId: "P", freezerListId: "F" };

  it("previews without writing and sorts every need into the right bucket", async () => {
    const acc = account();
    const result = await buildShoppingCandidates(acc.client, { ...options, dryRun: true });
    expect(acc.writes).toEqual([]);
    expect(result.meals.map((m) => m.recipe)).toEqual(["Burger", "Pariserbøf", "Rester", "Æggekage"]);
    expect(result.from_freezer.map((f) => f.recipe)).toEqual(["Burger", "Pariserbøf"]);
    expect(result.skipped_pantry).toEqual(["Salt"]);
    expect(result.recipes_missing_ingredients.map((m) => m.recipe)).toEqual(["Pariserbøf"]);
    const byVare = Object.fromEntries(result.results.map((r) => [r.vare, r]));
    expect(byVare["Rugbrød"]!.status).toBe("already_on_shopping");
    expect(byVare["Burgerboller"]!.status).toBe("already_staged");
    expect(byVare["Bacon i tern"]!.status).toBe("preview");
    expect(byVare["Bacon i tern"]!.reuse_shopping_task).toBeUndefined();
    // Most recently completed match wins; "Æg julefrokost" is "Æg" plus one word.
    expect(byVare["Tomater"]!.reuse_shopping_task).toEqual({ id: "s3", title: "Tomater – 6 stk" });
    expect(byVare["Tomater"]!.uses).toEqual(["2 – Burger (tir 13/10)", "3 – Æggekage (lør 17/10)"]);
    expect(byVare["Æg"]!.reuse_shopping_task?.id).toBe("s4");
  });

  it("writes one staging task per item with the shopping id Home Assistant reopens", async () => {
    const acc = account();
    const result = await buildShoppingCandidates(acc.client, { ...options, dryRun: false });
    expect(result.staged).toBe(3);
    const created = acc.lists.A!.slice(1);
    expect(created.map((t) => t.title).sort()).toEqual(["Bacon i tern", "Tomater", "Æg"]);
    const tomater = created.find((t) => t.title === "Tomater")!;
    const note = (tomater.body as { content: string }).content;
    expect(note).toContain("shopping_id: s3");
    expect(note).toContain("2 – Burger (tir 13/10)");
    expect((created.find((t) => t.title === "Bacon i tern")!.body as { content: string }).content).toContain("Ny vare på Shopping");

    const rerun = await buildShoppingCandidates(acc.client, { ...options, dryRun: false });
    expect(rerun.staged).toBe(0);
    expect(rerun.already_staged).toBe(4);
  });

  it("rejects a window that ends before it starts", async () => {
    await expect(buildShoppingCandidates(account().client, { ...options, from: "2026-10-20", dryRun: true })).rejects.toThrow();
  });
});
