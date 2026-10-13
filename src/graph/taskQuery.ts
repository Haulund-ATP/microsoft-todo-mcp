import type { CompactTask } from "./todoApi.js";

/**
 * Pure, Graph-free helpers behind list_tasks filtering/sorting,
 * multi-query search and summarize_list. Graph's To Do API can't filter on
 * creation/completion dates reliably, search case-insensitively across
 * lists, or aggregate, so these run over a compact snapshot fetched once.
 */

/** Lowercased, accents and punctuation stripped, whitespace collapsed. "Mælk!" and "mælk" match. */
export function normalizeTitle(title: string): string {
  return title
    .normalize("NFKD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

export function matchesQuery(title: string, query: string): boolean {
  const q = normalizeTitle(query);
  return q.length > 0 && normalizeTitle(title).includes(q);
}

export interface TaskFilter {
  createdAfter?: string;
  createdBefore?: string;
  completedBefore?: string;
}

const time = (iso: string | undefined) => (iso ? new Date(iso).getTime() : NaN);

export function filterTasks(tasks: CompactTask[], filter: TaskFilter): CompactTask[] {
  const after = filter.createdAfter ? time(filter.createdAfter) : undefined;
  const before = filter.createdBefore ? time(filter.createdBefore) : undefined;
  const completedBefore = filter.completedBefore ? time(filter.completedBefore) : undefined;
  return tasks.filter((t) => {
    const created = time(t.created);
    if (after !== undefined && !(created >= after)) return false;
    if (before !== undefined && !(created < before)) return false;
    if (completedBefore !== undefined && !(time(t.completed) < completedBefore)) return false;
    return true;
  });
}

export type TaskSort = "none" | "created_asc" | "created_desc" | "title" | "due";

export function sortTasks(tasks: CompactTask[], sort: TaskSort): CompactTask[] {
  if (sort === "none") return tasks;
  const copy = [...tasks];
  // Missing dates sort last in every direction.
  const byTime = (a: number, b: number, dir: 1 | -1) =>
    Number.isNaN(a) ? (Number.isNaN(b) ? 0 : 1) : Number.isNaN(b) ? -1 : (a - b) * dir;
  switch (sort) {
    case "created_asc":
      return copy.sort((a, b) => byTime(time(a.created), time(b.created), 1));
    case "created_desc":
      return copy.sort((a, b) => byTime(time(a.created), time(b.created), -1));
    case "due":
      return copy.sort((a, b) => byTime(time(a.due), time(b.due), 1));
    case "title":
      return copy.sort((a, b) => a.title.localeCompare(b.title, "da"));
  }
}

const monthOf = (iso: string | undefined) => (iso && !Number.isNaN(time(iso)) ? new Date(iso).toISOString().slice(0, 7) : "unknown");

function countBy<T>(items: T[], key: (item: T) => string): Record<string, number> {
  const out: Record<string, number> = {};
  for (const item of items) out[key(item)] = (out[key(item)] ?? 0) + 1;
  return Object.fromEntries(Object.entries(out).sort(([a], [b]) => a.localeCompare(b)));
}

export interface DuplicateGroup {
  normalized_title: string;
  count: number;
  open: number;
  tasks: Array<Pick<CompactTask, "id" | "title" | "status" | "created">>;
}

export function findDuplicates(tasks: CompactTask[], limit = 50): { groups: DuplicateGroup[]; total_groups: number } {
  const groups = new Map<string, CompactTask[]>();
  for (const t of tasks) {
    const key = normalizeTitle(t.title);
    if (!key) continue;
    groups.set(key, [...(groups.get(key) ?? []), t]);
  }
  const dupes = [...groups.entries()]
    .filter(([, g]) => g.length > 1)
    .map(([key, g]) => ({
      normalized_title: key,
      count: g.length,
      open: g.filter((t) => t.status !== "completed").length,
      tasks: g.map(({ id, title, status, created }) => ({ id, title, status, created })),
    }))
    .sort((a, b) => b.count - a.count || a.normalized_title.localeCompare(b.normalized_title));
  return { groups: dupes.slice(0, limit), total_groups: dupes.length };
}

export function summarizeTasks(tasks: CompactTask[], now = Date.now(), sampleSize = 10) {
  const open = tasks.filter((t) => t.status !== "completed");
  const completed = tasks.filter((t) => t.status === "completed");
  const overdue = open.filter((t) => time(t.due) < now);
  const oldestOpen = [...open]
    .sort((a, b) => time(a.created) - time(b.created))
    .slice(0, sampleSize)
    .map(({ id, title, created, due }) => ({ id, title, created, due }));
  const lastCompleted = completed.map((t) => time(t.completed)).filter((n) => !Number.isNaN(n));
  return {
    total: tasks.length,
    by_status: countBy(tasks, (t) => t.status),
    by_importance: countBy(tasks, (t) => t.importance),
    open: open.length,
    completed: completed.length,
    overdue: overdue.length,
    recurring: tasks.filter((t) => t.recurring).length,
    created_by_month: countBy(tasks, (t) => monthOf(t.created)),
    completed_by_month: countBy(completed, (t) => monthOf(t.completed)),
    oldest_open: oldestOpen,
    most_recent_completion: lastCompleted.length ? new Date(Math.max(...lastCompleted)).toISOString() : undefined,
    duplicates: findDuplicates(tasks),
  };
}
