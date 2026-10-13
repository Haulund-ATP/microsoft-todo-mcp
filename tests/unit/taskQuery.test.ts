import { describe, expect, it } from "vitest";
import { filterTasks, findDuplicates, matchesQuery, normalizeTitle, sortTasks, summarizeTasks } from "../../src/graph/taskQuery.js";
import type { CompactTask } from "../../src/graph/todoApi.js";

const task = (id: string, title: string, extra: Partial<CompactTask> = {}): CompactTask => ({
  id, title, status: "notStarted", importance: "normal", ...extra,
});

describe("title matching", () => {
  it("ignores case, accents, punctuation and extra whitespace but keeps Danish letters", () => {
    expect(normalizeTitle("  Mælk,   ØL & Café! ")).toBe("mælk øl cafe");
    expect(matchesQuery("Køb MÆLK i Netto", "mælk")).toBe(true);
    expect(matchesQuery("Køb mælk", "smør")).toBe(false);
    expect(matchesQuery("anything", "  !! ")).toBe(false);
  });
});

describe("filterTasks / sortTasks", () => {
  const tasks = [
    task("a", "B", { created: "2025-01-10T00:00:00Z", status: "completed", completed: "2025-02-01T00:00:00Z" }),
    task("b", "a", { created: "2026-06-01T00:00:00Z", due: "2026-07-01T00:00:00Z" }),
    task("c", "C"),
  ];

  it("filters by created range and completion date, excluding tasks missing the date", () => {
    expect(filterTasks(tasks, { createdAfter: "2026-01-01" }).map((t) => t.id)).toEqual(["b"]);
    expect(filterTasks(tasks, { createdBefore: "2026-01-01" }).map((t) => t.id)).toEqual(["a"]);
    expect(filterTasks(tasks, { completedBefore: "2025-06-01" }).map((t) => t.id)).toEqual(["a"]);
    expect(filterTasks(tasks, {})).toHaveLength(3);
  });

  it("sorts by date with missing dates last, and by title", () => {
    expect(sortTasks(tasks, "created_asc").map((t) => t.id)).toEqual(["a", "b", "c"]);
    expect(sortTasks(tasks, "created_desc").map((t) => t.id)).toEqual(["b", "a", "c"]);
    expect(sortTasks(tasks, "due").map((t) => t.id)).toEqual(["b", "a", "c"]);
    expect(sortTasks(tasks, "title").map((t) => t.title)).toEqual(["a", "B", "C"]);
    expect(sortTasks(tasks, "none")).toBe(tasks);
  });
});

describe("summarizeTasks", () => {
  const now = Date.parse("2026-10-13T00:00:00Z");
  const tasks = [
    task("1", "Mælk", { created: "2025-01-05T00:00:00Z" }),
    task("2", "mælk!", { created: "2025-03-01T00:00:00Z", status: "completed", completed: "2025-03-02T00:00:00Z" }),
    task("3", "Støvsug", { created: "2026-10-01T00:00:00Z", due: "2026-10-01T00:00:00Z", recurring: true, importance: "high" }),
    task("4", "Mælk", { created: "2026-10-02T00:00:00Z" }),
  ];

  it("counts, buckets by month and finds duplicates and overdue tasks", () => {
    const summary = summarizeTasks(tasks, now);
    expect(summary).toMatchObject({
      total: 4, open: 3, completed: 1, overdue: 1, recurring: 1,
      by_status: { completed: 1, notStarted: 3 },
      by_importance: { high: 1, normal: 3 },
      created_by_month: { "2025-01": 1, "2025-03": 1, "2026-10": 2 },
      completed_by_month: { "2025-03": 1 },
      most_recent_completion: "2025-03-02T00:00:00.000Z",
    });
    expect(summary.oldest_open.map((t) => t.id)).toEqual(["1", "3", "4"]);
    expect(summary.duplicates).toEqual({
      total_groups: 1,
      groups: [{ normalized_title: "mælk", count: 3, open: 2, tasks: expect.any(Array) }],
    });
  });

  it("caps duplicate groups but reports the total", () => {
    const many = Array.from({ length: 6 }, (_, i) => [task(`${i}a`, `T${i}`), task(`${i}b`, `t${i}`)]).flat();
    const result = findDuplicates(many, 2);
    expect(result.groups).toHaveLength(2);
    expect(result.total_groups).toBe(6);
  });
});
