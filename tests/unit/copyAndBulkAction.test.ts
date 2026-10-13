import { beforeEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import { buildCopyPayload, copyTasks } from "../../src/graph/copyTasks.js";
import { bulkTaskAction } from "../../src/graph/bulkTaskAction.js";
import { searchTasksMulti } from "../../src/graph/search.js";
import { getTasksBatch } from "../../src/graph/getTasks.js";
import * as todo from "../../src/graph/todoApi.js";
import { bulkTaskActionInputShape, copyTasksInputShape } from "../../src/tools/bulkTools.js";
import { getTasksInputShape, listTasksInputShape, searchTasksInputShape } from "../../src/tools/readTools.js";

type Task = Record<string, unknown> & { id: string; title: string; status?: string };
type Item = { id: string; displayName: string; isChecked: boolean };

/** Stateful fake of the To Do endpoints these modules use, one account per instance. */
function fakeAccount(lists: Record<string, { name?: string; tasks: Task[] }>, opts: { rejectFields?: string[] } = {}) {
  const checklists = new Map<string, Item[]>();
  const writes: Array<{ method: string; path: string; body?: unknown }> = [];
  let seq = 0;
  const notFound = () => Object.assign(new Error("not found"), { statusCode: 404 });
  const client = {
    api(path: string) {
      const m = /^\/me\/todo\/lists(?:\/([^/]+)(?:\/tasks(?:\/([^/]+)(\/checklistItems)?)?)?)?$/.exec(path.split("?")[0]!)!;
      const [, listId, taskId, checklist] = m;
      const list = listId ? lists[decodeURIComponent(listId)] : undefined;
      const find = () => list?.tasks.find((t) => t.id === decodeURIComponent(taskId!));
      const chain = {
        top: () => chain, select: () => chain, filter: () => chain,
        async get() {
          if (!listId) return { value: Object.entries(lists).map(([id, l]) => ({ id, displayName: l.name ?? id })) };
          if (!list) throw notFound();
          if (!taskId) return { value: list.tasks } ;
          const t = find();
          if (!t) throw notFound();
          if (checklist) return { value: checklists.get(t.id) ?? [] };
          return t;
        },
        async post(body: Record<string, unknown>) {
          writes.push({ method: "post", path, body });
          if (checklist) {
            const item = { id: `i${++seq}`, displayName: String(body.displayName), isChecked: Boolean(body.isChecked) };
            checklists.set(find()!.id, [...(checklists.get(find()!.id) ?? []), item]);
            return item;
          }
          if (opts.rejectFields?.some((f) => f in body)) throw Object.assign(new Error("bad"), { statusCode: 400 });
          const created = { ...body, id: `n${++seq}` } as Task;
          list!.tasks.push(created);
          return created;
        },
        async patch(body: Record<string, unknown>) {
          writes.push({ method: "patch", path, body });
          Object.assign(find()!, body);
        },
        async delete() {
          writes.push({ method: "delete", path });
          list!.tasks = list!.tasks.filter((t) => t.id !== decodeURIComponent(taskId!));
        },
      };
      return chain;
    },
  } as unknown as Client;
  return { client, lists, checklists, writes };
}

beforeEach(() => todo._resetTaskPagingForTests());

describe("copyTasks", () => {
  const source = () => {
    const acc = fakeAccount({
      S: { tasks: [
        { id: "a", title: "Mælk", status: "completed", importance: "high", body: { content: "<b>2 L</b>", contentType: "html" },
          recurrence: { pattern: { type: "weekly" } }, completedDateTime: { dateTime: "2026-01-01T00:00:00", timeZone: "UTC" },
          createdDateTime: "2025-01-01T00:00:00Z", "@odata.etag": "x" },
        { id: "b", title: "Brød", status: "notStarted" },
      ] },
    });
    acc.checklists.set("a", [{ id: "c1", displayName: "Let", isChecked: true }, { id: "c2", displayName: "Sød", isChecked: false }]);
    return acc;
  };

  it("previews by default without writing to either account", async () => {
    const src = source();
    const dst = fakeAccount({ T: { tasks: [] } });
    const result = await copyTasks(src.client, "S", dst.client, "T", ["a", "b"], { dryRun: true, includeChecklist: true, allowDuplicates: false });
    expect(result.results.map((r) => r.status)).toEqual(["preview", "preview"]);
    expect(result.results[0].checklist_items).toBe(2);
    expect([...src.writes, ...dst.writes]).toEqual([]);
  });

  it("copies notes, status, recurrence and checked checklist items across accounts, skipping server-owned fields", async () => {
    const src = source();
    const dst = fakeAccount({ T: { tasks: [] } });
    const result = await copyTasks(src.client, "S", dst.client, "T", ["a"], { dryRun: false, includeChecklist: true, allowDuplicates: false });
    expect(result.copied).toBe(1);
    const created = dst.lists.T.tasks[0]!;
    expect(created).toMatchObject({ title: "Mælk", status: "completed", importance: "high", body: { content: "<b>2 L</b>" }, recurrence: { pattern: { type: "weekly" } } });
    expect(created).not.toHaveProperty("createdDateTime");
    expect(created).not.toHaveProperty("@odata.etag");
    expect(dst.checklists.get(created.id)).toEqual([
      expect.objectContaining({ displayName: "Let", isChecked: true }),
      expect.objectContaining({ displayName: "Sød", isChecked: false }),
    ]);
    expect(src.lists.S.tasks).toHaveLength(2); // copy leaves the source alone
  });

  it("never writes to the source, and skips titles already in the target", async () => {
    const src = source();
    const dst = fakeAccount({ T: { tasks: [{ id: "x", title: "brød" }] } });
    const result = await copyTasks(src.client, "S", dst.client, "T", ["a", "b"], { dryRun: false, includeChecklist: true, allowDuplicates: false });
    expect(result.results.map((r) => r.status)).toEqual(["copied", "exists"]);
    expect(src.writes).toEqual([]);
    expect(src.lists.S.tasks.map((t) => t.id)).toEqual(["a", "b"]);
  });

  it("falls back to essential fields when Graph rejects the full payload", async () => {
    const src = source();
    const dst = fakeAccount({ T: { tasks: [] } }, { rejectFields: ["recurrence"] });
    const result = await copyTasks(src.client, "S", dst.client, "T", ["a"], { dryRun: false, includeChecklist: false, allowDuplicates: false });
    expect(result.results[0]).toMatchObject({ status: "copied", warning: expect.stringContaining("only title") });
    expect(dst.lists.T.tasks[0]).not.toHaveProperty("recurrence");
  });

  it("reports a failed create without exposing the upstream error", async () => {
    const src = fakeAccount({ S: { tasks: [{ id: "b", title: "Brød" }] } });
    const dst = fakeAccount({ T: { tasks: [] } });
    vi.spyOn(dst.client, "api").mockImplementation(() => {
      const chain = { top: () => chain, select: () => chain, filter: () => chain, get: async () => ({ value: [] }), post: async () => { throw new Error("boom"); } };
      return chain as never;
    });
    const result = await copyTasks(src.client, "S", dst.client, "T", ["b"], { dryRun: false, includeChecklist: false, allowDuplicates: false });
    expect(result.results[0]).toMatchObject({ status: "error", warning: expect.stringContaining("check the target list") });
    expect(JSON.stringify(result)).not.toContain("boom");
  });

  it("only copies completedDateTime for completed tasks and drops empty notes", () => {
    expect(buildCopyPayload({ title: "t", status: "notStarted", completedDateTime: { dateTime: "x" }, body: { content: "" }, categories: [] }))
      .toEqual({ title: "t", status: "notStarted" });
  });
});

describe("bulkTaskAction", () => {
  const account = () => fakeAccount({ L: { tasks: [
    { id: "1", title: "Old", status: "completed" },
    { id: "2", title: "Open", status: "notStarted" },
    { id: "3", title: "Renamed elsewhere", status: "completed" },
  ] } });
  const items = [
    { task_id: "1", expected_title: "Old" },
    { task_id: "2", expected_title: "Open" },
    { task_id: "3", expected_title: "Stale title" },
    { task_id: "gone", expected_title: "x" },
  ];

  it("previews deletes, protects open tasks and stale titles by default", async () => {
    const acc = account();
    const result = await bulkTaskAction(acc.client, "L", "delete", items, { dryRun: true, onlyCompleted: true });
    expect(result.results.map((r) => r.status)).toEqual(["preview", "skipped_not_completed", "conflict", "unchanged"]);
    expect(acc.writes).toEqual([]);
  });

  it("deletes and verifies only the eligible tasks", async () => {
    const acc = account();
    const result = await bulkTaskAction(acc.client, "L", "delete", items, { dryRun: false, onlyCompleted: true });
    expect(result.verified).toBe(1);
    expect(acc.lists.L.tasks.map((t) => t.id)).toEqual(["2", "3"]);
  });

  it("completes open tasks and treats already-completed ones as unchanged", async () => {
    const acc = account();
    const result = await bulkTaskAction(acc.client, "L", "complete", items.slice(0, 2), { dryRun: false, onlyCompleted: true });
    expect(result.results.map((r) => r.status)).toEqual(["unchanged", "verified"]);
    expect(acc.lists.L.tasks[1]!.status).toBe("completed");
  });
});

describe("searchTasksMulti", () => {
  it("matches several queries across every list with one read per list", async () => {
    const acc = fakeAccount({
      A: { name: "Indkøb", tasks: [{ id: "1", title: "Køb mælk" }, { id: "2", title: "Brød" }] },
      B: { name: "Hjem", tasks: [{ id: "3", title: "MÆLKEBØTTE fjernes" }] },
    });
    const api = vi.spyOn(acc.client, "api");
    const result = await searchTasksMulti(acc.client, ["mælk", "brød", "ost"], { includeCompleted: true, maxPerQuery: 1 });
    expect(result.lists_scanned).toBe(2);
    expect(result.results.map((r) => [r.query, r.match_count, r.matches.length])).toEqual([["mælk", 2, 1], ["brød", 1, 1], ["ost", 0, 0]]);
    expect(result.results[0]!.matches[0]).toMatchObject({ id: "1", list_name: "Indkøb" });
    expect(api.mock.calls.filter(([p]) => String(p).endsWith("/tasks"))).toHaveLength(2);
  });
});

describe("listTasksWithMeta paging fallback", () => {
  it("retries with the old page shape once Graph rejects $top=100/$select, and remembers it", async () => {
    const tops: number[] = [];
    const client = {
      api: () => {
        let top = 0;
        const chain = {
          top: (n: number) => { top = n; tops.push(n); return chain; },
          select: () => chain, filter: () => chain,
          get: async () => {
            if (top === 100) throw Object.assign(new Error("bad"), { statusCode: 400 });
            return { value: [{ id: "1", title: "t" }] };
          },
        };
        return chain;
      },
    } as unknown as Client;
    expect((await todo.listTasksWithMeta(client, "L", { compact: true })).tasks).toHaveLength(1);
    await todo.listTasksWithMeta(client, "L", { compact: true });
    expect(tops).toEqual([100, 50, 50]);
  });

  it("keeps the larger page size when only $select is rejected", async () => {
    const shapes: string[] = [];
    const client = {
      api: () => {
        let top = 0;
        let selected = false;
        const chain = {
          top: (n: number) => { top = n; return chain; },
          select: () => { selected = true; return chain; },
          filter: () => chain,
          get: async () => {
            shapes.push(`${top}${selected ? "+select" : ""}`);
            if (selected) throw Object.assign(new Error("bad"), { statusCode: 400, code: "BadRequest" });
            return { value: [{ id: "1", title: "t" }] };
          },
        };
        return chain;
      },
    } as unknown as Client;
    await todo.listTasksWithMeta(client, "L", { compact: true });
    await todo.listTasksWithMeta(client, "L", { compact: true });
    expect(shapes).toEqual(["100+select", "50+select", "100", "100"]);
  });

  it("does not disable the optimization when the request fails for another reason", async () => {
    const tops: number[] = [];
    const client = {
      api: () => {
        const chain = {
          top: (n: number) => { tops.push(n); return chain; }, select: () => chain, filter: () => chain,
          get: async () => { throw Object.assign(new Error("bad filter"), { statusCode: 400 }); },
        };
        return chain;
      },
    } as unknown as Client;
    await expect(todo.listTasksWithMeta(client, "L", { filter: "bad" })).rejects.toThrow("bad filter");
    await expect(todo.listTasksWithMeta(client, "L", { filter: "bad" })).rejects.toThrow();
    expect(tops).toEqual([100, 50, 100, 50]);
  });
});

describe("tool input schemas", () => {
  it("defaults bulk writes to preview and rejects duplicate ids and oversized batches", () => {
    const copy = z.object(copyTasksInputShape);
    expect(copy.parse({ list_id: "S", task_ids: ["a"], target_list_id: "T" })).toMatchObject({ dry_run: true, allow_duplicates: false });
    expect(copy.safeParse({ list_id: "S", task_ids: ["a", "a"], target_list_id: "T" }).success).toBe(false);
    const bulk = z.object(bulkTaskActionInputShape);
    expect(bulk.parse({ list_id: "L", action: "delete", items: [{ task_id: "1", expected_title: "x" }] })).toMatchObject({ dry_run: true, only_completed: true });
    expect(bulk.safeParse({ list_id: "L", action: "delete", items: Array.from({ length: 201 }, (_, i) => ({ task_id: `${i}`, expected_title: "x" })) }).success).toBe(false);
  });

  it("keeps list_tasks and search_tasks backward compatible", () => {
    expect(z.object(listTasksInputShape).parse({ list_id: "L" })).toMatchObject({ max_items: 50, status: "all", sort: "none" });
    expect(z.object(searchTasksInputShape).parse({ list_id: "L", query: "mælk" })).toMatchObject({ include_completed: true, max_items: 50 });
  });
});

describe("tool annotations", () => {
  it("marks every tool as closed-world and copy_tasks as non-destructive", async () => {
    const { registerAllTools } = await import("../../src/tools/registerAll.js");
    const registerTool = vi.fn();
    registerAllTools({ registerTool } as never, { profileAlias: "nadia", correlationId: "t" });
    const annotations = Object.fromEntries(registerTool.mock.calls.map(([name, def]) => [name, def.annotations]));
    for (const [name, a] of Object.entries(annotations)) expect(a, name).toMatchObject({ openWorldHint: false });
    expect(annotations.copy_tasks).toMatchObject({ readOnlyHint: false, destructiveHint: false });
    expect(annotations.bulk_task_action).toMatchObject({ destructiveHint: true });
  });
});
describe("getTasksBatch", () => {
  it("returns full tasks in input order, with checklist on request, and reports missing ones", async () => {
    const acc = fakeAccount({ L: { tasks: [
      { id: "1", title: "Skammel", status: "completed", body: { content: "IKEA", contentType: "text" } },
      { id: "2", title: "Lampe" },
    ] } });
    acc.checklists.set("1", [{ id: "c", displayName: "Skruer", isChecked: true }]);
    const result = await getTasksBatch(acc.client, "L", ["2", "gone", "1"], true);
    expect(result).toMatchObject({ requested: 3, found: 2, not_found: 1, errors: 0 });
    expect(result.results.map((r) => r.status)).toEqual(["found", "not_found", "found"]);
    expect(result.results[2]).toMatchObject({ task: { title: "Skammel", body: { content: "IKEA" } }, checklist: [{ displayName: "Skruer" }] });
    expect(result.results[0]).toMatchObject({ checklist: [] });
  });

  it("skips checklist reads unless asked, and hides upstream error text", async () => {
    const acc = fakeAccount({ L: { tasks: [{ id: "1", title: "t" }] } });
    const api = vi.spyOn(acc.client, "api");
    const ok = await getTasksBatch(acc.client, "L", ["1"], false);
    expect(ok.results[0]).not.toHaveProperty("checklist", expect.anything());
    expect(api.mock.calls.some(([p]) => String(p).endsWith("/checklistItems"))).toBe(false);

    api.mockImplementation(() => ({ get: async () => { throw Object.assign(new Error("secret upstream"), { statusCode: 500 }); } }) as never);
    const failed = await getTasksBatch(acc.client, "L", ["1"], false);
    expect(failed.results[0]).toMatchObject({ status: "error" });
    expect(JSON.stringify(failed)).not.toContain("secret upstream");
  });

  it("caps batches at 100 unique ids", () => {
    const schema = z.object(getTasksInputShape);
    expect(schema.parse({ list_id: "L", task_ids: ["a"] }).include_checklist).toBe(false);
    expect(schema.safeParse({ list_id: "L", task_ids: ["a", "a"] }).success).toBe(false);
    expect(schema.safeParse({ list_id: "L", task_ids: Array.from({ length: 101 }, (_, i) => `${i}`) }).success).toBe(false);
  });
});