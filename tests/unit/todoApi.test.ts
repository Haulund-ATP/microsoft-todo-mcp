import { describe, it, expect, vi } from "vitest";
import * as todo from "../../src/graph/todoApi.js";
import type { Client } from "@microsoft/microsoft-graph-client";

/**
 * These tests lock in the finding from the shared-list checklist-item
 * data-loss investigation: every write this module sends to Graph must be
 * field-scoped to exactly the resource it targets — a task update must
 * never touch checklistItems (they are a separate navigation property per
 * https://learn.microsoft.com/en-us/graph/api/resources/todotask), and a
 * checklist-item write must never touch the parent task's body/title.
 */

interface RecordedCall {
  path: string;
  method: "get" | "post" | "patch" | "delete";
  body?: unknown;
}

function makeFakeClient(responses: Record<string, unknown> = {}) {
  const calls: RecordedCall[] = [];

  function apiFor(path: string) {
    const record = (method: RecordedCall["method"], body?: unknown) => {
      calls.push({ path, method, body });
      const response = responses[path];
      if (response instanceof Error) return Promise.reject(response);
      return Promise.resolve(response ?? { id: "generated-id" });
    };
    const chain = {
      top: () => chain,
      filter: () => chain,
      get: () => record("get"),
      post: (body?: unknown) => record("post", body),
      patch: (body?: unknown) => record("patch", body),
      delete: () => record("delete"),
    };
    return chain;
  }

  const client = { api: vi.fn(apiFor) } as unknown as Client;
  return { client, calls };
}

describe("todoApi request shape — field-scoped writes only", () => {
  it("updateTask only PATCHes the task endpoint, never a checklistItems path", async () => {
    const { client, calls } = makeFakeClient({
      "/me/todo/lists/L/tasks/T": { id: "T", title: "renamed" },
    });
    await todo.updateTask(client, "L", "T", { title: "renamed" });

    const patchCalls = calls.filter((c) => c.method === "patch");
    expect(patchCalls).toHaveLength(1);
    expect(patchCalls[0].path).toBe("/me/todo/lists/L/tasks/T");
    expect(patchCalls[0].path).not.toContain("checklistItems");
  });

  it("updateTask's PATCH body contains only the explicitly provided fields", async () => {
    const { client, calls } = makeFakeClient({ "/me/todo/lists/L/tasks/T": { id: "T" } });
    await todo.updateTask(client, "L", "T", { title: "only title changed" });

    const patchCall = calls.find((c) => c.method === "patch")!;
    expect(patchCall.body).toEqual({ title: "only title changed" });
    // Must not silently include body, importance, status, checklistItems, etc.
    expect(Object.keys(patchCall.body as object)).toEqual(["title"]);
  });

  it("completeTask/reopenTask only send a status field, never body or checklistItems", async () => {
    const { client, calls } = makeFakeClient({ "/me/todo/lists/L/tasks/T": { id: "T" } });
    await todo.completeTask(client, "L", "T");

    const patchCall = calls.find((c) => c.method === "patch")!;
    expect(patchCall.body).toEqual({ status: "completed" });
  });

  it("addChecklistItem only POSTs to the checklistItems sub-resource, never the task itself", async () => {
    const { client, calls } = makeFakeClient({
      "/me/todo/lists/L/tasks/T/checklistItems": { id: "item-1", displayName: "A", isChecked: false },
    });
    await todo.addChecklistItem(client, "L", "T", "A");

    const postCalls = calls.filter((c) => c.method === "post");
    expect(postCalls).toHaveLength(1);
    expect(postCalls[0].path).toBe("/me/todo/lists/L/tasks/T/checklistItems");
    expect(postCalls[0].body).toEqual({ displayName: "A" });
  });

  it("updateChecklistItem only PATCHes the specific item and never the task or other items", async () => {
    const { client, calls } = makeFakeClient({
      "/me/todo/lists/L/tasks/T/checklistItems/item-1": { id: "item-1", displayName: "A2", isChecked: true },
    });
    await todo.updateChecklistItem(client, "L", "T", "item-1", { isChecked: true });

    const patchCall = calls.find((c) => c.method === "patch")!;
    expect(patchCall.path).toBe("/me/todo/lists/L/tasks/T/checklistItems/item-1");
    // Only the provided field should be present — an undefined displayName
    // must not be sent as an explicit overwrite.
    expect(patchCall.body).toEqual({ isChecked: true });
  });

  it("deleteChecklistItem only DELETEs the specific item path", async () => {
    const { client, calls } = makeFakeClient({});
    await todo.deleteChecklistItem(client, "L", "T", "item-1");

    const deleteCall = calls.find((c) => c.method === "delete")!;
    expect(deleteCall.path).toBe("/me/todo/lists/L/tasks/T/checklistItems/item-1");
  });

  it("listChecklistItems maps id, displayName, isChecked, createdDateTime, and checkedDateTime", async () => {
    const { client } = makeFakeClient({
      "/me/todo/lists/L/tasks/T/checklistItems": {
        value: [
          { id: "item-1", displayName: "A", isChecked: true, createdDateTime: "2026-01-01T00:00:00Z", checkedDateTime: "2026-01-02T00:00:00Z" },
        ],
      },
    });
    const items = await todo.listChecklistItems(client, "L", "T");
    expect(items).toEqual([
      {
        id: "item-1",
        displayName: "A",
        isChecked: true,
        createdDateTime: "2026-01-01T00:00:00Z",
        checkedDateTime: "2026-01-02T00:00:00Z",
      },
    ]);
  });

  it("createTask omits body/dueDateTime/reminderDateTime entirely when not provided", async () => {
    const { client, calls } = makeFakeClient({ "/me/todo/lists/L/tasks": { id: "T" } });
    await todo.createTask(client, "L", { title: "bare task" });

    const postCall = calls.find((c) => c.method === "post")!;
    expect(postCall.body).toEqual({ title: "bare task", importance: "normal" });
  });
});

/**
 * A minimal stateful fake of the checklistItems sub-resource, so these
 * tests exercise actual add/list/update sequences (per the reported bug's
 * reproduction steps) rather than only asserting single-call shapes.
 * Deliberately does NOT model Microsoft's shared-list sync backend — that
 * external behavior is exactly what's in question and cannot be faithfully
 * simulated in-process. See docs/security.md for the live, two-account
 * manual procedure that covers cross-participant sync (reproduction
 * scenarios 2 and 3 from the bug report).
 */
function makeStatefulFakeClient(taskBody?: { content: string; contentType: "text" }) {
  const items = new Map<string, { id: string; displayName: string; isChecked: boolean }>();
  let nextId = 1;
  let task = { id: "T", title: "task", body: taskBody };

  function apiFor(path: string) {
    const itemMatch = /^\/me\/todo\/lists\/L\/tasks\/T\/checklistItems\/(.+)$/.exec(path);
    const chain = {
      top: () => chain,
      filter: () => chain,
      get: () => {
        if (path === "/me/todo/lists/L/tasks/T") return Promise.resolve(task);
        if (path === "/me/todo/lists/L/tasks/T/checklistItems") {
          return Promise.resolve({ value: [...items.values()] });
        }
        if (itemMatch) return Promise.resolve(items.get(itemMatch[1]));
        throw new Error(`unexpected GET ${path}`);
      },
      post: (body: { displayName: string }) => {
        const id = `item-${nextId++}`;
        const item = { id, displayName: body.displayName, isChecked: false };
        items.set(id, item);
        return Promise.resolve(item);
      },
      patch: (body: Record<string, unknown>) => {
        if (path === "/me/todo/lists/L/tasks/T") {
          task = { ...task, ...body } as typeof task;
          return Promise.resolve(task);
        }
        if (itemMatch) {
          const existing = items.get(itemMatch[1])!;
          items.set(itemMatch[1], { ...existing, ...body });
          return Promise.resolve(items.get(itemMatch[1]));
        }
        throw new Error(`unexpected PATCH ${path}`);
      },
      delete: () => {
        if (itemMatch) items.delete(itemMatch[1]);
        return Promise.resolve();
      },
    };
    return chain;
  }

  return { client: { api: vi.fn(apiFor) } as unknown as Client, getTaskBody: () => task.body };
}

describe("todoApi stateful reproduction — bug report scenarios 1, 4, 5", () => {
  it("scenario 1: multiple MCP-created checklist items all persist and are independently updatable", async () => {
    const { client } = makeStatefulFakeClient();
    const a = await todo.addChecklistItem(client, "L", "T", "Testpunkt 1");
    const b = await todo.addChecklistItem(client, "L", "T", "Testpunkt 2");

    let items = await todo.listChecklistItems(client, "L", "T");
    expect(items.map((i) => i.displayName).sort()).toEqual(["Testpunkt 1", "Testpunkt 2"]);

    await todo.updateChecklistItem(client, "L", "T", a.id, { isChecked: true });

    items = await todo.listChecklistItems(client, "L", "T");
    expect(items).toHaveLength(2);
    expect(items.find((i) => i.id === b.id)?.displayName).toBe("Testpunkt 2");
    expect(items.find((i) => i.id === a.id)?.isChecked).toBe(true);
  });

  it("scenario 4: task body is untouched by checklist-item add/update operations", async () => {
    const { client, getTaskBody } = makeStatefulFakeClient({ content: "original body", contentType: "text" });
    const item = await todo.addChecklistItem(client, "L", "T", "Testpunkt 1");
    await todo.updateChecklistItem(client, "L", "T", item.id, { displayName: "renamed" });

    expect(getTaskBody()).toEqual({ content: "original body", contentType: "text" });
  });

  it("scenario 5: updating unrelated task fields preserves existing checklist items", async () => {
    const { client } = makeStatefulFakeClient({ content: "original body", contentType: "text" });
    await todo.addChecklistItem(client, "L", "T", "Testpunkt 1");
    await todo.addChecklistItem(client, "L", "T", "Testpunkt 2");

    await todo.updateTask(client, "L", "T", { title: "renamed task", importance: "high" });

    const items = await todo.listChecklistItems(client, "L", "T");
    expect(items).toHaveLength(2);
  });
});
