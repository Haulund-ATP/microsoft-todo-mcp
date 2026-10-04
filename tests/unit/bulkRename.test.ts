import { describe, it, expect, vi } from "vitest";
import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import { bulkRenameTasks } from "../../src/graph/bulkRename.js";
import { bulkRenameInputShape, registerBulkRenameTools } from "../../src/tools/bulkRenameTools.js";

function fakeClient(options: { failRead?: string; failAfterPatch?: string; loseChecklist?: string; staleRead?: string } = {}) {
  const tasks = new Map(["A", "B", "C"].map((id) => [id, {
    id, title: id, status: "notStarted", importance: "normal",
    body: { content: "Skuffe 1", contentType: "text" },
  }]));
  const patches: Array<{ id: string; payload: Record<string, unknown> }> = [];
  let active = 0;
  let peak = 0;
  const client = { api(path: string) {
    const match = /^\/me\/todo\/lists\/L\/tasks\/([^/]+)(\/checklistItems)?$/.exec(path)!;
    const id = match[1]!;
    const patched = () => patches.some((p) => p.id === id);
    const chain = {
      top: () => chain,
      async get() {
        active++; peak = Math.max(peak, active);
        await new Promise((resolve) => setTimeout(resolve, 0));
        active--;
        if (options.failRead === id || (options.failAfterPatch === id && patched())) throw new Error("secret upstream data");
        if (match[2]) return { value: options.loseChecklist === id && patched() ? [] : [{ id: "step", displayName: "Keep me" }] };
        const task = tasks.get(id);
        if (!task) throw new Error("missing task");
        return { ...task, title: options.staleRead === id && patched() ? id : task.title };
      },
      async patch(payload: Record<string, unknown>) {
        patches.push({ id, payload });
        Object.assign(tasks.get(id)!, payload);
      },
    };
    return chain;
  } } as unknown as Client;
  return { client, patches, tasks, peak: () => peak };
}

const updates = [
  { task_id: "A", expected_title: "A", title: "Hylde – And" },
  { task_id: "B", expected_title: "B", title: "Skuffe 1 – Bacon" },
];

describe("bulk rename", () => {
  it("previews by default without writing", async () => {
    const fake = fakeClient();
    const result = await bulkRenameTasks(fake.client, "L", updates);
    expect(result.dry_run).toBe(true);
    expect(result.results.map((r) => r.status)).toEqual(["preview", "preview"]);
    expect(fake.patches).toEqual([]);
  });

  it("verifies title-only writes and preserves notes, status and checklist", async () => {
    const fake = fakeClient();
    const result = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(result.verified).toBe(2);
    expect(fake.patches).toEqual(expect.arrayContaining(updates.map((u) => ({ id: u.task_id, payload: { title: u.title } }))));
    expect(fake.tasks.get("A")).toMatchObject({ body: { content: "Skuffe 1" }, status: "notStarted", importance: "normal" });
    expect(result.results[0].checklistConsistency?.consistent).toBe(true);
    // Same explicit request can be retried without adding another prefix/PATCH.
    const retry = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(retry.unchanged).toBe(2);
    expect(fake.patches).toHaveLength(2);
  });

  it("skips stale titles while applying the remaining items", async () => {
    const fake = fakeClient();
    fake.tasks.get("A")!.title = "Changed elsewhere";
    const result = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(result.conflicts).toBe(1);
    expect(result.verified).toBe(1);
    expect(result.results[0]).toMatchObject({ status: "conflict", current_title: "Changed elsewhere" });
    expect(fake.patches.map((p) => p.id)).toEqual(["B"]);
  });

  it("reports partial failures without exposing upstream error text", async () => {
    const fake = fakeClient({ failRead: "A" });
    const result = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(result.errors).toBe(1);
    expect(result.verified).toBe(1);
    expect(JSON.stringify(result)).not.toContain("secret upstream data");
    expect(result.results[0].warning).toContain("no title update was attempted");
  });

  it("reports uncertainty if the verification read fails after a successful PATCH", async () => {
    const fake = fakeClient({ failAfterPatch: "A" });
    const result = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(fake.tasks.get("A")!.title).toBe(updates[0].title);
    expect(result.results[0].status).toBe("error");
    expect(result.results[0].warning).toContain("may already have applied");
    expect(result.verified).toBe(1);
  });

  it.each([{ loseChecklist: "A" }, { staleRead: "A" }])("never reports mismatched verification as success (%o)", async (options) => {
    const fake = fakeClient(options);
    const result = await bulkRenameTasks(fake.client, "L", updates, false);
    expect(result.results[0].status).toBe("inconsistent");
    expect(result.errors).toBe(1);
  });

  it("keeps input order and caps concurrent Graph operations at four", async () => {
    const fake = fakeClient();
    const many = Array.from({ length: 12 }, (_, i) => ({ task_id: String(i), expected_title: String(i), title: "new" }));
    const result = await bulkRenameTasks(fake.client, "L", many);
    expect(result.results.map((r) => r.task_id)).toEqual(many.map((u) => u.task_id));
    expect(fake.peak()).toBe(4);
  });
});

describe("bulk rename MCP schema", () => {
  const schema = z.object(bulkRenameInputShape);
  it("defaults to preview and rejects duplicates, empty titles and batches above 100", () => {
    expect(schema.parse({ list_id: "L", updates }).dry_run).toBe(true);
    expect(schema.safeParse({ list_id: "L", updates: [updates[0], updates[0]] }).success).toBe(false);
    expect(schema.safeParse({ list_id: "L", updates: [{ ...updates[0], title: "   " }] }).success).toBe(false);
    expect(schema.safeParse({ list_id: "L", updates: [] }).success).toBe(false);
    expect(schema.safeParse({ list_id: "L", updates: Array.from({ length: 101 }, (_, i) => ({ ...updates[0], task_id: String(i) })) }).success).toBe(false);
    expect(schema.safeParse({ list_id: "L", updates: [{ task_id: "A", title: "new" }] }).success).toBe(false);
  });

  it("registers the tool with write annotations and a clear sorting limitation", () => {
    const registerTool = vi.fn();
    registerBulkRenameTools({ registerTool } as never, { profileAlias: "personal", correlationId: "test" });
    expect(registerTool.mock.calls[0][0]).toBe("bulk_rename_tasks");
    expect(registerTool.mock.calls[0][1].annotations).toMatchObject({ readOnlyHint: false, destructiveHint: false });
    expect(registerTool.mock.calls[0][1].description).toContain("does NOT reorder");
  });
});
