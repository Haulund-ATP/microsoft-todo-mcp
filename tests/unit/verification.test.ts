import { describe, it, expect, vi } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import {
  verifyChecklistItemAdded,
  verifyChecklistItemUpdated,
  verifyChecklistItemDeleted,
  withChecklistConsistencyCheck,
} from "../../src/graph/verification.js";

/**
 * A fake checklistItems store whose GET only starts reflecting a given item
 * after `appearsAfterReads` prior reads have already happened — simulating
 * eventual-consistency without any wall-clock sleeping in the test.
 */
function makeEventualClient(options: {
  items: Array<{ id: string; displayName: string; isChecked: boolean }>;
  appearsAfterReads?: number; // for the item id in `pendingItemId`
  pendingItemId?: string;
  taskEtag?: string;
}): { client: Client } {
  const state = { reads: 0 };
  let lastPath = "";
  const chain = {
    top: () => chain,
    filter: () => chain,
    get: () => {
      if (lastPath.endsWith("/checklistItems")) {
        state.reads += 1;
        const visible = options.items.filter((i) => {
          if (i.id !== options.pendingItemId) return true;
          return state.reads > (options.appearsAfterReads ?? 0);
        });
        return Promise.resolve({ value: visible });
      }
      return Promise.resolve({
        id: "T",
        title: "t",
        "@odata.etag": options.taskEtag ?? 'W/"x"',
        lastModifiedDateTime: "2026-01-01T00:00:00Z",
      });
    },
    post: () => Promise.reject(new Error("unexpected post")),
    patch: () => Promise.reject(new Error("unexpected patch")),
    delete: () => Promise.reject(new Error("unexpected delete")),
  };
  const client = {
    api: vi.fn((path: string) => {
      lastPath = path;
      return chain;
    }),
  } as unknown as Client;
  return { client };
}

describe("verifyChecklistItemAdded", () => {
  it("returns verified when the item is visible on the first read", async () => {
    const { client } = makeEventualClient({ items: [{ id: "item-1", displayName: "A", isChecked: false }] });
    const result = await verifyChecklistItemAdded(client, "L", "T", "item-1", { retries: 3, delayMs: 0 });
    expect(result.verificationStatus).toBe("verified");
    expect(result.checklistCount).toBe(1);
    expect(result.taskLastModifiedDateTime).toBe("2026-01-01T00:00:00Z");
  });

  it("returns delayed when the item only appears after a retry", async () => {
    const { client } = makeEventualClient({
      items: [{ id: "item-1", displayName: "A", isChecked: false }],
      pendingItemId: "item-1",
      appearsAfterReads: 1,
    });
    const result = await verifyChecklistItemAdded(client, "L", "T", "item-1", { retries: 3, delayMs: 0 });
    expect(result.verificationStatus).toBe("delayed");
    expect(result.warning).toBeUndefined();
  });

  it("returns inconsistent (with a warning, no extra writes) when the item never appears within retries", async () => {
    const { client } = makeEventualClient({
      items: [{ id: "item-1", displayName: "A", isChecked: false }],
      pendingItemId: "item-1",
      appearsAfterReads: 999,
    });
    const result = await verifyChecklistItemAdded(client, "L", "T", "item-1", { retries: 2, delayMs: 0 });
    expect(result.verificationStatus).toBe("inconsistent");
    expect(result.warning).toMatch(/item-1/);
    // The fake client throws on post/patch/delete — reaching this assertion at all proves no extra write was attempted.
  });
});

describe("verifyChecklistItemUpdated", () => {
  it("verifies the expected displayName/isChecked combination", async () => {
    const { client } = makeEventualClient({ items: [{ id: "item-1", displayName: "renamed", isChecked: true }] });
    const result = await verifyChecklistItemUpdated(client, "L", "T", "item-1", { displayName: "renamed", isChecked: true }, { retries: 1, delayMs: 0 });
    expect(result.verificationStatus).toBe("verified");
  });

  it("is inconsistent if the observed value doesn't match the expected patch", async () => {
    const { client } = makeEventualClient({ items: [{ id: "item-1", displayName: "old-name", isChecked: false }] });
    const result = await verifyChecklistItemUpdated(client, "L", "T", "item-1", { displayName: "renamed" }, { retries: 1, delayMs: 0 });
    expect(result.verificationStatus).toBe("inconsistent");
  });
});

describe("verifyChecklistItemDeleted", () => {
  it("verifies the item is gone", async () => {
    const { client } = makeEventualClient({ items: [] });
    const result = await verifyChecklistItemDeleted(client, "L", "T", "item-1", { retries: 1, delayMs: 0 });
    expect(result.verificationStatus).toBe("verified");
  });

  it("is inconsistent if the item is still visible", async () => {
    const { client } = makeEventualClient({ items: [{ id: "item-1", displayName: "A", isChecked: false }] });
    const result = await verifyChecklistItemDeleted(client, "L", "T", "item-1", { retries: 1, delayMs: 0 });
    expect(result.verificationStatus).toBe("inconsistent");
  });
});

describe("withChecklistConsistencyCheck", () => {
  it("reports consistent when no items go missing across an unrelated task operation", async () => {
    const { client } = makeEventualClient({ items: [{ id: "item-1", displayName: "A", isChecked: false }] });
    const { check } = await withChecklistConsistencyCheck(client, "L", "T", async () => "op-result");
    expect(check.consistent).toBe(true);
    expect(check.missingItemIds).toEqual([]);
    expect(check.checklistCountBefore).toBe(1);
    expect(check.checklistCountAfter).toBe(1);
  });

  it("calls performOperation exactly once and never auto-retries the operation itself", async () => {
    const { client } = makeEventualClient({ items: [] });
    const op = vi.fn().mockResolvedValue("done");
    await withChecklistConsistencyCheck(client, "L", "T", op);
    expect(op).toHaveBeenCalledTimes(1);
  });
});
