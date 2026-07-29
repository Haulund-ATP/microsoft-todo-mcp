import { describe, it, expect, vi } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import { assertChecklistWriteAllowed } from "../../src/tools/sharedListGuard.js";
import { SharedListChecklistWriteBlockedError } from "../../src/tools/errors.js";

function fakeClientReturning(listResponse: Record<string, unknown>): Client {
  const chain = { get: () => Promise.resolve(listResponse) };
  return { api: vi.fn(() => chain) } as unknown as Client;
}

describe("assertChecklistWriteAllowed — shared-list checklist-write safety gate", () => {
  it("allows writes on a list the connection owns, even if shared", async () => {
    const client = fakeClientReturning({ id: "L", isOwner: true, isShared: true, displayName: "x" });
    await expect(
      assertChecklistWriteAllowed(client, "L", false, "corr-1", "add_checklist_item")
    ).resolves.toBeUndefined();
  });

  it("allows writes on a non-shared list regardless of ownership", async () => {
    const client = fakeClientReturning({ id: "L", isOwner: false, isShared: false, displayName: "x" });
    await expect(
      assertChecklistWriteAllowed(client, "L", false, "corr-1", "add_checklist_item")
    ).resolves.toBeUndefined();
  });

  it("blocks writes on a shared list the connection does not own, without acknowledgement", async () => {
    const client = fakeClientReturning({ id: "L", isOwner: false, isShared: true, displayName: "x" });
    await expect(
      assertChecklistWriteAllowed(client, "L", false, "corr-1", "add_checklist_item")
    ).rejects.toBeInstanceOf(SharedListChecklistWriteBlockedError);
  });

  it("allows writes on a shared, non-owned list once explicitly acknowledged", async () => {
    const client = fakeClientReturning({ id: "L", isOwner: false, isShared: true, displayName: "x" });
    await expect(
      assertChecklistWriteAllowed(client, "L", true, "corr-1", "add_checklist_item")
    ).resolves.toBeUndefined();
  });
});
