import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import type { ChecklistItem } from "./todoApi.js";

/**
 * Read-after-write verification for checklist operations. A Graph 2xx
 * response is not proof the write is durably visible — this module re-reads
 * the checklist (with a few short, bounded retries) and reports whether the
 * expected state was actually observed. It never retries the write itself,
 * never creates replacement items, and never deletes other items — only
 * reports what it saw.
 */

export type VerificationStatus = "verified" | "delayed" | "inconsistent";

export interface ChecklistWriteVerification {
  verificationStatus: VerificationStatus;
  checklist: ChecklistItem[];
  checklistCount: number;
  taskEtag?: string;
  taskLastModifiedDateTime?: string;
  warning?: string;
}

export interface VerifyOptions {
  retries?: number;
  delayMs?: number;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function verifyChecklistState(
  client: Client,
  listId: string,
  taskId: string,
  expect: (items: ChecklistItem[]) => boolean,
  warningIfInconsistent: string,
  options: VerifyOptions = {}
): Promise<ChecklistWriteVerification> {
  const retries = options.retries ?? 3;
  const delayMs = options.delayMs ?? 400;

  let items = await todo.listChecklistItems(client, listId, taskId);
  let attempt = 0;
  while (!expect(items) && attempt < retries) {
    await sleep(delayMs);
    items = await todo.listChecklistItems(client, listId, taskId);
    attempt += 1;
  }

  const satisfied = expect(items);
  const task = await todo.getTask(client, listId, taskId);
  return {
    verificationStatus: satisfied ? (attempt === 0 ? "verified" : "delayed") : "inconsistent",
    checklist: items,
    checklistCount: items.length,
    taskEtag: task.etag,
    taskLastModifiedDateTime: task.lastModifiedDateTime,
    warning: satisfied ? undefined : warningIfInconsistent,
  };
}

export function verifyChecklistItemAdded(
  client: Client,
  listId: string,
  taskId: string,
  itemId: string,
  options?: VerifyOptions
): Promise<ChecklistWriteVerification> {
  return verifyChecklistState(
    client,
    listId,
    taskId,
    (items) => items.some((i) => i.id === itemId),
    `Checklist item ${itemId} was reported created by Graph, but a fresh read of the checklist does not (yet) show it.`,
    options
  );
}

export function verifyChecklistItemUpdated(
  client: Client,
  listId: string,
  taskId: string,
  itemId: string,
  expectedPatch: { displayName?: string; isChecked?: boolean },
  options?: VerifyOptions
): Promise<ChecklistWriteVerification> {
  return verifyChecklistState(
    client,
    listId,
    taskId,
    (items) => {
      const item = items.find((i) => i.id === itemId);
      if (!item) return false;
      if (expectedPatch.displayName !== undefined && item.displayName !== expectedPatch.displayName) return false;
      if (expectedPatch.isChecked !== undefined && item.isChecked !== expectedPatch.isChecked) return false;
      return true;
    },
    `Checklist item ${itemId} was reported updated by Graph, but a fresh read of the checklist does not (yet) reflect the expected change.`,
    options
  );
}

export function verifyChecklistItemDeleted(
  client: Client,
  listId: string,
  taskId: string,
  itemId: string,
  options?: VerifyOptions
): Promise<ChecklistWriteVerification> {
  return verifyChecklistState(
    client,
    listId,
    taskId,
    (items) => !items.some((i) => i.id === itemId),
    `Checklist item ${itemId} was reported deleted by Graph, but a fresh read of the checklist still shows it.`,
    options
  );
}

/** Result of comparing a task's checklist before and after an unrelated task-level write (update/complete/reopen). */
export interface TaskOperationChecklistCheck {
  checklistCountBefore: number;
  checklistCountAfter: number;
  missingItemIds: string[];
  consistent: boolean;
}

/**
 * Reads the checklist before and after `performOperation`, comparing item
 * IDs. Never auto-reconstructs anything and never blocks the operation —
 * `performOperation` always runs; this only reports what changed.
 */
export async function withChecklistConsistencyCheck<T>(
  client: Client,
  listId: string,
  taskId: string,
  performOperation: () => Promise<T>
): Promise<{ result: T; check: TaskOperationChecklistCheck }> {
  const before = await todo.listChecklistItems(client, listId, taskId);
  const result = await performOperation();
  const after = await todo.listChecklistItems(client, listId, taskId);

  const afterIds = new Set(after.map((i) => i.id));
  const missingItemIds = before.filter((i) => !afterIds.has(i.id)).map((i) => i.id);

  return {
    result,
    check: {
      checklistCountBefore: before.length,
      checklistCountAfter: after.length,
      missingItemIds,
      consistent: missingItemIds.length === 0,
    },
  };
}
