import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "../graph/todoApi.js";
import { SharedListChecklistWriteBlockedError } from "./errors.js";
import { childLogger } from "../logging/logger.js";

export const acknowledgeSharedListRiskField = z
  .boolean()
  .optional()
  .default(false)
  .describe(
    "Must be true to write a checklist item on a list this connection does not own (isOwner: false). " +
      "Shared-list checklist writes can be silently lost due to a known Microsoft To Do sync issue — see docs/security.md."
  );

/**
 * Blocks checklist-item writes on shared lists the current connection does
 * not own, unless explicitly acknowledged. See docs/security.md for why:
 * Microsoft's shared-list sync has repeatedly been reported to silently
 * drop checklist items (and reset task body) on subsequent edits by any
 * participant — a server-side behavior this client cannot prevent or
 * detect via the documented Graph API surface (todoTask/checklistItem
 * expose no ETag or version field to guard against it).
 */
export async function assertChecklistWriteAllowed(
  client: Client,
  listId: string,
  acknowledge: boolean,
  correlationId: string,
  toolName: string
): Promise<void> {
  const list = await todo.getTaskList(client, listId);
  const log = childLogger({
    correlationId,
    toolName,
    taskListId: listId,
    isSharedList: list.isShared,
    isListOwner: list.isOwner,
  });
  if (list.isShared && !list.isOwner && !acknowledge) {
    log.warn({ resultStatus: "error" }, "checklist write blocked: shared list, not owner, not acknowledged");
    throw new SharedListChecklistWriteBlockedError(listId);
  }
  log.info({}, "checklist write allowed");
}
