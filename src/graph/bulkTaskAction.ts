import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { toGraphErrorInfo } from "./client.js";
import { withChecklistConsistencyCheck } from "./verification.js";

/**
 * Completes, reopens or deletes many tasks in one list — e.g. clearing out
 * hundreds of long-finished tasks. Same safety model as bulk rename:
 * dry_run previews by default, every item names the title the caller
 * expects (a stale read is skipped as a conflict, never acted on), and each
 * write is confirmed by a fresh read. Deleting is limited to completed tasks
 * unless the caller explicitly opts out.
 */

export type BulkAction = "complete" | "reopen" | "delete";

export interface BulkActionItem {
  task_id: string;
  expected_title: string;
}

export type BulkActionStatus =
  | "preview"
  | "unchanged"
  | "conflict"
  | "skipped_not_completed"
  | "verified"
  | "inconsistent"
  | "error";

export interface BulkActionResult extends BulkActionItem {
  status: BulkActionStatus;
  current_title?: string;
  current_status?: string;
  warning?: string;
}

const isNotFound = (err: unknown) => toGraphErrorInfo(err).statusCode === 404;

export async function bulkTaskAction(
  client: Client,
  listId: string,
  action: BulkAction,
  items: BulkActionItem[],
  options: { dryRun: boolean; onlyCompleted: boolean }
) {
  const results: BulkActionResult[] = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const index = next++;
      const item = items[index]!;
      let writeAttempted = false;
      try {
        let current: todo.TaskItem;
        try {
          current = await todo.getTask(client, listId, item.task_id);
        } catch (err) {
          // Deleting something already gone is the desired end state.
          if (action === "delete" && isNotFound(err)) {
            results[index] = { ...item, status: "unchanged", warning: "Task no longer exists." };
            continue;
          }
          throw err;
        }
        const seen = { current_title: current.title, current_status: current.status };
        const alreadyDone =
          (action === "complete" && current.status === "completed") ||
          (action === "reopen" && current.status !== "completed");

        if (current.title !== item.expected_title) {
          results[index] = { ...item, ...seen, status: "conflict" };
        } else if (alreadyDone) {
          results[index] = { ...item, ...seen, status: "unchanged" };
        } else if (action === "delete" && options.onlyCompleted && current.status !== "completed") {
          results[index] = { ...item, ...seen, status: "skipped_not_completed" };
        } else if (options.dryRun) {
          results[index] = { ...item, ...seen, status: "preview" };
        } else if (action === "delete") {
          writeAttempted = true;
          await todo.deleteTask(client, listId, item.task_id);
          const gone = await todo.getTask(client, listId, item.task_id).then(() => false, (err) => isNotFound(err));
          results[index] = gone
            ? { ...item, ...seen, status: "verified" }
            : { ...item, ...seen, status: "inconsistent", warning: "The task was still readable after deletion." };
        } else {
          // Same checklist before/after comparison as complete_task/reopen_task,
          // given the shared-list checklist-loss history (docs/security.md).
          const { result: updated, check } = await withChecklistConsistencyCheck(client, listId, item.task_id, () => {
            writeAttempted = true;
            return action === "complete"
              ? todo.completeTask(client, listId, item.task_id)
              : todo.reopenTask(client, listId, item.task_id);
          });
          const ok = (action === "complete" ? updated.status === "completed" : updated.status !== "completed") && check.consistent;
          results[index] = ok
            ? { ...item, current_title: updated.title, current_status: updated.status, status: "verified" }
            : {
                ...item, current_title: updated.title, current_status: updated.status, status: "inconsistent",
                warning: check.consistent
                  ? "The fresh read did not show the new status."
                  : "Checklist items were missing after the status change. Inspect this task.",
              };
        }
      } catch {
        results[index] = {
          ...item,
          status: "error",
          warning: writeAttempted
            ? "The change may already have applied; read this task before retrying."
            : "The task could not be read; nothing was changed.",
        };
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(4, items.length) }, () => worker()));

  const count = (s: BulkActionStatus) => results.filter((r) => r.status === s).length;
  return {
    action,
    dry_run: options.dryRun,
    requested: items.length,
    preview: count("preview"),
    verified: count("verified"),
    unchanged: count("unchanged"),
    conflicts: count("conflict"),
    skipped_not_completed: count("skipped_not_completed"),
    errors: count("error") + count("inconsistent"),
    results,
  };
}
