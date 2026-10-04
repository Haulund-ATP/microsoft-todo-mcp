import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { withChecklistConsistencyCheck } from "./verification.js";

export interface TaskRename {
  task_id: string;
  expected_title: string;
  title: string;
}

export interface TaskRenameResult extends TaskRename {
  status: "preview" | "unchanged" | "conflict" | "verified" | "inconsistent" | "error";
  current_title?: string;
  checklistConsistency?: Awaited<ReturnType<typeof withChecklistConsistencyCheck>>["check"];
  warning?: string;
}

/** Bounded parallelism reduces MCP round trips; each PATCH still uses the
 * existing Graph throttling and field-scoped update implementation.
 * expected_title is a best-effort stale-read guard, not an atomic If-Match.
 * Retries are safe: an already-matching title is reported without a write.
 */
export async function bulkRenameTasks(client: Client, listId: string, updates: TaskRename[], dryRun = true) {
  const results: TaskRenameResult[] = new Array(updates.length);
  let next = 0;
  async function worker() {
    while (next < updates.length) {
      const index = next++;
      const update = updates[index]!;
      let writeAttempted = false;
      try {
        const current = await todo.getTask(client, listId, update.task_id);
        if (current.title === update.title) {
          results[index] = { ...update, status: "unchanged", current_title: current.title };
        } else if (current.title !== update.expected_title) {
          results[index] = { ...update, status: "conflict", current_title: current.title };
        } else if (dryRun) {
          results[index] = { ...update, status: "preview", current_title: current.title };
        } else {
          const { result, check } = await withChecklistConsistencyCheck(client, listId, update.task_id, () => {
            writeAttempted = true;
            return todo.updateTask(client, listId, update.task_id, { title: update.title });
          });
          const verified = result.title === update.title && check.consistent;
          results[index] = {
            ...update, status: verified ? "verified" : "inconsistent", current_title: result.title,
            checklistConsistency: check,
            warning: verified ? undefined : "The fresh read did not confirm the title and checklist preservation. Inspect before retrying.",
          };
        }
      } catch {
        // Never expose raw Graph/auth errors (which can contain internal data).
        // A PATCH may have succeeded even if a subsequent read failed.
        results[index] = {
          ...update, status: "error",
          warning: writeAttempted
            ? "The update may already have applied; read this task before retrying."
            : "The task could not be read or checked; no title update was attempted.",
        };
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(4, updates.length) }, () => worker()));
  return {
    dry_run: dryRun,
    requested: updates.length,
    verified: results.filter((r) => r.status === "verified").length,
    unchanged: results.filter((r) => r.status === "unchanged").length,
    conflicts: results.filter((r) => r.status === "conflict").length,
    errors: results.filter((r) => r.status === "error" || r.status === "inconsistent").length,
    results,
  };
}
