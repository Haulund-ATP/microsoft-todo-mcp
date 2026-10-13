import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { toGraphErrorInfo } from "./client.js";

/**
 * Reads many tasks (with notes, and optionally checklist) in one tool call.
 * Verifying 24 copied tasks took ChatGPT 26 separate get_task calls and hit
 * its own per-turn tool-call limit (2026-10-13). Reads run with bounded
 * parallelism; a missing task is reported per item instead of failing the
 * whole call.
 */

export type GetTaskResult =
  | { task_id: string; status: "found"; task: todo.TaskItem; checklist?: todo.ChecklistItem[] }
  | { task_id: string; status: "not_found" | "error"; warning?: string };

const CONCURRENCY = 6;

export async function getTasksBatch(client: Client, listId: string, taskIds: string[], includeChecklist: boolean) {
  const results: GetTaskResult[] = new Array(taskIds.length);
  let next = 0;
  async function worker() {
    while (next < taskIds.length) {
      const index = next++;
      const taskId = taskIds[index]!;
      try {
        const [task, checklist] = await Promise.all([
          todo.getTask(client, listId, taskId),
          includeChecklist ? todo.listChecklistItems(client, listId, taskId) : Promise.resolve(undefined),
        ]);
        results[index] = { task_id: taskId, status: "found", task, checklist };
      } catch (err) {
        // Never surface raw Graph/auth errors (they can contain internal data).
        results[index] =
          toGraphErrorInfo(err).statusCode === 404
            ? { task_id: taskId, status: "not_found" }
            : { task_id: taskId, status: "error", warning: "The task could not be read; retry this one with get_task." };
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, taskIds.length) }, () => worker()));
  return {
    requested: taskIds.length,
    found: results.filter((r) => r.status === "found").length,
    not_found: results.filter((r) => r.status === "not_found").length,
    errors: results.filter((r) => r.status === "error").length,
    results,
  };
}
