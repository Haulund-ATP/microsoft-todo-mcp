import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { toGraphErrorInfo } from "./client.js";
import { normalizeTitle } from "./taskQuery.js";

/**
 * Copies tasks between lists, optionally across Microsoft accounts,
 * including their checklist. Replaces the create_task_list → list_tasks →
 * list_checklist_items → create_task × N sequence a chat client otherwise
 * has to run by hand, which loses notes formatting, recurrence and checked
 * state along the way.
 *
 * Copy-only: the source is never modified. Moving is copy, check, then
 * bulk_task_action delete — keeping deletion in its own tool lets this one
 * be non-destructive, which ChatGPT's safety checks require before they let
 * a cross-account write through.
 *
 * Safety model mirrors bulk rename: dry_run previews by default; a task
 * whose title already exists in the target is skipped unless duplicates are
 * allowed, so a retried request doesn't create a second copy; every copy is
 * read back and compared.
 */

/** Fields copied verbatim. Read-only/server-owned ones (id, dates, etag, linkedResources) are not. */
const COPIED_FIELDS = [
  "title",
  "body",
  "importance",
  "status",
  "categories",
  "dueDateTime",
  "startDateTime",
  "reminderDateTime",
  "isReminderOn",
  "recurrence",
  "completedDateTime",
] as const;

/** Retry set if Graph rejects the full payload (e.g. a recurrence it won't accept on create). */
const ESSENTIAL_FIELDS = ["title", "body", "importance", "status", "dueDateTime"] as const;

export function buildCopyPayload(raw: Record<string, unknown>, fields: readonly string[] = COPIED_FIELDS): Record<string, unknown> {
  const payload: Record<string, unknown> = {};
  for (const field of fields) {
    const value = raw[field];
    if (value === undefined || value === null) continue;
    if (field === "body" && !(value as { content?: string }).content) continue;
    if (field === "categories" && Array.isArray(value) && value.length === 0) continue;
    if (field === "completedDateTime" && raw.status !== "completed") continue;
    payload[field] = value;
  }
  return payload;
}

export type CopyStatus = "preview" | "exists" | "copied" | "inconsistent" | "error";

export interface CopyResult {
  task_id: string;
  status: CopyStatus;
  title?: string;
  new_task_id?: string;
  checklist_items?: number;
  warning?: string;
}

export interface CopyOptions {
  dryRun: boolean;
  includeChecklist: boolean;
  allowDuplicates: boolean;
}

export async function copyTasks(
  source: Client,
  sourceListId: string,
  target: Client,
  targetListId: string,
  taskIds: string[],
  options: CopyOptions
) {
  const existingTitles = options.allowDuplicates
    ? new Set<string>()
    : new Set(
        (await todo.listTasks(target, targetListId, { compact: true, maxItems: 5000 })).map((t) => normalizeTitle(t.title))
      );

  const results: CopyResult[] = new Array(taskIds.length);
  let next = 0;
  async function worker() {
    while (next < taskIds.length) {
      const index = next++;
      const taskId = taskIds[index]!;
      let stage: "read" | "create" = "read";
      let newTaskId: string | undefined;
      try {
        const raw = await todo.getTaskRaw(source, sourceListId, taskId);
        const title = String(raw.title ?? "");
        const checklist = options.includeChecklist ? await todo.listChecklistItems(source, sourceListId, taskId) : [];
        const base = { task_id: taskId, title, checklist_items: checklist.length };

        if (existingTitles.has(normalizeTitle(title))) {
          results[index] = { ...base, status: "exists", warning: "A task with this title already exists in the target list; skipped." };
          continue;
        }
        if (options.dryRun) {
          results[index] = { ...base, status: "preview" };
          continue;
        }

        stage = "create";
        let warning: string | undefined;
        let created: todo.TaskItem;
        try {
          created = await todo.createTaskRaw(target, targetListId, buildCopyPayload(raw));
        } catch (err) {
          if (toGraphErrorInfo(err).statusCode !== 400) throw err;
          created = await todo.createTaskRaw(target, targetListId, buildCopyPayload(raw, ESSENTIAL_FIELDS));
          warning = "Graph rejected some fields (reminder, recurrence or categories); only title, notes, importance, status and due date were copied.";
        }
        newTaskId = created.id;
        for (const item of checklist) {
          await todo.addChecklistItem(target, targetListId, created.id, item.displayName, item.isChecked);
        }

        const [readBack, readBackChecklist] = await Promise.all([
          todo.getTask(target, targetListId, created.id),
          checklist.length ? todo.listChecklistItems(target, targetListId, created.id) : Promise.resolve([]),
        ]);
        if (readBack.title !== title || readBackChecklist.length !== checklist.length) {
          results[index] = {
            ...base, status: "inconsistent", new_task_id: created.id,
            warning: "The copy could not be confirmed on read-back. Inspect the new task before retrying or deleting the original.",
          };
          continue;
        }
        results[index] = { ...base, status: "copied", new_task_id: created.id, warning };
      } catch {
        // Never surface raw Graph/auth errors (they can contain internal data).
        results[index] = {
          task_id: taskId,
          status: "error",
          new_task_id: newTaskId,
          warning:
            stage === "read"
              ? "The source task could not be read; nothing was copied."
              : newTaskId
                ? "The task was created but its checklist may be incomplete; inspect new_task_id before retrying."
                : "Creating the copy failed or its result is unknown; check the target list before retrying.",
        };
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(4, taskIds.length) }, () => worker()));

  const count = (s: CopyStatus) => results.filter((r) => r.status === s).length;
  return {
    dry_run: options.dryRun,
    requested: taskIds.length,
    copied: count("copied"),
    skipped_existing: count("exists"),
    errors: count("error") + count("inconsistent"),
    results,
  };
}
