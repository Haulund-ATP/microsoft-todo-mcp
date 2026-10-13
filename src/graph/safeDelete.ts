import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { buildCopyPayload } from "./copyTasks.js";
import type { DeletedTaskSnapshot, DeletedTasksRepo } from "../storage/deletedTasksRepo.js";

/**
 * Deletion guard and undo for every tool that deletes a task.
 *
 * On 2026-10-13 ChatGPT "tidied up" a list by deleting open tasks and
 * re-adding them as checklist items. That silently dropped a planned
 * purchase date, and Microsoft To Do keeps nothing Graph can recover.
 * So: an open task that carries a schedule (due date, reminder or
 * recurrence) can't be deleted at all — it has to be completed first,
 * which is a separate, visible step — and every delete that does happen
 * snapshots the task first so restore_deleted_task can bring it back.
 */

export class DeletionBlockedError extends Error {
  constructor(readonly reason: string) {
    super(reason);
    this.name = "DeletionBlockedError";
  }
}

/** Why an open task must not be deleted, or undefined if it may be. */
export function deletionBlocker(raw: Record<string, unknown>): string | undefined {
  if (raw.status === "completed") return undefined;
  const schedule = [
    raw.dueDateTime ? "a due date" : undefined,
    raw.reminderDateTime && raw.isReminderOn !== false ? "a reminder" : undefined,
    raw.recurrence ? "a recurrence" : undefined,
  ].filter(Boolean);
  if (schedule.length === 0) return undefined;
  return (
    `This open task has ${schedule.join(" and ")}, so it can't be deleted. Keep it, or complete it first ` +
    `(a separate step the user should confirm) and then delete it. Don't merge it into another task's checklist: ` +
    `checklist items can't hold dates.`
  );
}

/** deletionBlocker for the mapped TaskItem the bulk tools already have in hand. */
export function deletionBlockerForTask(task: todo.TaskItem): string | undefined {
  return deletionBlocker({
    status: task.status,
    dueDateTime: task.dueDateTime,
    reminderDateTime: task.reminderDateTime,
    isReminderOn: task.reminderOn,
    recurrence: task.recurring,
  });
}

export interface SnapshotDeleter {
  /** Snapshots then deletes the task; throws DeletionBlockedError if the guard refuses. Returns the snapshot id. */
  (listId: string, taskId: string): Promise<string>;
}

export function snapshotDeleter(
  client: Client,
  connectionId: string,
  repo: Pick<DeletedTasksRepo, "save">,
  deletedBy: string,
  now: () => Date = () => new Date()
): SnapshotDeleter {
  return async (listId, taskId) => {
    const [raw, checklist] = await Promise.all([
      todo.getTaskRaw(client, listId, taskId),
      todo.listChecklistItems(client, listId, taskId),
    ]);
    const blocker = deletionBlocker(raw);
    if (blocker) throw new DeletionBlockedError(blocker);
    const { "@odata.etag": _etag, "@odata.context": _ctx, ...task } = raw;
    // Snapshot first: if saving fails, nothing is deleted.
    const snapshotId = await repo.save({
      connectionId, listId, taskId, deletedBy, deletedAt: now().toISOString(), task, checklist,
    });
    await todo.deleteTask(client, listId, taskId);
    return snapshotId;
  };
}

/** Larger lists are refused rather than deleted with a partial undo set. */
const LIST_SNAPSHOT_CAP = 2000;

/**
 * delete_task_list removes every task in one Graph call, so the guard and
 * the snapshots have to run over the whole list first. Refuses (naming up
 * to ten) if any open task has a schedule.
 */
export async function snapshotAndDeleteList(
  client: Client,
  connectionId: string,
  repo: Pick<DeletedTasksRepo, "save">,
  listId: string,
  now: () => Date = () => new Date()
): Promise<{ snapshots: number }> {
  const [list, { items, truncated }] = await Promise.all([
    todo.getTaskList(client, listId),
    todo.listTasksRaw(client, listId, LIST_SNAPSHOT_CAP),
  ]);
  if (truncated) throw new DeletionBlockedError(`This list has more than ${LIST_SNAPSHOT_CAP} tasks; delete it in the To Do app instead.`);
  const blocked = items.filter((raw) => deletionBlocker(raw));
  if (blocked.length) {
    const names = blocked.slice(0, 10).map((raw) => `"${String(raw.title ?? "")}"`).join(", ");
    throw new DeletionBlockedError(
      `The list can't be deleted: ${blocked.length} open task(s) have a due date, reminder or recurrence (${names}${blocked.length > 10 ? ", …" : ""}). ` +
        "Complete or move them first."
    );
  }

  const deletedAt = now().toISOString();
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const raw = items[next++]!;
      const taskId = String(raw.id);
      const checklist = await todo.listChecklistItems(client, listId, taskId);
      const { "@odata.etag": _etag, ...task } = raw;
      await repo.save({
        connectionId, listId, listName: list.displayName, taskId, deletedBy: "delete_task_list", deletedAt, task, checklist,
      });
    }
  };
  // All snapshots must succeed before the list is deleted.
  await Promise.all(Array.from({ length: Math.min(6, items.length) }, worker));
  await todo.deleteTaskList(client, listId);
  return { snapshots: items.length };
}

export async function restoreSnapshot(
  client: Client,
  snapshot: DeletedTaskSnapshot,
  targetListId: string | undefined
): Promise<{ task: todo.TaskItem; listId: string; checklistItems: number }> {
  const listId = targetListId ?? snapshot.listId;
  const created = await todo.createTaskRaw(client, listId, buildCopyPayload(snapshot.task));
  for (const item of snapshot.checklist) {
    await todo.addChecklistItem(client, listId, created.id, item.displayName, item.isChecked);
  }
  return { task: created, listId, checklistItems: snapshot.checklist.length };
}
