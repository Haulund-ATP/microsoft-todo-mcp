import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import { copyTasks } from "../graph/copyTasks.js";
import { bulkTaskAction } from "../graph/bulkTaskAction.js";
import { restoreSnapshot, snapshotDeleter } from "../graph/safeDelete.js";
import { DeletedTasksRepo } from "../storage/deletedTasksRepo.js";
import { createGraphClientForConnection } from "../graph/client.js";
import { ProfilesRepo } from "../storage/profilesRepo.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";

const uniqueIds = <T>(key: (item: T) => string) => (items: T[], ctx: z.RefinementCtx) => {
  const seen = new Set<string>();
  items.forEach((item, index) => {
    if (seen.has(key(item))) ctx.addIssue({ code: z.ZodIssueCode.custom, path: [index], message: "Duplicate task_id." });
    seen.add(key(item));
  });
};

export const copyTasksInputShape = {
  list_id: z.string().min(1).describe("Source list."),
  task_ids: z.array(z.string().min(1)).min(1).max(100).superRefine(uniqueIds((id: string) => id)),
  target_list_id: z.string().min(1),
  target_profile: z
    .string()
    .min(1)
    .optional()
    .describe("Profile alias (e.g. 'personal') whose Microsoft account owns target_list_id. Omit for the same account."),
  target_connection_id: z.string().min(1).optional().describe("Alternative to target_profile."),
  include_checklist: z.boolean().default(true),
  allow_duplicates: z.boolean().default(false),
  dry_run: z.boolean().default(true),
  ...connectionIdInputShape,
};

export const bulkTaskActionInputShape = {
  list_id: z.string().min(1),
  action: z.enum(["complete", "reopen", "delete"]),
  items: z
    .array(z.object({ task_id: z.string().min(1), expected_title: z.string().max(255) }))
    .min(1)
    .max(200)
    .superRefine(uniqueIds((item: { task_id: string }) => item.task_id)),
  only_completed: z
    .boolean()
    .default(true)
    .describe("For delete: skip tasks that are not completed. Set false only when the user explicitly wants open tasks deleted."),
  dry_run: z.boolean().default(true),
  ...connectionIdInputShape,
};

async function resolveTargetConnection(sourceConnectionId: string, profile?: string, connectionId?: string): Promise<string> {
  if (profile && connectionId) throw new Error("Pass target_profile or target_connection_id, not both.");
  if (profile) {
    const binding = await new ProfilesRepo().get(profile);
    if (!binding) throw new Error(`Unknown target_profile "${profile}". Call list_profiles to see the bound profiles.`);
    return binding.connectionId;
  }
  if (connectionId) {
    if (!(await new ConnectionsRepo().get(connectionId))) throw new Error(`Unknown target_connection_id: ${connectionId}`);
    return connectionId;
  }
  return sourceConnectionId;
}

export function registerBulkTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "copy_tasks",
    {
      title: "Copy tasks",
      description:
        "Copies up to 100 tasks from one list to another — also to a list in another of the owner's own " +
        "profiles (target_profile, e.g. 'personal') — keeping notes, importance, status, due/reminder dates, " +
        "recurrence, categories and the checklist with checked state. Never changes or deletes the source tasks. " +
        "Defaults to a read-only preview (dry_run=true); use dry_run=false to apply. Tasks whose title already " +
        "exists in the target are skipped unless allow_duplicates=true, so retrying is safe. To move tasks, copy " +
        "them, check the result, then delete the originals with bulk_task_action. Returns per-task results; " +
        "partial success is possible.",
      inputSchema: copyTasksInputShape,
      // Copy-only, so not destructive. Moving used to be a mode of this tool,
      // which forced destructiveHint on every call; ChatGPT's safety check
      // then blocked a plain cross-profile copy before it reached the server.
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) =>
      runGraphTool(ctx, "copy_tasks", args.connection_id, async (source, sourceConnectionId) => {
        const targetConnectionId = await resolveTargetConnection(sourceConnectionId, args.target_profile, args.target_connection_id);
        const target = targetConnectionId === sourceConnectionId ? source : createGraphClientForConnection(targetConnectionId);
        return copyTasks(source, args.list_id, target, args.target_list_id, args.task_ids, {
          dryRun: args.dry_run,
          includeChecklist: args.include_checklist,
          allowDuplicates: args.allow_duplicates,
        });
      })
  );

  server.registerTool(
    "bulk_task_action",
    {
      title: "Complete, reopen or delete many tasks",
      description:
        "Completes, reopens or deletes up to 200 tasks in one list — e.g. clearing out old completed tasks found " +
        "with summarize_list or list_tasks(completed_before=…). Defaults to a read-only preview (dry_run=true); use " +
        "dry_run=false to apply. Supply each task's exact expected_title from a fresh read; mismatches are skipped as " +
        "conflicts. delete only touches completed tasks unless only_completed=false, and never deletes an open task " +
        "that has a due date, reminder or recurrence (status blocked_has_schedule) — don't fold such tasks into " +
        "another task's checklist, which can't hold dates. Each deleted task is kept for 30 days and returns a " +
        "snapshot_id: list_deleted_tasks / restore_deleted_task bring it back. Each change is verified by a fresh " +
        "read; partial success is possible.",
      inputSchema: bulkTaskActionInputShape,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    },
    async ({ list_id, action, items, only_completed, dry_run, connection_id }) =>
      runGraphTool(ctx, "bulk_task_action", connection_id, (client, connectionId) =>
        bulkTaskAction(client, list_id, action, items, {
          dryRun: dry_run,
          onlyCompleted: only_completed,
          deleter: snapshotDeleter(client, connectionId, new DeletedTasksRepo(), "bulk_task_action"),
        })
      )
  );

  server.registerTool(
    "list_deleted_tasks",
    {
      title: "List deleted tasks",
      description:
        "Lists tasks deleted through this server (delete_task, bulk_task_action) in the last 30 days for this " +
        "connection, newest first: snapshot_id, title, original list, due date, status, checklist size, and whether " +
        "it was already restored. Use restore_deleted_task to bring one back. Tasks deleted directly in the To Do " +
        "app aren't here — those are in Outlook's Deleted Items folder.",
      inputSchema: {
        include_restored: z.boolean().default(false),
        max_items: z.number().int().positive().max(200).default(50),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ include_restored, max_items, connection_id }) =>
      runGraphTool(ctx, "list_deleted_tasks", connection_id, async (_client, connectionId) => {
        const snapshots = await new DeletedTasksRepo().list(connectionId, include_restored ? max_items : 200);
        return snapshots
          .filter((s) => include_restored || !s.restoredAt)
          .slice(0, max_items)
          .map((s) => ({
            snapshot_id: s.snapshotId,
            title: s.task.title,
            list_id: s.listId,
            deleted_at: s.deletedAt,
            deleted_by: s.deletedBy,
            status: s.task.status,
            due: (s.task.dueDateTime as { dateTime?: string } | undefined)?.dateTime,
            checklist_items: s.checklist.length,
            restored_at: s.restoredAt,
            restored_task_id: s.restoredTaskId,
          }));
      })
  );

  server.registerTool(
    "restore_deleted_task",
    {
      title: "Restore a deleted task",
      description:
        "Recreates a task deleted through this server from its snapshot (see list_deleted_tasks): title, notes, " +
        "importance, status, dates, reminder, recurrence, categories and checklist with checked state. Goes back to " +
        "its original list unless target_list_id is given (needed if that list was deleted). The task gets a new id.",
      inputSchema: {
        snapshot_id: z.string().min(1),
        target_list_id: z.string().min(1).optional(),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async ({ snapshot_id, target_list_id, connection_id }) =>
      runGraphTool(ctx, "restore_deleted_task", connection_id, async (client, connectionId) => {
        const repo = new DeletedTasksRepo();
        const snapshot = await repo.get(connectionId, snapshot_id);
        if (!snapshot) throw new Error(`No deleted-task snapshot ${snapshot_id} for this connection (snapshots are kept 30 days).`);
        if (snapshot.restoredAt) {
          return { already_restored: true, restored_at: snapshot.restoredAt, restored_task_id: snapshot.restoredTaskId };
        }
        const restored = await restoreSnapshot(client, snapshot, target_list_id);
        await repo.markRestored(snapshot, restored.task.id, new Date().toISOString());
        return { restored: true, list_id: restored.listId, task: restored.task, checklist_items: restored.checklistItems };
      })
  );
}
