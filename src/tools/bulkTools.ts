import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import { copyTasks } from "../graph/copyTasks.js";
import { bulkTaskAction } from "../graph/bulkTaskAction.js";
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
  mode: z.enum(["copy", "move"]).default("copy"),
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
      title: "Copy or move tasks",
      description:
        "Copies or moves up to 100 tasks from one list to another — also to a list in a different profile's " +
        "Microsoft account (target_profile) — keeping notes, importance, status, due/reminder dates, recurrence, " +
        "categories and the checklist with checked state. Defaults to a read-only preview (dry_run=true); use " +
        "dry_run=false to apply. Tasks whose title already exists in the target are skipped unless " +
        "allow_duplicates=true, so retrying is safe. mode='move' deletes each source task only after its copy was " +
        "read back and matched. Returns per-task results; partial success is possible.",
      inputSchema: copyTasksInputShape,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true },
    },
    async (args) =>
      runGraphTool(ctx, "copy_tasks", args.connection_id, async (source, sourceConnectionId) => {
        const targetConnectionId = await resolveTargetConnection(sourceConnectionId, args.target_profile, args.target_connection_id);
        const target = targetConnectionId === sourceConnectionId ? source : createGraphClientForConnection(targetConnectionId);
        return copyTasks(source, args.list_id, target, args.target_list_id, args.task_ids, {
          mode: args.mode,
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
        "Completes, reopens or permanently deletes up to 200 tasks in one list — e.g. clearing out old completed " +
        "tasks found with summarize_list or list_tasks(completed_before=…). Defaults to a read-only preview " +
        "(dry_run=true); use dry_run=false to apply. Supply each task's exact expected_title from a fresh read; " +
        "mismatches are skipped as conflicts. delete only touches completed tasks unless only_completed=false. " +
        "Each change is verified by a fresh read; partial success is possible. Deletion cannot be undone.",
      inputSchema: bulkTaskActionInputShape,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true },
    },
    async ({ list_id, action, items, only_completed, dry_run, connection_id }) =>
      runGraphTool(ctx, "bulk_task_action", connection_id, (client) =>
        bulkTaskAction(client, list_id, action, items, { dryRun: dry_run, onlyCompleted: only_completed })
      )
  );
}
