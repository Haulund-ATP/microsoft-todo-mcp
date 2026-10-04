import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import { bulkRenameTasks } from "../graph/bulkRename.js";

export const bulkRenameInputShape = {
  list_id: z.string().min(1),
  updates: z.array(z.object({
    task_id: z.string().min(1),
    expected_title: z.string().max(255),
    title: z.string().trim().min(1).max(255),
  })).min(1).max(100).superRefine((updates, ctx) => {
    const seen = new Set<string>();
    updates.forEach((update, index) => {
      if (seen.has(update.task_id)) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: [index, "task_id"], message: "Duplicate task_id." });
      }
      seen.add(update.task_id);
    });
  }),
  dry_run: z.boolean().default(true),
  ...connectionIdInputShape,
};

export function registerBulkRenameTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool("bulk_rename_tasks", {
    title: "Bulk rename tasks",
    description: "Renames up to 100 existing tasks in one list, preserving notes and checking checklist IDs. " +
      "Defaults to a read-only preview (dry_run=true); use dry_run=false to apply. " +
      "Supply each task's exact expected_title from a fresh read; conflicting titles are skipped. " +
      "Returns per-task verification/errors; partial success is possible. " +
      "Use location prefixes such as 'Hylde – ' to prepare alphabetical sorting. " +
      "This does NOT reorder tasks or change the To Do app's sort setting; select Alphabetically in the app.",
    inputSchema: bulkRenameInputShape,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
  }, async ({ list_id, updates, dry_run, connection_id }) =>
    runGraphTool(ctx, "bulk_rename_tasks", connection_id, (client) => bulkRenameTasks(client, list_id, updates, dry_run))
  );
}
