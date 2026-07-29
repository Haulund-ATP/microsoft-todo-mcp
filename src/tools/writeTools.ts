import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import * as todo from "../graph/todoApi.js";
import { withGraphDiagnostics, type GraphCallDiagnostics } from "../graph/diagnostics.js";
import {
  verifyChecklistItemAdded,
  verifyChecklistItemUpdated,
  verifyChecklistItemDeleted,
  withChecklistConsistencyCheck,
} from "../graph/verification.js";
import { childLogger } from "../logging/logger.js";

const isoDateTime = z.string().datetime({ offset: true }).or(z.string().date());
const importance = z.enum(["low", "normal", "high"]);
const timezoneField = z.string().optional().describe("IANA timezone, e.g. Europe/Copenhagen. Defaults to the server's configured timezone.");

/** Picks the diagnostic entry for the actual write call (first non-GET) out of a scope that may also contain verification GETs. */
function writeDiagnostic(diagnostics: GraphCallDiagnostics[]): GraphCallDiagnostics | undefined {
  return diagnostics.find((d) => d.method !== "GET") ?? diagnostics[0];
}

export function registerWriteTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "create_task",
    {
      title: "Create task",
      description: "Creates a new task in a task list. Works on shared lists.",
      inputSchema: {
        list_id: z.string().min(1),
        title: z.string().min(1).max(255),
        body: z.string().max(10000).optional(),
        importance: importance.optional(),
        due_date_time: isoDateTime.optional(),
        reminder_date_time: isoDateTime.optional(),
        time_zone: timezoneField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, title, body, importance: imp, due_date_time, reminder_date_time, time_zone, connection_id }) =>
      runGraphTool(ctx, "create_task", connection_id, (client) =>
        todo.createTask(client, list_id, {
          title,
          body,
          importance: imp,
          dueDateTime: due_date_time,
          reminderDateTime: reminder_date_time,
          timeZone: time_zone,
        })
      )
  );

  server.registerTool(
    "update_task",
    {
      title: "Update task",
      description:
        "Updates fields on an existing task. Pass null for due_date_time/reminder_date_time to clear them. Works on shared lists.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        title: z.string().min(1).max(255).optional(),
        body: z.string().max(10000).optional(),
        importance: importance.optional(),
        due_date_time: isoDateTime.nullable().optional(),
        reminder_date_time: isoDateTime.nullable().optional(),
        time_zone: timezoneField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, title, body, importance: imp, due_date_time, reminder_date_time, time_zone, connection_id }) =>
      runGraphTool(ctx, "update_task", connection_id, async (client, connectionId) => {
        const { result, check } = await withChecklistConsistencyCheck(client, list_id, task_id, () =>
          todo.updateTask(client, list_id, task_id, {
            title,
            body,
            importance: imp,
            dueDateTime: due_date_time,
            reminderDateTime: reminder_date_time,
            timeZone: time_zone,
          })
        );
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "update_task",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          taskEtag: result.etag,
          taskLastModifiedDateTime: result.lastModifiedDateTime,
          checklistCountBefore: check.checklistCountBefore,
          checklistCountAfter: check.checklistCountAfter,
          missingItemIds: check.missingItemIds,
        });
        if (!check.consistent) {
          log.warn({}, "update_task: checklist items missing after task update");
        } else {
          log.info({}, "update_task completed");
        }
        return { task: result, checklistConsistency: check };
      })
  );

  server.registerTool(
    "complete_task",
    {
      title: "Complete task",
      description: "Marks a task as completed. Works on shared lists.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "complete_task", connection_id, async (client, connectionId) => {
        const { result, check } = await withChecklistConsistencyCheck(client, list_id, task_id, () =>
          todo.completeTask(client, list_id, task_id)
        );
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "complete_task",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          taskEtag: result.etag,
          checklistCountBefore: check.checklistCountBefore,
          checklistCountAfter: check.checklistCountAfter,
          missingItemIds: check.missingItemIds,
        });
        if (!check.consistent) log.warn({}, "complete_task: checklist items missing after completion");
        else log.info({}, "complete_task completed");
        return { task: result, checklistConsistency: check };
      })
  );

  server.registerTool(
    "reopen_task",
    {
      title: "Reopen task",
      description: "Reopens a completed task (sets status back to notStarted). Works on shared lists.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "reopen_task", connection_id, async (client, connectionId) => {
        const { result, check } = await withChecklistConsistencyCheck(client, list_id, task_id, () =>
          todo.reopenTask(client, list_id, task_id)
        );
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "reopen_task",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          taskEtag: result.etag,
          checklistCountBefore: check.checklistCountBefore,
          checklistCountAfter: check.checklistCountAfter,
          missingItemIds: check.missingItemIds,
        });
        if (!check.consistent) log.warn({}, "reopen_task: checklist items missing after reopen");
        else log.info({}, "reopen_task completed");
        return { task: result, checklistConsistency: check };
      })
  );

  server.registerTool(
    "delete_task",
    {
      title: "Delete task",
      description: "Permanently deletes a task. This cannot be undone.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "delete_task", connection_id, async (client) => {
        await todo.deleteTask(client, list_id, task_id);
        return { deleted: true, list_id, task_id };
      })
  );

  server.registerTool(
    "create_task_list",
    {
      title: "Create task list",
      description: "Creates a new task list.",
      inputSchema: { display_name: z.string().min(1).max(255), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ display_name, connection_id }) =>
      runGraphTool(ctx, "create_task_list", connection_id, (client) => todo.createTaskList(client, display_name))
  );

  server.registerTool(
    "rename_task_list",
    {
      title: "Rename task list",
      description: "Renames an existing task list.",
      inputSchema: { list_id: z.string().min(1), display_name: z.string().min(1).max(255), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, display_name, connection_id }) =>
      runGraphTool(ctx, "rename_task_list", connection_id, (client) => todo.renameTaskList(client, list_id, display_name))
  );

  server.registerTool(
    "delete_task_list",
    {
      title: "Delete task list",
      description: "Permanently deletes a task list and all tasks in it. This cannot be undone.",
      inputSchema: { list_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ list_id, connection_id }) =>
      runGraphTool(ctx, "delete_task_list", connection_id, async (client) => {
        await todo.deleteTaskList(client, list_id);
        return { deleted: true, list_id };
      })
  );

  server.registerTool(
    "add_checklist_item",
    {
      title: "Add checklist item",
      description:
        "Adds a checklist (sub-item) entry to a task, preserving all other existing checklist items. Works on shared " +
        "lists. Re-reads the checklist after writing and reports verificationStatus (verified/delayed/inconsistent) " +
        "so callers can detect sync issues rather than trusting Graph's success response alone.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        display_name: z.string().min(1).max(255),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, display_name, connection_id }) =>
      runGraphTool(ctx, "add_checklist_item", connection_id, async (client, connectionId) => {
        const { result, diagnostics } = await withGraphDiagnostics(async () => {
          const added = await todo.addChecklistItem(client, list_id, task_id, display_name);
          const verification = await verifyChecklistItemAdded(client, list_id, task_id, added.id);
          return { added, verification };
        });
        const gd = writeDiagnostic(diagnostics);
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "add_checklist_item",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          checklistItemId: result.added.id,
          taskEtag: result.verification.taskEtag,
          taskLastModifiedDateTime: result.verification.taskLastModifiedDateTime,
          checklistCountAfter: result.verification.checklistCount,
          verificationStatus: result.verification.verificationStatus,
          graphRequestId: gd?.graphRequestId,
          graphClientRequestId: gd?.clientRequestId,
        });
        if (result.verification.verificationStatus === "inconsistent") log.warn({}, "add_checklist_item: verification failed");
        else log.info({}, "checklist item added");
        return {
          item: result.added,
          checklist: result.verification.checklist,
          checklistCount: result.verification.checklistCount,
          taskEtag: result.verification.taskEtag,
          taskLastModifiedDateTime: result.verification.taskLastModifiedDateTime,
          verificationStatus: result.verification.verificationStatus,
          warning: result.verification.warning,
        };
      })
  );

  server.registerTool(
    "update_checklist_item",
    {
      title: "Update checklist item",
      description:
        "Updates only the specified checklist item's text and/or checked state — never touches the task's title, " +
        "body, status, or other checklist items. At least one of display_name/is_checked must be provided. Works " +
        "on shared lists. Re-reads the checklist after writing and reports verificationStatus.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        item_id: z.string().min(1),
        display_name: z.string().min(1).max(255).optional(),
        is_checked: z.boolean().optional(),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, item_id, display_name, is_checked, connection_id }) => {
      if (display_name === undefined && is_checked === undefined) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: JSON.stringify(
                { errorType: "validation_error", message: "At least one of display_name or is_checked must be provided." },
                null,
                2
              ),
            },
          ],
        };
      }
      return runGraphTool(ctx, "update_checklist_item", connection_id, async (client, connectionId) => {
        const fieldsSent = [
          ...(display_name !== undefined ? ["display_name"] : []),
          ...(is_checked !== undefined ? ["is_checked"] : []),
        ];
        const { result, diagnostics } = await withGraphDiagnostics(async () => {
          const updated = await todo.updateChecklistItem(client, list_id, task_id, item_id, {
            displayName: display_name,
            isChecked: is_checked,
          });
          const verification = await verifyChecklistItemUpdated(client, list_id, task_id, item_id, {
            displayName: display_name,
            isChecked: is_checked,
          });
          return { updated, verification };
        });
        const gd = writeDiagnostic(diagnostics);
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "update_checklist_item",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          checklistItemId: item_id,
          taskEtag: result.verification.taskEtag,
          taskLastModifiedDateTime: result.verification.taskLastModifiedDateTime,
          checklistCountAfter: result.verification.checklistCount,
          verificationStatus: result.verification.verificationStatus,
          graphRequestId: gd?.graphRequestId,
          graphClientRequestId: gd?.clientRequestId,
          fieldsSent,
        });
        if (result.verification.verificationStatus === "inconsistent") log.warn({}, "update_checklist_item: verification failed");
        else log.info({}, "checklist item updated");
        return {
          item: result.updated,
          fieldsSent,
          checklist: result.verification.checklist,
          checklistCount: result.verification.checklistCount,
          taskEtag: result.verification.taskEtag,
          taskLastModifiedDateTime: result.verification.taskLastModifiedDateTime,
          verificationStatus: result.verification.verificationStatus,
          warning: result.verification.warning,
        };
      });
    }
  );

  server.registerTool(
    "delete_checklist_item",
    {
      title: "Delete checklist item",
      description:
        "Permanently deletes a checklist item. This cannot be undone. Works on shared lists. Re-reads the " +
        "checklist after deleting and reports verificationStatus.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        item_id: z.string().min(1),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ list_id, task_id, item_id, connection_id }) =>
      runGraphTool(ctx, "delete_checklist_item", connection_id, async (client, connectionId) => {
        const { result, diagnostics } = await withGraphDiagnostics(async () => {
          await todo.deleteChecklistItem(client, list_id, task_id, item_id);
          return verifyChecklistItemDeleted(client, list_id, task_id, item_id);
        });
        const gd = writeDiagnostic(diagnostics);
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "delete_checklist_item",
          connectionId,
          taskListId: list_id,
          taskId: task_id,
          checklistItemId: item_id,
          taskEtag: result.taskEtag,
          taskLastModifiedDateTime: result.taskLastModifiedDateTime,
          checklistCountAfter: result.checklistCount,
          verificationStatus: result.verificationStatus,
          graphRequestId: gd?.graphRequestId,
          graphClientRequestId: gd?.clientRequestId,
        });
        if (result.verificationStatus === "inconsistent") log.warn({}, "delete_checklist_item: verification failed");
        else log.info({}, "checklist item deleted");
        return {
          deleted: true,
          list_id,
          task_id,
          item_id,
          checklist: result.checklist,
          checklistCount: result.checklistCount,
          taskEtag: result.taskEtag,
          taskLastModifiedDateTime: result.taskLastModifiedDateTime,
          verificationStatus: result.verificationStatus,
          warning: result.warning,
        };
      })
  );
}
