import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import * as todo from "../graph/todoApi.js";
import { assertChecklistWriteAllowed, acknowledgeSharedListRiskField } from "./sharedListGuard.js";
import { childLogger } from "../logging/logger.js";

const isoDateTime = z.string().datetime({ offset: true }).or(z.string().date());
const importance = z.enum(["low", "normal", "high"]);
const timezoneField = z.string().optional().describe("IANA timezone, e.g. Europe/Copenhagen. Defaults to the server's configured timezone.");

export function registerWriteTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "create_task",
    {
      title: "Create task",
      description: "Creates a new task in a task list.",
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
      description: "Updates fields on an existing task. Pass null for due_date_time/reminder_date_time to clear them.",
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
      runGraphTool(ctx, "update_task", connection_id, (client) =>
        todo.updateTask(client, list_id, task_id, {
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
    "complete_task",
    {
      title: "Complete task",
      description: "Marks a task as completed.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "complete_task", connection_id, (client) => todo.completeTask(client, list_id, task_id))
  );

  server.registerTool(
    "reopen_task",
    {
      title: "Reopen task",
      description: "Reopens a completed task (sets status back to notStarted).",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "reopen_task", connection_id, (client) => todo.reopenTask(client, list_id, task_id))
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
        "Adds a checklist (sub-item) entry to a task, preserving all other existing checklist items. " +
        "On shared lists this connection does not own, requires acknowledge_shared_list_risk: true — see docs/security.md.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        display_name: z.string().min(1).max(255),
        acknowledge_shared_list_risk: acknowledgeSharedListRiskField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, display_name, acknowledge_shared_list_risk, connection_id }) =>
      runGraphTool(ctx, "add_checklist_item", connection_id, async (client) => {
        await assertChecklistWriteAllowed(client, list_id, acknowledge_shared_list_risk, ctx.correlationId, "add_checklist_item");
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "add_checklist_item",
          taskListId: list_id,
          taskId: task_id,
          graphEndpoint: "/todo/lists/{id}/tasks/{id}/checklistItems",
          graphMethod: "POST",
        });
        try {
          const result = await todo.addChecklistItem(client, list_id, task_id, display_name);
          log.info({ checklistItemId: result.id }, "checklist item added");
          return result;
        } catch (err) {
          log.warn({}, "add_checklist_item Graph call failed");
          throw err;
        }
      })
  );

  server.registerTool(
    "update_checklist_item",
    {
      title: "Update checklist item",
      description:
        "Updates only the specified checklist item's text and/or checked state — never touches the task's " +
        "title, body, status, or other checklist items. On shared lists this connection does not own, requires " +
        "acknowledge_shared_list_risk: true — see docs/security.md.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        item_id: z.string().min(1),
        display_name: z.string().min(1).max(255).optional(),
        is_checked: z.boolean().optional(),
        acknowledge_shared_list_risk: acknowledgeSharedListRiskField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, item_id, display_name, is_checked, acknowledge_shared_list_risk, connection_id }) =>
      runGraphTool(ctx, "update_checklist_item", connection_id, async (client) => {
        await assertChecklistWriteAllowed(client, list_id, acknowledge_shared_list_risk, ctx.correlationId, "update_checklist_item");
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "update_checklist_item",
          taskListId: list_id,
          taskId: task_id,
          checklistItemId: item_id,
          graphEndpoint: "/todo/lists/{id}/tasks/{id}/checklistItems/{id}",
          graphMethod: "PATCH",
        });
        try {
          const result = await todo.updateChecklistItem(client, list_id, task_id, item_id, {
            displayName: display_name,
            isChecked: is_checked,
          });
          log.info({}, "checklist item updated");
          return result;
        } catch (err) {
          log.warn({}, "update_checklist_item Graph call failed");
          throw err;
        }
      })
  );

  server.registerTool(
    "delete_checklist_item",
    {
      title: "Delete checklist item",
      description:
        "Permanently deletes a checklist item. This cannot be undone. On shared lists this connection does not " +
        "own, requires acknowledge_shared_list_risk: true — see docs/security.md.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        item_id: z.string().min(1),
        acknowledge_shared_list_risk: acknowledgeSharedListRiskField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ list_id, task_id, item_id, acknowledge_shared_list_risk, connection_id }) =>
      runGraphTool(ctx, "delete_checklist_item", connection_id, async (client) => {
        await assertChecklistWriteAllowed(client, list_id, acknowledge_shared_list_risk, ctx.correlationId, "delete_checklist_item");
        const log = childLogger({
          correlationId: ctx.correlationId,
          toolName: "delete_checklist_item",
          taskListId: list_id,
          taskId: task_id,
          checklistItemId: item_id,
          graphEndpoint: "/todo/lists/{id}/tasks/{id}/checklistItems/{id}",
          graphMethod: "DELETE",
        });
        try {
          await todo.deleteChecklistItem(client, list_id, task_id, item_id);
          log.info({}, "checklist item deleted");
          return { deleted: true, list_id, task_id, item_id };
        } catch (err) {
          log.warn({}, "delete_checklist_item Graph call failed");
          throw err;
        }
      })
  );
}
