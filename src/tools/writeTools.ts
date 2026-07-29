import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import * as todo from "../graph/todoApi.js";

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
      description: "Adds a checklist (sub-item) entry to a task.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        display_name: z.string().min(1).max(255),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ list_id, task_id, display_name, connection_id }) =>
      runGraphTool(ctx, "add_checklist_item", connection_id, (client) =>
        todo.addChecklistItem(client, list_id, task_id, display_name)
      )
  );

  server.registerTool(
    "update_checklist_item",
    {
      title: "Update checklist item",
      description: "Updates a checklist item's text and/or checked state.",
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
    async ({ list_id, task_id, item_id, display_name, is_checked, connection_id }) =>
      runGraphTool(ctx, "update_checklist_item", connection_id, (client) =>
        todo.updateChecklistItem(client, list_id, task_id, item_id, { displayName: display_name, isChecked: is_checked })
      )
  );

  server.registerTool(
    "delete_checklist_item",
    {
      title: "Delete checklist item",
      description: "Permanently deletes a checklist item.",
      inputSchema: {
        list_id: z.string().min(1),
        task_id: z.string().min(1),
        item_id: z.string().min(1),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ list_id, task_id, item_id, connection_id }) =>
      runGraphTool(ctx, "delete_checklist_item", connection_id, async (client) => {
        await todo.deleteChecklistItem(client, list_id, task_id, item_id);
        return { deleted: true, list_id, task_id, item_id };
      })
  );
}
