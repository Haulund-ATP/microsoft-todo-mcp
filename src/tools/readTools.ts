import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import * as todo from "../graph/todoApi.js";

const isoDateTime = z.string().datetime({ offset: true }).or(z.string().date());

export function registerReadTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "list_task_lists",
    {
      title: "List task lists",
      description: "Lists all Microsoft To Do task lists (e.g. 'Tasks', 'Groceries') for the resolved connection.",
      inputSchema: { ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ connection_id }) =>
      runGraphTool(ctx, "list_task_lists", connection_id, (client) => todo.listTaskLists(client))
  );

  server.registerTool(
    "get_task_list",
    {
      title: "Get task list",
      description: "Gets a single task list by id.",
      inputSchema: { list_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, connection_id }) =>
      runGraphTool(ctx, "get_task_list", connection_id, (client) => todo.getTaskList(client, list_id))
  );

  server.registerTool(
    "list_tasks",
    {
      title: "List tasks",
      description: "Lists tasks within a task list, optionally capped at max_items (default 50, max 500).",
      inputSchema: {
        list_id: z.string().min(1),
        max_items: z.number().int().positive().max(500).default(50),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, max_items, connection_id }) =>
      runGraphTool(ctx, "list_tasks", connection_id, (client) => todo.listTasks(client, list_id, { maxItems: max_items }))
  );

  server.registerTool(
    "get_task",
    {
      title: "Get task",
      description: "Gets a single task by id within a task list.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "get_task", connection_id, (client) => todo.getTask(client, list_id, task_id))
  );

  server.registerTool(
    "list_checklist_items",
    {
      title: "List checklist items",
      description: "Lists the checklist (sub-item) entries on a task, including their item_id, text, and checked state.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "list_checklist_items", connection_id, (client) =>
        todo.listChecklistItems(client, list_id, task_id)
      )
  );

  server.registerTool(
    "get_task_with_checklist",
    {
      title: "Get task with checklist",
      description:
        "Fetches a task together with its full checklist and parent-list metadata in a single call, always with " +
        "fresh Graph reads (list, task, and checklistItems are each fetched live — nothing is cached). Useful for " +
        "verifying actual server-side state, e.g. across two accounts on a shared list.",
      inputSchema: { list_id: z.string().min(1), task_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, task_id, connection_id }) =>
      runGraphTool(ctx, "get_task_with_checklist", connection_id, async (client) => {
        const [list, task, checklist] = await Promise.all([
          todo.getTaskList(client, list_id),
          todo.getTask(client, list_id, task_id),
          todo.listChecklistItems(client, list_id, task_id),
        ]);
        return {
          list: { id: list.id, displayName: list.displayName, isOwner: list.isOwner, isShared: list.isShared },
          task: { id: task.id, title: task.title, status: task.status, etag: task.etag, lastModifiedDateTime: task.lastModifiedDateTime },
          checklist,
          checklistCount: checklist.length,
        };
      })
  );

  server.registerTool(
    "search_tasks",
    {
      title: "Search tasks",
      description: "Searches for tasks whose title contains the given text, within a specific task list.",
      inputSchema: {
        list_id: z.string().min(1),
        query: z.string().min(1).max(200),
        max_items: z.number().int().positive().max(200).default(50),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, query, max_items, connection_id }) =>
      runGraphTool(ctx, "search_tasks", connection_id, (client) => todo.searchTasks(client, list_id, query, max_items))
  );

  server.registerTool(
    "list_overdue_tasks",
    {
      title: "List overdue tasks",
      description: "Lists incomplete tasks in a list whose due date is before now.",
      inputSchema: {
        list_id: z.string().min(1),
        max_items: z.number().int().positive().max(500).default(100),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, max_items, connection_id }) =>
      runGraphTool(ctx, "list_overdue_tasks", connection_id, async (client) => {
        const nowIso = new Date().toISOString().replace("Z", "");
        return todo.listTasks(client, list_id, {
          filter: `status ne 'completed' and dueDateTime/dateTime le '${nowIso}'`,
          maxItems: max_items,
        });
      })
  );

  server.registerTool(
    "list_tasks_due_between",
    {
      title: "List tasks due between two dates",
      description: "Lists tasks in a list whose due date falls within [from, to] (inclusive), ISO 8601.",
      inputSchema: {
        list_id: z.string().min(1),
        from: isoDateTime,
        to: isoDateTime,
        max_items: z.number().int().positive().max(500).default(100),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, from, to, max_items, connection_id }) =>
      runGraphTool(ctx, "list_tasks_due_between", connection_id, async (client) => {
        const fromRaw = new Date(from).toISOString().replace("Z", "");
        const toRaw = new Date(to).toISOString().replace("Z", "");
        return todo.listTasks(client, list_id, {
          filter: `dueDateTime/dateTime ge '${fromRaw}' and dueDateTime/dateTime le '${toRaw}'`,
          maxItems: max_items,
        });
      })
  );

  server.registerTool(
    "list_completed_tasks",
    {
      title: "List completed tasks",
      description: "Lists tasks in a list with status = completed.",
      inputSchema: {
        list_id: z.string().min(1),
        max_items: z.number().int().positive().max(500).default(100),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, max_items, connection_id }) =>
      runGraphTool(ctx, "list_completed_tasks", connection_id, (client) =>
        todo.listTasks(client, list_id, { filter: `status eq 'completed'`, maxItems: max_items })
      )
  );
}
