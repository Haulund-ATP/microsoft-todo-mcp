import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool, type ToolTextResult } from "./helpers.js";
import * as todo from "../graph/todoApi.js";
import { filterTasks, sortTasks, summarizeTasks } from "../graph/taskQuery.js";
import { searchTasksMulti } from "../graph/search.js";

const isoDateTime = z.string().datetime({ offset: true }).or(z.string().date());

/** Whole-list read cap for filtering/sorting/summarizing; far above any list seen (~600). */
const SCAN_CAP = 5000;

const STATUS_FILTERS = {
  all: undefined,
  open: "status ne 'completed'",
  completed: "status eq 'completed'",
} as const;

export const listTasksInputShape = {
  list_id: z.string().min(1),
  max_items: z.number().int().positive().max(500).default(50),
  status: z.enum(["all", "open", "completed"]).default("all"),
  created_after: isoDateTime.optional().describe("Only tasks created at or after this date/time."),
  created_before: isoDateTime.optional().describe("Only tasks created before this date/time."),
  completed_before: isoDateTime.optional().describe("Only completed tasks finished before this date/time (e.g. to find old completed tasks)."),
  sort: z.enum(["none", "created_asc", "created_desc", "title", "due"]).default("none"),
  ...connectionIdInputShape,
};

export const searchTasksInputShape = {
  list_id: z.string().min(1).optional().describe("Search only this list. Omit to search every list."),
  query: z.string().min(1).max(200).optional(),
  queries: z.array(z.string().min(1).max(200)).max(50).optional().describe("Several title fragments to look up in one call."),
  include_completed: z.boolean().default(true),
  max_items: z.number().int().positive().max(200).default(50).describe("Maximum matches returned per query."),
  ...connectionIdInputShape,
};

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
      description:
        "Lists tasks within a task list as compact rows (id, title, status, importance, created/due/completed " +
        "dates, recurring) — no notes; use get_task for a task's notes. Filter by status and by creation/completion " +
        "date, and sort. Returns `matched` (how many tasks pass the filters, even beyond max_items) so counts don't " +
        "need a second call. For counts, duplicates and age breakdowns of a whole list, prefer summarize_list.",
      inputSchema: listTasksInputShape,
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, max_items, status, created_after, created_before, completed_before, sort, connection_id }) =>
      runGraphTool(ctx, "list_tasks", connection_id, async (client) => {
        const clientSide = Boolean(created_after || created_before || completed_before) || sort !== "none";
        const { tasks, truncated } = await todo.listTasksWithMeta(client, list_id, {
          compact: true,
          filter: STATUS_FILTERS[status],
          // Client-side filters/sorts need the whole list to be correct.
          maxItems: clientSide ? SCAN_CAP : max_items,
        });
        const matched = sortTasks(
          filterTasks(tasks.map(todo.toCompactTask), {
            createdAfter: created_after,
            createdBefore: created_before,
            completedBefore: completed_before,
          }),
          sort
        );
        return {
          tasks: matched.slice(0, max_items),
          returned: Math.min(matched.length, max_items),
          matched: clientSide ? matched.length : undefined,
          more_available: matched.length > max_items || truncated,
        };
      })
  );

  server.registerTool(
    "summarize_list",
    {
      title: "Summarize task list",
      description:
        "Overview of a whole task list in one call, computed server-side: totals by status and importance, open, " +
        "completed, overdue and recurring counts, tasks created and completed per month, the oldest open tasks, " +
        "and groups of likely duplicates (titles equal ignoring case, accents and punctuation). Use this instead of " +
        "paging through list_tasks to count, age or de-duplicate a list.",
      inputSchema: { list_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, connection_id }) =>
      runGraphTool(ctx, "summarize_list", connection_id, async (client) => {
        const [list, { tasks, truncated }] = await Promise.all([
          todo.getTaskList(client, list_id),
          todo.listTasksWithMeta(client, list_id, { compact: true, maxItems: SCAN_CAP }),
        ]);
        return {
          list: { id: list.id, displayName: list.displayName, isShared: list.isShared },
          truncated: truncated || undefined,
          ...summarizeTasks(tasks.map(todo.toCompactTask)),
        };
      })
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
      description:
        "Finds tasks whose title contains a text, ignoring case and accents. Pass several texts in `queries` to " +
        "look them all up in one call (preferred over calling this repeatedly), and omit list_id to search every " +
        "list. Returns compact rows grouped per query, each with its list.",
      inputSchema: searchTasksInputShape,
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ list_id, query, queries, include_completed, max_items, connection_id }) => {
      const all = [...new Set([...(query ? [query] : []), ...(queries ?? [])])];
      if (all.length === 0) return validationError("Provide `query` or a non-empty `queries` array.");
      return runGraphTool(ctx, "search_tasks", connection_id, (client) =>
        searchTasksMulti(client, all, { listId: list_id, includeCompleted: include_completed, maxPerQuery: max_items })
      );
    }
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
        const tasks = await todo.listTasks(client, list_id, {
          filter: `status ne 'completed' and dueDateTime/dateTime le '${nowIso}'`,
          maxItems: max_items,
          compact: true,
        });
        return tasks.map(todo.toCompactTask);
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
        const tasks = await todo.listTasks(client, list_id, {
          filter: `dueDateTime/dateTime ge '${fromRaw}' and dueDateTime/dateTime le '${toRaw}'`,
          maxItems: max_items,
          compact: true,
        });
        return tasks.map(todo.toCompactTask);
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
      runGraphTool(ctx, "list_completed_tasks", connection_id, async (client) => {
        const tasks = await todo.listTasks(client, list_id, { filter: `status eq 'completed'`, maxItems: max_items, compact: true });
        return tasks.map(todo.toCompactTask);
      })
  );
}

function validationError(message: string): ToolTextResult {
  return { isError: true, content: [{ type: "text", text: JSON.stringify({ errorType: "validation_error", message }, null, 2) }] };
}
