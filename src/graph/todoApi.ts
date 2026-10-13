import type { Client } from "@microsoft/microsoft-graph-client";
import { fetchAllPages, fetchAllPagesWithMeta, withThrottleRetry } from "./pagination.js";
import { toGraphErrorInfo } from "./client.js";
import { logger } from "../logging/logger.js";
import { toGraphDateTimeTimeZone, fromGraphDateTimeTimeZone, type DisplayableDateTime } from "./time.js";

/**
 * Thin, typed wrapper around the Microsoft Graph "To Do" API
 * (`/me/todo/lists`, `.../tasks`, `.../checklistItems`). Kept separate from
 * MCP tool handlers so the handlers stay focused on schema/validation and
 * this module stays focused on Graph request shape.
 */

export interface TaskList {
  id: string;
  displayName: string;
  isOwner: boolean;
  isShared: boolean;
  wellknownListName?: string;
}

export interface TaskItem {
  id: string;
  title: string;
  status: "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred";
  importance: "low" | "normal" | "high";
  body?: { content: string; contentType: "text" | "html" };
  dueDateTime?: DisplayableDateTime;
  reminderDateTime?: DisplayableDateTime;
  completedDateTime?: DisplayableDateTime;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  /**
   * Captured from the Graph response's `@odata.etag` annotation. Confirmed
   * present on todoTask responses via the official example at
   * https://learn.microsoft.com/en-us/graph/api/todotask-update (the
   * "Properties" table on the todoTask resource page omits it, since it's an
   * OData protocol annotation rather than a named entity property — that
   * omission is what led to it being wrongly assumed absent previously).
   * Used for diagnostics only; If-Match/optimistic-concurrency support has
   * not been verified and must not be assumed.
   */
  etag?: string;
  /** True when the task has a recurrence pattern. */
  recurring?: boolean;
  categories?: string[];
}

/**
 * Small per-task view for listing, searching and summarizing many tasks:
 * no notes body, dates flattened to their stored value. Keeps a 500-task
 * response well inside a chat model's context; get_task returns the rest.
 */
export interface CompactTask {
  id: string;
  title: string;
  status: TaskItem["status"];
  importance: TaskItem["importance"];
  created?: string;
  lastModified?: string;
  due?: string;
  completed?: string;
  recurring?: boolean;
}

export function toCompactTask(task: TaskItem): CompactTask {
  return {
    id: task.id,
    title: task.title,
    status: task.status,
    importance: task.importance,
    created: task.createdDateTime,
    lastModified: task.lastModifiedDateTime,
    due: task.dueDateTime?.iso,
    completed: task.completedDateTime?.iso,
    recurring: task.recurring || undefined,
  };
}

export interface ChecklistItem {
  id: string;
  displayName: string;
  isChecked: boolean;
  createdDateTime?: string;
  /**
   * Set only once the item is checked off. This is the only
   * "last changed" signal the Graph checklistItem resource exposes — the
   * documented example response
   * (https://learn.microsoft.com/en-us/graph/api/todotask-post-checklistitems)
   * shows only `@odata.context`, `displayName`, `createdDateTime`,
   * `isChecked`, `id` — no `@odata.etag` or `lastModifiedDateTime`, unlike
   * todoTask. Verified against the actual response example, not just the
   * resource's "Properties" table.
   */
  checkedDateTime?: string;
}

/**
 * Microsoft's current todoTask-update reference documents the `body`
 * property as: "Note that only HTML type is supported." (the response
 * example on the same page still shows `contentType: "text"` for a task
 * created before that note was added, which is why this wasn't caught
 * earlier — Graph accepts "text" without error but the docs say new writes
 * should use HTML). Plain text is escaped and newlines converted to `<br>`
 * so the caller's text renders unchanged; no semantic change to the text
 * itself.
 */
function textToHtmlBody(content: string): { content: string; contentType: "html" } {
  const escaped = content
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
  return { content: escaped.replace(/\r\n|\r|\n/g, "<br>"), contentType: "html" };
}

function mapTaskList(raw: Record<string, unknown>): TaskList {
  return {
    id: String(raw.id),
    displayName: String(raw.displayName ?? ""),
    isOwner: Boolean(raw.isOwner),
    isShared: Boolean(raw.isShared),
    wellknownListName: raw.wellknownListName as string | undefined,
  };
}

function mapTask(raw: Record<string, unknown>): TaskItem {
  return {
    id: String(raw.id),
    title: String(raw.title ?? ""),
    status: (raw.status as TaskItem["status"]) ?? "notStarted",
    importance: (raw.importance as TaskItem["importance"]) ?? "normal",
    body: raw.body as TaskItem["body"],
    dueDateTime: fromGraphDateTimeTimeZone(raw.dueDateTime as never),
    reminderDateTime: fromGraphDateTimeTimeZone(raw.reminderDateTime as never),
    completedDateTime: fromGraphDateTimeTimeZone(raw.completedDateTime as never),
    createdDateTime: raw.createdDateTime as string | undefined,
    lastModifiedDateTime: raw.lastModifiedDateTime as string | undefined,
    etag: raw["@odata.etag"] as string | undefined,
    recurring: raw.recurrence ? true : undefined,
    categories: Array.isArray(raw.categories) && raw.categories.length ? (raw.categories as string[]) : undefined,
  };
}

function mapChecklistItem(raw: Record<string, unknown>): ChecklistItem {
  return {
    id: String(raw.id),
    displayName: String(raw.displayName ?? ""),
    isChecked: Boolean(raw.isChecked),
    createdDateTime: raw.createdDateTime as string | undefined,
    checkedDateTime: raw.checkedDateTime as string | undefined,
  };
}

export async function listTaskLists(client: Client): Promise<TaskList[]> {
  const raw = await fetchAllPages<Record<string, unknown>>(client, client.api("/me/todo/lists").top(50));
  return raw.map(mapTaskList);
}

export async function getTaskList(client: Client, listId: string): Promise<TaskList> {
  const raw = await withThrottleRetry(() => client.api(`/me/todo/lists/${encodeURIComponent(listId)}`).get());
  return mapTaskList(raw);
}

export async function createTaskList(client: Client, displayName: string): Promise<TaskList> {
  const raw = await withThrottleRetry(() => client.api("/me/todo/lists").post({ displayName }));
  return mapTaskList(raw);
}

export async function renameTaskList(client: Client, listId: string, displayName: string): Promise<TaskList> {
  await withThrottleRetry(() => client.api(`/me/todo/lists/${encodeURIComponent(listId)}`).patch({ displayName }));
  return getTaskList(client, listId);
}

export async function deleteTaskList(client: Client, listId: string): Promise<void> {
  await withThrottleRetry(() => client.api(`/me/todo/lists/${encodeURIComponent(listId)}`).delete());
}

export interface ListTasksOptions {
  filter?: string;
  maxItems?: number;
  /** Fetch only the fields CompactTask needs (no notes body). */
  compact?: boolean;
}

/** Fields behind CompactTask, requested with $select when `compact` is set. */
const COMPACT_TASK_FIELDS = [
  "id",
  "title",
  "status",
  "importance",
  "createdDateTime",
  "lastModifiedDateTime",
  "dueDateTime",
  "completedDateTime",
  "recurrence",
];

/**
 * Listing ~500 tasks took 4.5–8.5 s at 50 per page (10 sequential Graph
 * round trips). Larger pages and $select cut that, but Graph's To Do docs
 * don't state the page-size ceiling or $select support for tasks, so the
 * first 400 Bad Request switches this process back to the old shape for
 * good rather than failing the call.
 */
let optimizedTaskPaging = true;

export function _resetTaskPagingForTests(): void {
  optimizedTaskPaging = true;
}

export async function listTasksWithMeta(
  client: Client,
  listId: string,
  options: ListTasksOptions = {}
): Promise<{ tasks: TaskItem[]; truncated: boolean }> {
  const build = (optimized: boolean) => {
    let request = client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`).top(optimized ? 100 : 50);
    if (optimized && options.compact) request = request.select(COMPACT_TASK_FIELDS.join(","));
    if (options.filter) request = request.filter(options.filter);
    return request;
  };
  const run = async (optimized: boolean) => {
    const { items, truncated } = await fetchAllPagesWithMeta<Record<string, unknown>>(client, build(optimized), {
      maxItems: options.maxItems,
    });
    return { tasks: items.map(mapTask), truncated };
  };
  if (!optimizedTaskPaging) return run(false);
  try {
    return await run(true);
  } catch (err) {
    if (toGraphErrorInfo(err).statusCode !== 400) throw err;
    // Only blame the paging shape if the old shape then succeeds; a 400
    // caused by the filter itself fails both ways and is rethrown.
    const result = await run(false);
    optimizedTaskPaging = false;
    logger.warn({ graphStatusCode: 400 }, "Graph rejected $top=100/$select on tasks; falling back to 50 per page");
    return result;
  }
}

export async function listTasks(client: Client, listId: string, options: ListTasksOptions = {}): Promise<TaskItem[]> {
  return (await listTasksWithMeta(client, listId, options)).tasks;
}

export async function getTask(client: Client, listId: string, taskId: string): Promise<TaskItem> {
  const raw = await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`).get()
  );
  return mapTask(raw);
}

export interface CreateTaskInput {
  title: string;
  body?: string;
  importance?: TaskItem["importance"];
  dueDateTime?: string;
  reminderDateTime?: string;
  timeZone?: string;
}

export async function createTask(client: Client, listId: string, input: CreateTaskInput): Promise<TaskItem> {
  const payload: Record<string, unknown> = {
    title: input.title,
    importance: input.importance ?? "normal",
  };
  if (input.body) payload.body = textToHtmlBody(input.body);
  if (input.dueDateTime) payload.dueDateTime = toGraphDateTimeTimeZone(input.dueDateTime, input.timeZone);
  if (input.reminderDateTime) {
    payload.reminderDateTime = toGraphDateTimeTimeZone(input.reminderDateTime, input.timeZone);
    payload.isReminderOn = true;
  }
  const raw = await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`).post(payload)
  );
  return mapTask(raw);
}

export interface UpdateTaskInput {
  title?: string;
  body?: string;
  importance?: TaskItem["importance"];
  dueDateTime?: string | null;
  reminderDateTime?: string | null;
  timeZone?: string;
  status?: TaskItem["status"];
}

export async function updateTask(client: Client, listId: string, taskId: string, input: UpdateTaskInput): Promise<TaskItem> {
  const payload: Record<string, unknown> = {};
  if (input.title !== undefined) payload.title = input.title;
  if (input.body !== undefined) payload.body = textToHtmlBody(input.body);
  if (input.importance !== undefined) payload.importance = input.importance;
  if (input.status !== undefined) payload.status = input.status;
  if (input.dueDateTime !== undefined) {
    payload.dueDateTime = input.dueDateTime === null ? null : toGraphDateTimeTimeZone(input.dueDateTime, input.timeZone);
  }
  if (input.reminderDateTime !== undefined) {
    payload.reminderDateTime =
      input.reminderDateTime === null ? null : toGraphDateTimeTimeZone(input.reminderDateTime, input.timeZone);
    payload.isReminderOn = input.reminderDateTime !== null;
  }
  await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`).patch(payload)
  );
  return getTask(client, listId, taskId);
}

export async function completeTask(client: Client, listId: string, taskId: string): Promise<TaskItem> {
  return updateTask(client, listId, taskId, { status: "completed" });
}

export async function reopenTask(client: Client, listId: string, taskId: string): Promise<TaskItem> {
  return updateTask(client, listId, taskId, { status: "notStarted" });
}

export async function deleteTask(client: Client, listId: string, taskId: string): Promise<void> {
  await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`).delete()
  );
}

export async function listChecklistItems(client: Client, listId: string, taskId: string): Promise<ChecklistItem[]> {
  const raw = await fetchAllPages<Record<string, unknown>>(
    client,
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`).top(50)
  );
  return raw.map(mapChecklistItem);
}

export async function addChecklistItem(
  client: Client,
  listId: string,
  taskId: string,
  displayName: string,
  isChecked?: boolean
): Promise<ChecklistItem> {
  const raw = await withThrottleRetry(() =>
    client
      .api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`)
      .post(isChecked === undefined ? { displayName } : { displayName, isChecked })
  );
  return mapChecklistItem(raw);
}

/** The task exactly as Graph returns it, for copying fields this module doesn't model. */
export async function getTaskRaw(client: Client, listId: string, taskId: string): Promise<Record<string, unknown>> {
  return withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`).get()
  );
}

/** Creates a task from an already Graph-shaped payload (used when copying a task verbatim). */
export async function createTaskRaw(client: Client, listId: string, payload: Record<string, unknown>): Promise<TaskItem> {
  const raw = await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`).post(payload)
  );
  return mapTask(raw);
}

export async function updateChecklistItem(
  client: Client,
  listId: string,
  taskId: string,
  itemId: string,
  patch: { displayName?: string; isChecked?: boolean }
): Promise<ChecklistItem> {
  const payload: Record<string, unknown> = {};
  if (patch.displayName !== undefined) payload.displayName = patch.displayName;
  if (patch.isChecked !== undefined) payload.isChecked = patch.isChecked;
  await withThrottleRetry(() =>
    client
      .api(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(itemId)}`
      )
      .patch(payload)
  );
  const raw = await withThrottleRetry(() =>
    client
      .api(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(itemId)}`
      )
      .get()
  );
  return mapChecklistItem(raw);
}

export async function deleteChecklistItem(client: Client, listId: string, taskId: string, itemId: string): Promise<void> {
  await withThrottleRetry(() =>
    client
      .api(
        `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems/${encodeURIComponent(itemId)}`
      )
      .delete()
  );
}
