import type { Client } from "@microsoft/microsoft-graph-client";
import { fetchAllPages, withThrottleRetry } from "./pagination.js";
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
  top?: number;
  maxItems?: number;
}

export async function listTasks(client: Client, listId: string, options: ListTasksOptions = {}): Promise<TaskItem[]> {
  let request = client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks`).top(options.top ?? 50);
  if (options.filter) request = request.filter(options.filter);
  const raw = await fetchAllPages<Record<string, unknown>>(client, request, { maxItems: options.maxItems });
  return raw.map(mapTask);
}

export async function getTask(client: Client, listId: string, taskId: string): Promise<TaskItem> {
  const raw = await withThrottleRetry(() =>
    client.api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`).get()
  );
  return mapTask(raw);
}

export async function searchTasks(client: Client, listId: string, query: string, maxItems = 50): Promise<TaskItem[]> {
  // Graph To Do tasks do not support full-text $search; approximate with a
  // contains() filter on title, which is the documented workaround.
  const escaped = query.replace(/'/g, "''");
  return listTasks(client, listId, { filter: `contains(title,'${escaped}')`, maxItems });
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

export async function addChecklistItem(client: Client, listId: string, taskId: string, displayName: string): Promise<ChecklistItem> {
  const raw = await withThrottleRetry(() =>
    client
      .api(`/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`)
      .post({ displayName })
  );
  return mapChecklistItem(raw);
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
