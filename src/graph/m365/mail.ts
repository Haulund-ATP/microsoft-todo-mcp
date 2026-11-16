import type { Client } from "@microsoft/microsoft-graph-client";
import { withThrottleRetry } from "../pagination.js";
import { graphSearchValue, htmlToText, nextPageRequest, pageOf, truncateText, type Page } from "../../tools/m365/common.js";

/**
 * Microsoft Graph v1.0 mail operations for the signed-in user (`/me/...`,
 * delegated Mail.ReadWrite + Mail.Send). Pure functions over a Graph
 * `Client` so request and output shaping can be tested with a fake client;
 * the MCP tool layer (tools/m365/mailTools.ts) handles connection
 * resolution, send gating and auditing.
 */

export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;
/** Plain-text cap for one message body (mail_get_message). */
export const MESSAGE_BODY_MAX_CHARS = 20_000;
/** Plain-text cap per message in a thread, which can hold many messages. */
export const THREAD_BODY_MAX_CHARS = 4_000;
export const THREAD_MAX_MESSAGES = 100;
/** Text-like file attachments up to this size are decoded and returned. */
export const ATTACHMENT_TEXT_MAX_BYTES = 256 * 1024;
export const ATTACHMENT_TEXT_MAX_CHARS = 20_000;

/** Well-known folder names Graph accepts wherever a mail folder id is expected. */
export const WELL_KNOWN_FOLDERS = ["inbox", "drafts", "sentitems", "deleteditems", "archive", "junkemail", "outbox", "msgfolderroot"] as const;

/** Asks Graph to return bodies as plain text; html bodies that still come back are converted locally. */
const PREFER_TEXT_BODY = 'outlook.body-content-type="text"';

const SUMMARY_SELECT = [
  "id",
  "conversationId",
  "subject",
  "from",
  "toRecipients",
  "receivedDateTime",
  "isRead",
  "isDraft",
  "hasAttachments",
  "importance",
  "flag",
  "categories",
  "parentFolderId",
  "bodyPreview",
].join(",");

const DETAIL_SELECT = [
  "id",
  "conversationId",
  "internetMessageId",
  "subject",
  "from",
  "sender",
  "toRecipients",
  "ccRecipients",
  "bccRecipients",
  "replyTo",
  "receivedDateTime",
  "sentDateTime",
  "isRead",
  "isDraft",
  "hasAttachments",
  "importance",
  "flag",
  "categories",
  "parentFolderId",
  "webLink",
  "body",
].join(",");

const THREAD_SELECT = ["id", "subject", "from", "toRecipients", "ccRecipients", "receivedDateTime", "sentDateTime", "isRead", "isDraft", "hasAttachments", "parentFolderId", "uniqueBody", "bodyPreview"].join(",");

const DRAFT_SELECT = ["id", "conversationId", "subject", "toRecipients", "ccRecipients", "bccRecipients", "importance", "isDraft", "webLink"].join(",");

const FOLDER_SELECT = ["id", "displayName", "parentFolderId", "childFolderCount", "unreadItemCount", "totalItemCount", "isHidden"].join(",");

const ATTACHMENT_SELECT = ["id", "name", "contentType", "size", "isInline", "lastModifiedDateTime"].join(",");

// ---------------------------------------------------------------- Graph shapes

interface GraphRecipient {
  emailAddress?: { name?: string; address?: string };
}

interface GraphBody {
  contentType?: string;
  content?: string;
}

interface GraphMessage {
  id: string;
  conversationId?: string;
  internetMessageId?: string;
  subject?: string;
  from?: GraphRecipient;
  sender?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  bccRecipients?: GraphRecipient[];
  replyTo?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  flag?: { flagStatus?: string };
  categories?: string[];
  parentFolderId?: string;
  webLink?: string;
  bodyPreview?: string;
  body?: GraphBody;
  uniqueBody?: GraphBody;
}

interface GraphFolder {
  id: string;
  displayName?: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
  isHidden?: boolean;
}

interface GraphAttachment {
  "@odata.type"?: string;
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
  contentBytes?: string;
}

type Collection<T> = { value?: T[]; "@odata.nextLink"?: string };

// ---------------------------------------------------------------- helpers

export interface Address {
  name?: string;
  address?: string;
}

function toAddress(r: GraphRecipient | undefined): Address | undefined {
  if (!r?.emailAddress) return undefined;
  const { name, address } = r.emailAddress;
  return { ...(name ? { name } : {}), ...(address ? { address } : {}) };
}

function toAddresses(list: GraphRecipient[] | undefined): Address[] {
  return (list ?? []).map(toAddress).filter((a): a is Address => a !== undefined);
}

function messagePath(messageId: string): string {
  return `/me/messages/${encodeURIComponent(messageId)}`;
}

function folderPath(folder: string): string {
  return `/me/mailFolders/${encodeURIComponent(folder)}`;
}

/** OData string literal: single quotes are doubled. */
function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** Accepts YYYY-MM-DD or an ISO date-time and returns an OData DateTimeOffset literal in UTC. */
export function toODataDateTime(value: string): string {
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T00:00:00Z` : value;
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) throw new Error(`Invalid date: ${value}. Use YYYY-MM-DD or an ISO 8601 date-time.`);
  return date.toISOString().replace(/\.000Z$/, "Z");
}

/** YYYY-MM-DD part of a date or date-time, for KQL date restrictions. */
function toKqlDate(value: string): string {
  return toODataDateTime(value).slice(0, 10);
}

/** Converts a Graph body to capped plain text. */
export function bodyToText(body: GraphBody | undefined, maxChars: number): { text: string; truncated: boolean } {
  const content = body?.content ?? "";
  const text = body?.contentType?.toLowerCase() === "html" ? htmlToText(content) : content.trim();
  return truncateText(text, maxChars);
}

function clampPageSize(pageSize: number | undefined): number {
  if (pageSize === undefined) return DEFAULT_PAGE_SIZE;
  return Math.min(Math.max(Math.trunc(pageSize), 1), MAX_PAGE_SIZE);
}

/** A page token must be a Graph URL from the matching kind of listing (checked again by nextPageRequest). */
function pagedRequest(client: Client, token: string, expect: RegExp) {
  const path = token.split("?")[0] ?? "";
  if (!expect.test(path)) throw new Error("Invalid next_page_token: pass back the exact value from the previous result of this tool.");
  return nextPageRequest(client, token);
}

// ---------------------------------------------------------------- output shapes

export interface MessageSummary {
  id: string;
  conversationId?: string;
  subject?: string;
  from?: Address;
  to: Address[];
  receivedDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  flagStatus?: string;
  categories?: string[];
  parentFolderId?: string;
  preview?: string;
}

export function toMessageSummary(m: GraphMessage): MessageSummary {
  return {
    id: m.id,
    conversationId: m.conversationId,
    subject: m.subject,
    from: toAddress(m.from),
    to: toAddresses(m.toRecipients),
    receivedDateTime: m.receivedDateTime,
    isRead: m.isRead,
    isDraft: m.isDraft,
    hasAttachments: m.hasAttachments,
    importance: m.importance,
    flagStatus: m.flag?.flagStatus,
    categories: m.categories,
    parentFolderId: m.parentFolderId,
    preview: m.bodyPreview,
  };
}

export interface DraftSummary {
  id: string;
  conversationId?: string;
  subject?: string;
  to: Address[];
  cc: Address[];
  bcc: Address[];
  importance?: string;
  isDraft?: boolean;
  webLink?: string;
}

function toDraftSummary(m: GraphMessage): DraftSummary {
  return {
    id: m.id,
    conversationId: m.conversationId,
    subject: m.subject,
    to: toAddresses(m.toRecipients),
    cc: toAddresses(m.ccRecipients),
    bcc: toAddresses(m.bccRecipients),
    importance: m.importance,
    isDraft: m.isDraft,
    webLink: m.webLink,
  };
}

function recipientCount(m: Pick<GraphMessage, "toRecipients" | "ccRecipients" | "bccRecipients">): number {
  const unique = new Set<string>();
  for (const r of [...(m.toRecipients ?? []), ...(m.ccRecipients ?? []), ...(m.bccRecipients ?? [])]) {
    const address = r.emailAddress?.address?.toLowerCase();
    if (address) unique.add(address);
  }
  return unique.size;
}

// ---------------------------------------------------------------- folders

export interface MailFolder {
  id: string;
  displayName?: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  totalItemCount?: number;
  isHidden?: boolean;
  childFolders?: MailFolder[];
}

function toFolder(f: GraphFolder): MailFolder {
  return {
    id: f.id,
    displayName: f.displayName,
    parentFolderId: f.parentFolderId,
    childFolderCount: f.childFolderCount,
    unreadItemCount: f.unreadItemCount,
    totalItemCount: f.totalItemCount,
    ...(f.isHidden ? { isHidden: true } : {}),
  };
}

export interface ListFoldersOptions {
  parentFolder?: string;
  includeChildren?: boolean;
  includeHidden?: boolean;
  pageSize?: number;
  nextPageToken?: string;
}

export async function listMailFolders(
  client: Client,
  opts: ListFoldersOptions = {}
): Promise<Page<MailFolder> & { well_known_names: readonly string[] }> {
  let response: Collection<GraphFolder>;
  if (opts.nextPageToken) {
    const request = pagedRequest(client, opts.nextPageToken, /\/(mailFolders|childFolders)$/i);
    response = await withThrottleRetry(() => request.get());
  } else {
    let request = client
      .api(opts.parentFolder ? `${folderPath(opts.parentFolder)}/childFolders` : "/me/mailFolders")
      .select(FOLDER_SELECT)
      .top(clampPageSize(opts.pageSize));
    if (opts.includeHidden) request = request.query({ includeHiddenFolders: "true" });
    response = await withThrottleRetry(() => request.get());
  }
  const page = pageOf(response);
  const items = page.items.map(toFolder);
  if (opts.includeChildren) {
    for (const folder of items) {
      if (!folder.childFolderCount) continue;
      const children = await withThrottleRetry(() =>
        client.api(`${folderPath(folder.id)}/childFolders`).select(FOLDER_SELECT).top(MAX_PAGE_SIZE).get()
      ) as Collection<GraphFolder>;
      folder.childFolders = (children.value ?? []).map(toFolder);
    }
  }
  return { ...page, items, well_known_names: WELL_KNOWN_FOLDERS };
}

export async function createMailFolder(client: Client, displayName: string, parentFolder?: string): Promise<MailFolder> {
  const path = parentFolder ? `${folderPath(parentFolder)}/childFolders` : "/me/mailFolders";
  const created = (await withThrottleRetry(() => client.api(path).post({ displayName }))) as GraphFolder;
  return toFolder(created);
}

// ---------------------------------------------------------------- search

export type Importance = "low" | "normal" | "high";
export type FlagStatus = "notFlagged" | "flagged" | "complete";

export interface SearchMessagesOptions {
  query?: string;
  folder?: string;
  from?: string;
  unreadOnly?: boolean;
  receivedAfter?: string;
  receivedBefore?: string;
  hasAttachments?: boolean;
  importance?: Importance;
  flagStatus?: FlagStatus;
  pageSize?: number;
  nextPageToken?: string;
}

/**
 * Structured filters as OData. Graph rejects `$filter` + `$orderby` on
 * messages ("InefficientFilter") unless every `$orderby` property also
 * appears first in `$filter`, so receivedDateTime always leads (with an
 * open lower bound when no date was given).
 */
export function buildMessageFilter(opts: SearchMessagesOptions): string | undefined {
  const rest: string[] = [];
  if (opts.receivedBefore) rest.push(`receivedDateTime lt ${toODataDateTime(opts.receivedBefore)}`);
  if (opts.from) rest.push(`from/emailAddress/address eq ${odataString(opts.from.trim())}`);
  if (opts.unreadOnly) rest.push("isRead eq false");
  if (opts.hasAttachments !== undefined) rest.push(`hasAttachments eq ${opts.hasAttachments}`);
  if (opts.importance) rest.push(`importance eq ${odataString(opts.importance)}`);
  if (opts.flagStatus) rest.push(`flag/flagStatus eq ${odataString(opts.flagStatus)}`);
  if (!opts.receivedAfter && rest.length === 0) return undefined;
  const lower = `receivedDateTime ge ${opts.receivedAfter ? toODataDateTime(opts.receivedAfter) : "1900-01-01T00:00:00Z"}`;
  return [lower, ...rest].join(" and ");
}

/**
 * Free text plus the structured filters KQL can express, as one `$search`
 * value. Graph cannot combine `$search` with `$filter` or `$orderby` on
 * messages, and KQL has no read or flag state, so those two filters are
 * refused together with a query instead of being silently dropped.
 */
export function buildMessageSearch(opts: SearchMessagesOptions): string {
  if (opts.unreadOnly || opts.flagStatus) {
    throw new Error(
      "unread_only and flag_status cannot be combined with query: Microsoft Graph does not allow $search together with $filter on messages. " +
        "Drop query and use the structured filters, or drop unread_only/flag_status."
    );
  }
  const parts = [(opts.query ?? "").replace(/"/g, "").trim()];
  if (opts.from) parts.push(`from:${opts.from.trim().replace(/\s+/g, "")}`);
  if (opts.receivedAfter) parts.push(`received>=${toKqlDate(opts.receivedAfter)}`);
  if (opts.receivedBefore) parts.push(`received<${toKqlDate(opts.receivedBefore)}`);
  if (opts.hasAttachments !== undefined) parts.push(`hasattachment:${opts.hasAttachments}`);
  if (opts.importance) parts.push(`importance:${opts.importance === "normal" ? "medium" : opts.importance}`);
  return graphSearchValue(parts.filter(Boolean).join(" AND "));
}

export interface SearchMessagesResult extends Page<MessageSummary> {
  mode: "search" | "filter" | "next_page";
  ordering: string;
}

export async function searchMessages(client: Client, opts: SearchMessagesOptions): Promise<SearchMessagesResult> {
  if (opts.nextPageToken) {
    const request = pagedRequest(client, opts.nextPageToken, /\/messages$/i);
    const response = (await withThrottleRetry(() => request.get())) as Collection<GraphMessage>;
    const page = pageOf(response);
    return { mode: "next_page", ordering: "same as the first page", ...page, items: page.items.map(toMessageSummary) };
  }
  const base = opts.folder ? `${folderPath(opts.folder)}/messages` : "/me/messages";
  let request = client.api(base).select(SUMMARY_SELECT).top(clampPageSize(opts.pageSize));
  let mode: SearchMessagesResult["mode"];
  let ordering: string;
  if (opts.query && opts.query.trim()) {
    request = request.search(buildMessageSearch(opts));
    mode = "search";
    ordering = "newest first by sent date (Graph orders search results itself; at most 1000 results)";
  } else {
    const filter = buildMessageFilter(opts);
    if (filter) request = request.filter(filter);
    request = request.orderby("receivedDateTime desc");
    mode = "filter";
    ordering = "newest first by received date";
  }
  const response = (await withThrottleRetry(() => request.get())) as Collection<GraphMessage>;
  const page = pageOf(response);
  return { mode, ordering, ...page, items: page.items.map(toMessageSummary) };
}

// ---------------------------------------------------------------- read

export interface AttachmentInfo {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  kind: "file" | "item" | "reference" | "unknown";
}

function attachmentKind(odataType: string | undefined): AttachmentInfo["kind"] {
  switch (odataType) {
    case "#microsoft.graph.fileAttachment":
      return "file";
    case "#microsoft.graph.itemAttachment":
      return "item";
    case "#microsoft.graph.referenceAttachment":
      return "reference";
    default:
      return "unknown";
  }
}

function toAttachmentInfo(a: GraphAttachment): AttachmentInfo {
  return { id: a.id, name: a.name, contentType: a.contentType, size: a.size, isInline: a.isInline, kind: attachmentKind(a["@odata.type"]) };
}

export interface MessageDetail {
  id: string;
  conversationId?: string;
  internetMessageId?: string;
  subject?: string;
  from?: Address;
  sender?: Address;
  to: Address[];
  cc: Address[];
  bcc: Address[];
  replyTo: Address[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  importance?: string;
  flagStatus?: string;
  categories?: string[];
  parentFolderId?: string;
  webLink?: string;
  body: string;
  bodyTruncated: boolean;
  attachments: AttachmentInfo[];
}

export async function getMessage(client: Client, messageId: string, maxBodyChars = MESSAGE_BODY_MAX_CHARS): Promise<MessageDetail> {
  const m = (await withThrottleRetry(() =>
    client.api(messagePath(messageId)).header("Prefer", PREFER_TEXT_BODY).select(DETAIL_SELECT).get()
  )) as GraphMessage;
  const attachments = (await withThrottleRetry(() =>
    client.api(`${messagePath(messageId)}/attachments`).select(ATTACHMENT_SELECT).get()
  )) as Collection<GraphAttachment>;
  const body = bodyToText(m.body, maxBodyChars);
  return {
    id: m.id,
    conversationId: m.conversationId,
    internetMessageId: m.internetMessageId,
    subject: m.subject,
    from: toAddress(m.from),
    sender: toAddress(m.sender),
    to: toAddresses(m.toRecipients),
    cc: toAddresses(m.ccRecipients),
    bcc: toAddresses(m.bccRecipients),
    replyTo: toAddresses(m.replyTo),
    receivedDateTime: m.receivedDateTime,
    sentDateTime: m.sentDateTime,
    isRead: m.isRead,
    isDraft: m.isDraft,
    importance: m.importance,
    flagStatus: m.flag?.flagStatus,
    categories: m.categories,
    parentFolderId: m.parentFolderId,
    webLink: m.webLink,
    body: body.text,
    bodyTruncated: body.truncated,
    attachments: (attachments.value ?? []).map(toAttachmentInfo),
  };
}

export interface ThreadMessage {
  id: string;
  subject?: string;
  from?: Address;
  to: Address[];
  cc: Address[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  hasAttachments?: boolean;
  parentFolderId?: string;
  /** Only the part this message added (Graph uniqueBody), without the quoted history. */
  body: string;
  bodyTruncated: boolean;
}

export interface ThreadResult {
  conversationId: string;
  messageCount: number;
  /** True when the conversation has more messages than were returned. */
  moreMessages: boolean;
  messages: ThreadMessage[];
}

export interface GetThreadOptions {
  messageId?: string;
  conversationId?: string;
  maxMessages?: number;
  maxBodyChars?: number;
}

export async function getThread(client: Client, opts: GetThreadOptions): Promise<ThreadResult> {
  let conversationId = opts.conversationId;
  if (!conversationId) {
    if (!opts.messageId) throw new Error("Pass message_id or conversation_id.");
    const m = (await withThrottleRetry(() => client.api(messagePath(opts.messageId!)).select("id,conversationId").get())) as GraphMessage;
    if (!m.conversationId) throw new Error("This message has no conversationId.");
    conversationId = m.conversationId;
  }
  const maxMessages = Math.min(Math.max(opts.maxMessages ?? 50, 1), THREAD_MAX_MESSAGES);
  const maxBodyChars = opts.maxBodyChars ?? THREAD_BODY_MAX_CHARS;
  // Graph rejects $orderby together with a conversationId filter, so the
  // thread is sorted here. /me/messages spans all folders (sent items too).
  let request = client
    .api("/me/messages")
    .header("Prefer", PREFER_TEXT_BODY)
    .filter(`conversationId eq ${odataString(conversationId)}`)
    .select(THREAD_SELECT)
    .top(Math.min(maxMessages, 50));
  const collected: GraphMessage[] = [];
  let more = false;
  // eslint-disable-next-line no-constant-condition -- bounded by maxMessages and nextLink
  while (true) {
    const current = request;
    const response = (await withThrottleRetry(() => current.get())) as Collection<GraphMessage>;
    collected.push(...(response.value ?? []));
    const next = response["@odata.nextLink"];
    if (!next) break;
    if (collected.length >= maxMessages) {
      more = true;
      break;
    }
    request = nextPageRequest(client, next).header("Prefer", PREFER_TEXT_BODY);
  }
  if (collected.length > maxMessages) more = true;
  const time = (m: GraphMessage) => Date.parse(m.receivedDateTime ?? m.sentDateTime ?? "") || 0;
  const messages = collected
    .sort((a, b) => time(a) - time(b))
    .slice(0, maxMessages)
    .map((m): ThreadMessage => {
      const body = m.uniqueBody?.content ? bodyToText(m.uniqueBody, maxBodyChars) : truncateText(m.bodyPreview ?? "", maxBodyChars);
      return {
        id: m.id,
        subject: m.subject,
        from: toAddress(m.from),
        to: toAddresses(m.toRecipients),
        cc: toAddresses(m.ccRecipients),
        receivedDateTime: m.receivedDateTime,
        sentDateTime: m.sentDateTime,
        isRead: m.isRead,
        isDraft: m.isDraft,
        hasAttachments: m.hasAttachments,
        parentFolderId: m.parentFolderId,
        body: body.text,
        bodyTruncated: body.truncated,
      };
    });
  return { conversationId, messageCount: messages.length, moreMessages: more, messages };
}

const TEXT_CONTENT_TYPES = new Set([
  "application/json",
  "application/xml",
  "application/csv",
  "application/ics",
  "application/x-yaml",
  "application/yaml",
  "application/javascript",
  "application/x-sh",
]);
const TEXT_EXTENSIONS = new Set(["txt", "csv", "tsv", "json", "xml", "ics", "vcs", "md", "log", "yaml", "yml", "html", "htm", "eml", "vcf", "ini", "sql"]);

/** Whether an attachment can be returned as decoded text (mail often labels files application/octet-stream, so the extension counts too). */
export function isTextLikeAttachment(contentType: string | undefined, name: string | undefined): boolean {
  const type = (contentType ?? "").toLowerCase().split(";")[0]!.trim();
  if (type.startsWith("text/") || TEXT_CONTENT_TYPES.has(type) || type.endsWith("+json") || type.endsWith("+xml")) return true;
  const ext = (name ?? "").toLowerCase().split(".").pop() ?? "";
  return (name ?? "").includes(".") && TEXT_EXTENSIONS.has(ext);
}

export interface AttachmentResult extends AttachmentInfo {
  contentReturned: boolean;
  text?: string;
  truncated?: boolean;
  reason?: string;
}

export async function getAttachment(
  client: Client,
  messageId: string,
  attachmentId: string,
  opts: { maxBytes?: number; maxChars?: number } = {}
): Promise<AttachmentResult> {
  const maxBytes = opts.maxBytes ?? ATTACHMENT_TEXT_MAX_BYTES;
  const maxChars = opts.maxChars ?? ATTACHMENT_TEXT_MAX_CHARS;
  const path = `${messagePath(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
  // Metadata first, so large files are never downloaded as base64.
  const meta = (await withThrottleRetry(() => client.api(path).select(ATTACHMENT_SELECT).get())) as GraphAttachment;
  const info = toAttachmentInfo(meta);
  if (info.kind !== "file") {
    return { ...info, contentReturned: false, reason: "Only file attachments have content; open item or cloud attachments in Outlook." };
  }
  if (!isTextLikeAttachment(meta.contentType, meta.name)) {
    return { ...info, contentReturned: false, reason: "Not a text file; only metadata is returned. Open it in Outlook." };
  }
  if ((meta.size ?? Number.POSITIVE_INFINITY) > maxBytes) {
    return { ...info, contentReturned: false, reason: `Larger than ${maxBytes} bytes; only metadata is returned. Open it in Outlook.` };
  }
  const full = (await withThrottleRetry(() => client.api(path).get())) as GraphAttachment;
  const bytes = Buffer.from(full.contentBytes ?? "", "base64");
  if (bytes.length > maxBytes) {
    return { ...info, contentReturned: false, reason: `Larger than ${maxBytes} bytes; only metadata is returned. Open it in Outlook.` };
  }
  const decoded = bytes.toString("utf8").replace(/^\uFEFF/, "");
  const { text, truncated } = truncateText(decoded, maxChars);
  return { ...info, contentReturned: true, text, truncated };
}

// ---------------------------------------------------------------- drafts

export interface MessageContentInput {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  bodyText?: string;
  bodyHtml?: string;
  importance?: Importance;
}

function recipients(list: string[]): GraphRecipient[] {
  return list.map((address) => ({ emailAddress: { address: address.trim() } }));
}

/** Graph message payload with only the fields that were given. */
export function buildMessagePayload(input: MessageContentInput): Record<string, unknown> {
  if (input.bodyText !== undefined && input.bodyHtml !== undefined) {
    throw new Error("Pass either body_text or body_html, not both.");
  }
  const payload: Record<string, unknown> = {};
  if (input.subject !== undefined) payload.subject = input.subject;
  if (input.bodyText !== undefined) payload.body = { contentType: "text", content: input.bodyText };
  if (input.bodyHtml !== undefined) payload.body = { contentType: "html", content: input.bodyHtml };
  if (input.to !== undefined) payload.toRecipients = recipients(input.to);
  if (input.cc !== undefined) payload.ccRecipients = recipients(input.cc);
  if (input.bcc !== undefined) payload.bccRecipients = recipients(input.bcc);
  if (input.importance !== undefined) payload.importance = input.importance;
  return payload;
}

export async function createDraft(client: Client, input: MessageContentInput): Promise<DraftSummary> {
  const created = (await withThrottleRetry(() => client.api("/me/messages").post(buildMessagePayload(input)))) as GraphMessage;
  return toDraftSummary(created);
}

async function requireDraft(client: Client, messageId: string): Promise<GraphMessage> {
  const m = (await withThrottleRetry(() => client.api(messagePath(messageId)).select(DRAFT_SELECT).get())) as GraphMessage;
  if (m.isDraft !== true) throw new Error("This message is not a draft; only drafts can be edited or sent.");
  return m;
}

export async function updateDraft(client: Client, messageId: string, input: MessageContentInput): Promise<DraftSummary> {
  const payload = buildMessagePayload(input);
  if (Object.keys(payload).length === 0) throw new Error("Nothing to update: pass at least one field.");
  await requireDraft(client, messageId);
  const updated = (await withThrottleRetry(() => client.api(messagePath(messageId)).patch(payload))) as GraphMessage | undefined;
  return toDraftSummary(updated?.id ? updated : await requireDraft(client, messageId));
}

export async function createReplyDraft(client: Client, messageId: string, opts: { replyAll?: boolean; comment?: string }): Promise<DraftSummary> {
  const action = opts.replyAll ? "createReplyAll" : "createReply";
  const body = opts.comment ? { comment: opts.comment } : {};
  const draft = (await withThrottleRetry(() => client.api(`${messagePath(messageId)}/${action}`).post(body))) as GraphMessage;
  return toDraftSummary(draft);
}

export async function createForwardDraft(client: Client, messageId: string, opts: { to: string[]; comment?: string }): Promise<DraftSummary> {
  const body: Record<string, unknown> = { toRecipients: recipients(opts.to) };
  if (opts.comment) body.comment = opts.comment;
  const draft = (await withThrottleRetry(() => client.api(`${messagePath(messageId)}/createForward`).post(body))) as GraphMessage;
  return toDraftSummary(draft);
}

// ---------------------------------------------------------------- organize

export async function moveMessage(client: Client, messageId: string, destination: string): Promise<{ id: string; parentFolderId?: string }> {
  // Graph accepts a folder id or a well-known name as destinationId; the moved message gets a new id.
  const moved = (await withThrottleRetry(() => client.api(`${messagePath(messageId)}/move`).post({ destinationId: destination }))) as GraphMessage;
  return { id: moved.id, parentFolderId: moved.parentFolderId };
}

/** Soft delete only: the message goes to Deleted Items, where the user can restore it. */
export async function softDeleteMessage(client: Client, messageId: string): Promise<{ deleted: true; movedTo: "deleteditems"; id: string }> {
  const moved = await moveMessage(client, messageId, "deleteditems");
  return { deleted: true, movedTo: "deleteditems", id: moved.id };
}

export interface MessageUpdateInput {
  isRead?: boolean;
  flagStatus?: FlagStatus;
  categories?: string[];
  importance?: Importance;
}

export async function updateMessage(client: Client, messageId: string, input: MessageUpdateInput): Promise<MessageSummary> {
  const payload: Record<string, unknown> = {};
  if (input.isRead !== undefined) payload.isRead = input.isRead;
  if (input.flagStatus !== undefined) payload.flag = { flagStatus: input.flagStatus };
  if (input.categories !== undefined) payload.categories = input.categories;
  if (input.importance !== undefined) payload.importance = input.importance;
  if (Object.keys(payload).length === 0) throw new Error("Nothing to update: pass at least one of is_read, flag_status, categories, importance.");
  await withThrottleRetry(() => client.api(messagePath(messageId)).patch(payload));
  const m = (await withThrottleRetry(() => client.api(messagePath(messageId)).select(SUMMARY_SELECT).get())) as GraphMessage;
  return toMessageSummary(m);
}

// ---------------------------------------------------------------- send

export interface SendResult {
  sent: true;
  recipients: number;
  draftId?: string;
}

export async function sendDraft(client: Client, messageId: string): Promise<SendResult> {
  const draft = await requireDraft(client, messageId);
  const count = recipientCount(draft);
  if (count === 0) throw new Error("The draft has no recipients.");
  await withThrottleRetry(() => client.api(`${messagePath(messageId)}/send`).post({}));
  return { sent: true, recipients: count, draftId: messageId };
}

export async function sendMail(client: Client, input: MessageContentInput): Promise<SendResult> {
  const message = buildMessagePayload(input);
  const count = recipientCount(message as Pick<GraphMessage, "toRecipients" | "ccRecipients" | "bccRecipients">);
  if (count === 0) throw new Error("Pass at least one recipient in to, cc or bcc.");
  await withThrottleRetry(() => client.api("/me/sendMail").post({ message, saveToSentItems: true }));
  return { sent: true, recipients: count };
}

/**
 * Reply/forward and send. Done as create-draft-then-send (not /reply) so the
 * exact recipient count Graph resolved is known for the audit log. If the
 * send step fails, the draft stays in Drafts and its id is in the error.
 */
async function sendCreatedDraft(client: Client, draft: GraphMessage): Promise<SendResult> {
  const count = recipientCount(draft);
  if (count === 0) throw new Error(`The reply has no recipients; the draft ${draft.id} was left in Drafts.`);
  try {
    await withThrottleRetry(() => client.api(`${messagePath(draft.id)}/send`).post({}));
  } catch (err) {
    if (err instanceof Error) err.message = `${err.message} (the unsent draft ${draft.id} is in Drafts)`;
    throw err;
  }
  return { sent: true, recipients: count, draftId: draft.id };
}

export async function replyAndSend(client: Client, messageId: string, opts: { replyAll?: boolean; comment: string }): Promise<SendResult> {
  const action = opts.replyAll ? "createReplyAll" : "createReply";
  const draft = (await withThrottleRetry(() => client.api(`${messagePath(messageId)}/${action}`).post({ comment: opts.comment }))) as GraphMessage;
  return sendCreatedDraft(client, draft);
}

export async function forwardAndSend(client: Client, messageId: string, opts: { to: string[]; comment?: string }): Promise<SendResult> {
  const body: Record<string, unknown> = { toRecipients: recipients(opts.to) };
  if (opts.comment) body.comment = opts.comment;
  const draft = (await withThrottleRetry(() => client.api(`${messagePath(messageId)}/createForward`).post(body))) as GraphMessage;
  return sendCreatedDraft(client, draft);
}
