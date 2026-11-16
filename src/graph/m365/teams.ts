import type { Client } from "@microsoft/microsoft-graph-client";
import { htmlToText, nextPageRequest, pageOf, truncateText, type Page } from "../../tools/m365/common.js";

/**
 * Microsoft Teams through Graph v1.0, delegated (work/school accounts only).
 * Endpoint notes (verified against Microsoft Learn):
 *
 * - GET /me/joinedTeams (Team.ReadBasic.All) supports no OData query options.
 * - GET /teams/{id}/channels (Channel.ReadBasic.All) supports $select/$filter, no $top.
 * - GET /me/chats (Chat.ReadWrite) supports $expand=members,lastMessagePreview,
 *   $orderby=lastMessagePreview/createdDateTime desc, $top <= 50, no $select.
 *   Expanded members are capped at 25 per chat by Graph.
 * - GET /me/chats/{id}/messages (Chat.ReadWrite) supports $top <= 50 and
 *   $orderby=createdDateTime desc / lastModifiedDateTime desc only, no $select.
 * - GET .../messages/{id} (chat: Chat.ReadWrite, channel: ChannelMessage.Read.All)
 *   supports no OData query options.
 * - GET /teams/{t}/channels/{c}/messages and .../replies (ChannelMessage.Read.All,
 *   admin consent) support $top <= 50 only (no $orderby/$select); top-level posts are
 *   sorted by last activity in the thread.
 * - POST /chats/{id}/messages (ChatMessage.Send), POST .../channels/{c}/messages and
 *   .../replies (ChannelMessage.Send).
 * - POST /chats (least privileged Chat.Create; Chat.ReadWrite is accepted as the
 *   higher-privileged alternative, so the `teams` module covers it).
 */

export const MESSAGE_PAGE_MAX = 50;
/** Characters of message text returned per message in list results. */
export const LIST_TEXT_LIMIT = 2000;
/** Characters of message text returned for a single message. */
export const GET_TEXT_LIMIT = 20000;

interface GraphIdentity {
  id?: string | null;
  displayName?: string | null;
}

interface GraphIdentitySet {
  user?: GraphIdentity | null;
  application?: GraphIdentity | null;
  device?: GraphIdentity | null;
}

export interface GraphChatMessage {
  id: string;
  replyToId?: string | null;
  messageType?: string;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  lastEditedDateTime?: string | null;
  deletedDateTime?: string | null;
  subject?: string | null;
  importance?: string;
  webUrl?: string | null;
  from?: GraphIdentitySet | null;
  body?: { contentType?: string; content?: string | null } | null;
  attachments?: Array<{ id?: string | null; name?: string | null; contentType?: string | null; contentUrl?: string | null }>;
  mentions?: Array<{ id?: number; mentionText?: string | null }>;
  reactions?: unknown[];
  eventDetail?: { "@odata.type"?: string } | null;
  replies?: GraphChatMessage[];
}

export interface ShapedMessage {
  id: string;
  messageType: string;
  from: string | null;
  createdDateTime: string | null;
  lastEditedDateTime?: string;
  deleted?: true;
  subject?: string;
  importance?: string;
  replyToId?: string;
  text: string;
  truncated?: true;
  attachments?: Array<{ name: string | null; contentType: string | null; url?: string }>;
  mentions?: string[];
  reactionCount?: number;
  systemEvent?: string;
  webUrl?: string;
}

function senderName(from: GraphIdentitySet | null | undefined): string | null {
  if (!from) return null;
  return from.user?.displayName ?? from.application?.displayName ?? from.device?.displayName ?? null;
}

function bodyText(body: GraphChatMessage["body"]): string {
  const content = body?.content ?? "";
  if (!content) return "";
  return body?.contentType === "html" ? htmlToText(content) : content.trim();
}

/**
 * Reduces a Graph chatMessage to what a model needs: sender display name,
 * timestamps, plain text (html converted, then capped), and short
 * summaries of attachments, mentions and system events.
 */
export function shapeMessage(m: GraphChatMessage, maxChars: number): ShapedMessage {
  const isSystem = m.messageType !== undefined && m.messageType !== "message";
  const { text, truncated } = truncateText(isSystem ? "" : bodyText(m.body), maxChars);
  const out: ShapedMessage = {
    id: m.id,
    messageType: m.messageType ?? "message",
    from: senderName(m.from),
    createdDateTime: m.createdDateTime ?? null,
    text,
  };
  if (truncated) out.truncated = true;
  if (m.lastEditedDateTime) out.lastEditedDateTime = m.lastEditedDateTime;
  if (m.deletedDateTime) out.deleted = true;
  if (m.subject) out.subject = m.subject;
  if (m.importance && m.importance !== "normal") out.importance = m.importance;
  if (m.replyToId) out.replyToId = m.replyToId;
  if (m.attachments?.length) {
    out.attachments = m.attachments.map((a) => {
      const shaped: { name: string | null; contentType: string | null; url?: string } = { name: a.name ?? null, contentType: a.contentType ?? null };
      // "reference" attachments are shared files; the link lets the files tools find them.
      if (a.contentType === "reference" && a.contentUrl) shaped.url = a.contentUrl;
      return shaped;
    });
  }
  if (m.mentions?.length) out.mentions = m.mentions.map((x) => x.mentionText ?? "").filter((x) => x !== "");
  if (m.reactions?.length) out.reactionCount = m.reactions.length;
  if (isSystem && m.eventDetail?.["@odata.type"]) out.systemEvent = m.eventDetail["@odata.type"].replace(/^#microsoft\.graph\./, "");
  if (m.webUrl) out.webUrl = m.webUrl;
  return out;
}

const seg = (id: string) => encodeURIComponent(id);

function clampTop(pageSize: number | undefined, max: number, fallback = 25): number {
  const n = pageSize ?? fallback;
  return Math.max(1, Math.min(max, Math.floor(n)));
}

function shapePage(response: { value?: GraphChatMessage[]; "@odata.nextLink"?: string }, maxChars: number): Page<ShapedMessage> {
  const page = pageOf(response);
  return { ...page, items: page.items.map((m) => shapeMessage(m, maxChars)) };
}

export interface PageOptions {
  pageSize?: number;
  nextPageToken?: string;
}

// ---- Teams and channels (module "teams") ----

export async function listJoinedTeams(client: Client, opts: { nextPageToken?: string } = {}) {
  const req = opts.nextPageToken ? nextPageRequest(client, opts.nextPageToken) : client.api("/me/joinedTeams");
  const res = (await req.get()) as { value?: Array<{ id: string; displayName?: string; description?: string | null; isArchived?: boolean }>; "@odata.nextLink"?: string };
  const page = pageOf(res);
  return {
    ...page,
    items: page.items.map((t) => ({ id: t.id, displayName: t.displayName ?? null, description: t.description ?? null, isArchived: t.isArchived ?? false })),
  };
}

export async function listChannels(client: Client, teamId: string) {
  const res = (await client
    .api(`/teams/${seg(teamId)}/channels`)
    .select("id,displayName,description,membershipType,webUrl")
    .get()) as { value?: Array<{ id: string; displayName?: string; description?: string | null; membershipType?: string; webUrl?: string }>; "@odata.nextLink"?: string };
  const page = pageOf(res);
  return {
    ...page,
    items: page.items.map((c) => ({
      id: c.id,
      displayName: c.displayName ?? null,
      description: c.description ?? null,
      membershipType: c.membershipType ?? null,
      webUrl: c.webUrl ?? null,
    })),
  };
}

// ---- Chats (module "teams") ----

interface GraphChat {
  id: string;
  topic?: string | null;
  chatType?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
  viewpoint?: { isHidden?: boolean; lastMessageReadDateTime?: string } | null;
  members?: Array<{ displayName?: string | null }>;
  lastMessagePreview?: {
    id?: string;
    createdDateTime?: string;
    isDeleted?: boolean;
    messageType?: string;
    from?: GraphIdentitySet | null;
    body?: { contentType?: string; content?: string | null } | null;
  } | null;
}

const PREVIEW_CHARS = 200;

export function shapeChat(c: GraphChat) {
  const preview = c.lastMessagePreview;
  let lastMessage: { from: string | null; createdDateTime: string | null; text: string; unread?: true } | null = null;
  if (preview && !preview.isDeleted) {
    const isSystem = preview.messageType !== undefined && preview.messageType !== "message";
    lastMessage = {
      from: senderName(preview.from),
      createdDateTime: preview.createdDateTime ?? null,
      text: isSystem ? "" : truncateText(bodyText(preview.body), PREVIEW_CHARS).text,
    };
    const read = c.viewpoint?.lastMessageReadDateTime;
    if (preview.createdDateTime && read && preview.createdDateTime > read) lastMessage.unread = true;
  }
  return {
    id: c.id,
    chatType: c.chatType ?? null,
    topic: c.topic ?? null,
    members: (c.members ?? []).map((m) => m.displayName ?? "").filter((n) => n !== ""),
    lastUpdatedDateTime: c.lastUpdatedDateTime ?? null,
    lastMessage,
    webUrl: c.webUrl ?? null,
  };
}

export async function listChats(client: Client, opts: PageOptions = {}) {
  type ChatPage = { value?: GraphChat[]; "@odata.nextLink"?: string };
  const firstPage = (expand: string) =>
    client
      .api("/me/chats")
      .expand(expand)
      .orderby("lastMessagePreview/createdDateTime desc")
      .top(clampTop(opts.pageSize, 50));
  let res: ChatPage;
  if (opts.nextPageToken) {
    res = (await nextPageRequest(client, opts.nextPageToken).get()) as ChatPage;
  } else {
    // Graph documents expanding members and lastMessagePreview separately;
    // if the combination is rejected, fall back to the preview (which the
    // ordering needs) and leave member names out rather than failing.
    try {
      res = (await firstPage("members,lastMessagePreview").get()) as ChatPage;
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode !== 400) throw err;
      res = (await firstPage("lastMessagePreview").get()) as ChatPage;
    }
  }
  const page = pageOf(res);
  return { ...page, items: page.items.map(shapeChat) };
}

export async function listChatMessages(client: Client, chatId: string, opts: PageOptions = {}) {
  const req = opts.nextPageToken
    ? nextPageRequest(client, opts.nextPageToken)
    : client.api(`/me/chats/${seg(chatId)}/messages`).orderby("createdDateTime desc").top(clampTop(opts.pageSize, MESSAGE_PAGE_MAX));
  return shapePage((await req.get()) as { value?: GraphChatMessage[] }, LIST_TEXT_LIMIT);
}

export async function getChatMessage(client: Client, chatId: string, messageId: string) {
  const m = (await client.api(`/me/chats/${seg(chatId)}/messages/${seg(messageId)}`).get()) as GraphChatMessage;
  return shapeMessage(m, GET_TEXT_LIMIT);
}

// ---- Channel messages (module "teams_channels") ----

export async function listChannelMessages(client: Client, teamId: string, channelId: string, opts: PageOptions = {}) {
  const req = opts.nextPageToken
    ? nextPageRequest(client, opts.nextPageToken)
    : client.api(`/teams/${seg(teamId)}/channels/${seg(channelId)}/messages`).top(clampTop(opts.pageSize, MESSAGE_PAGE_MAX));
  return shapePage((await req.get()) as { value?: GraphChatMessage[] }, LIST_TEXT_LIMIT);
}

export async function listChannelMessageReplies(client: Client, teamId: string, channelId: string, messageId: string, opts: PageOptions = {}) {
  const req = opts.nextPageToken
    ? nextPageRequest(client, opts.nextPageToken)
    : client
        .api(`/teams/${seg(teamId)}/channels/${seg(channelId)}/messages/${seg(messageId)}/replies`)
        .top(clampTop(opts.pageSize, MESSAGE_PAGE_MAX));
  return shapePage((await req.get()) as { value?: GraphChatMessage[] }, LIST_TEXT_LIMIT);
}

export async function getChannelMessage(client: Client, teamId: string, channelId: string, messageId: string, replyId?: string) {
  const base = `/teams/${seg(teamId)}/channels/${seg(channelId)}/messages/${seg(messageId)}`;
  const m = (await client.api(replyId ? `${base}/replies/${seg(replyId)}` : base).get()) as GraphChatMessage;
  return shapeMessage(m, GET_TEXT_LIMIT);
}

// ---- Sending (module "teams", only registered when the profile allows sending) ----

export type BodyType = "text" | "html";

export interface OutgoingMessage {
  content: string;
  contentType?: BodyType;
  subject?: string;
}

function messagePayload(msg: OutgoingMessage): Record<string, unknown> {
  const payload: Record<string, unknown> = { body: { contentType: msg.contentType ?? "text", content: msg.content } };
  if (msg.subject) payload.subject = msg.subject;
  return payload;
}

function sentResult(m: GraphChatMessage) {
  return { sent: true, id: m.id, createdDateTime: m.createdDateTime ?? null, webUrl: m.webUrl ?? null };
}

export async function sendChatMessage(client: Client, chatId: string, msg: OutgoingMessage) {
  const m = (await client.api(`/chats/${seg(chatId)}/messages`).post(messagePayload({ content: msg.content, contentType: msg.contentType }))) as GraphChatMessage;
  return sentResult(m);
}

export async function sendChannelMessage(client: Client, teamId: string, channelId: string, msg: OutgoingMessage) {
  const m = (await client.api(`/teams/${seg(teamId)}/channels/${seg(channelId)}/messages`).post(messagePayload(msg))) as GraphChatMessage;
  return sentResult(m);
}

export async function replyToChannelMessage(client: Client, teamId: string, channelId: string, messageId: string, msg: OutgoingMessage) {
  const m = (await client
    .api(`/teams/${seg(teamId)}/channels/${seg(channelId)}/messages/${seg(messageId)}/replies`)
    .post(messagePayload({ content: msg.content, contentType: msg.contentType }))) as GraphChatMessage;
  return sentResult(m);
}

const GRAPH_USERS = "https://graph.microsoft.com/v1.0/users";

/** A user id (GUID) or user principal name, as accepted inside users('...'). */
export const USER_REF_PATTERN = /^[^\s/\\?#"'%]+$/;

function memberBind(userRef: string) {
  if (!USER_REF_PATTERN.test(userRef)) throw new Error(`Invalid member "${userRef}": pass a user id or user principal name.`);
  return { "@odata.type": "#microsoft.graph.aadUserConversationMember", roles: ["owner"], "user@odata.bind": `${GRAPH_USERS}('${userRef}')` };
}

export interface NewChat {
  chatType: "oneOnOne" | "group";
  /** The other participants (user ids or UPNs); the signed-in user is added automatically. */
  members: string[];
  topic?: string;
}

/**
 * Creates a oneOnOne or group chat (or returns the existing oneOnOne chat,
 * which Graph does for a pair that already has one). Graph requires the
 * signed-in user in the member list, so their id is looked up first.
 */
export async function createChat(client: Client, chat: NewChat) {
  const others = [...new Set(chat.members.map((m) => m.trim()).filter((m) => m !== ""))];
  if (chat.chatType === "oneOnOne" && others.length !== 1) throw new Error("A oneOnOne chat needs exactly one other member.");
  if (chat.chatType === "group" && others.length < 2) throw new Error("A group chat needs at least two other members.");
  if (chat.topic && chat.chatType !== "group") throw new Error("Only group chats can have a topic.");
  const binds = others.map(memberBind);
  const me = (await client.api("/me").select("id").get()) as { id: string };
  const body: Record<string, unknown> = { chatType: chat.chatType, members: [memberBind(me.id), ...binds] };
  if (chat.topic) body.topic = chat.topic;
  const created = (await client.api("/chats").post(body)) as GraphChat;
  return { id: created.id, chatType: created.chatType ?? chat.chatType, topic: created.topic ?? null, webUrl: created.webUrl ?? null, memberCount: others.length + 1 };
}
