import type { Client } from "@microsoft/microsoft-graph-client";
import { htmlToText, nextPageRequest, pageOf, truncateText, type Page } from "../../tools/m365/common.js";

/**
 * Outlook calendar on Microsoft Graph v1.0 (delegated, /me). Pure functions
 * over a Graph client; the tool layer resolves the default time zone,
 * `allowSend` and auditing.
 *
 * Sending rule: anything that emails other people counts as sending —
 * creating or updating a meeting that has attendees (invitations/updates),
 * cancelling a meeting you organize, and replying to an invitation with
 * sendResponse. Those paths throw SendNotAllowedError unless `allowSend`.
 */

export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;

export class SendNotAllowedError extends Error {
  constructor(what: string) {
    super(
      `${what} would email other people, and sending is not enabled for this profile (allowSend is off). ` +
        "Do it without attendees / without notifying anyone, or ask the owner to enable sending for this profile."
    );
    this.name = "SendNotAllowedError";
  }
}

// ---------------------------------------------------------------------------
// Time zones
// ---------------------------------------------------------------------------

/**
 * IANA names mapped to the Windows names Outlook knows for sure. Graph only
 * accepts a subset of IANA names in event payloads and the Prefer header
 * (Europe/Copenhagen is not among them), so known zones are sent as Windows
 * names. The first IANA name per Windows zone is used for the reverse lookup.
 */
const IANA_TO_WINDOWS: Record<string, string> = {
  "Europe/Copenhagen": "Romance Standard Time",
  "Europe/Paris": "Romance Standard Time",
  "Europe/Brussels": "Romance Standard Time",
  "Europe/Madrid": "Romance Standard Time",
  "Europe/Berlin": "W. Europe Standard Time",
  "Europe/Oslo": "W. Europe Standard Time",
  "Europe/Stockholm": "W. Europe Standard Time",
  "Europe/Amsterdam": "W. Europe Standard Time",
  "Europe/Rome": "W. Europe Standard Time",
  "Europe/Vienna": "W. Europe Standard Time",
  "Europe/Zurich": "W. Europe Standard Time",
  "Europe/Warsaw": "Central European Standard Time",
  "Europe/Prague": "Central Europe Standard Time",
  "Europe/Budapest": "Central Europe Standard Time",
  "Europe/London": "GMT Standard Time",
  "Europe/Dublin": "GMT Standard Time",
  "Europe/Lisbon": "GMT Standard Time",
  "Atlantic/Reykjavik": "Greenwich Standard Time",
  "Europe/Helsinki": "FLE Standard Time",
  "Europe/Tallinn": "FLE Standard Time",
  "Europe/Riga": "FLE Standard Time",
  "Europe/Vilnius": "FLE Standard Time",
  "Europe/Athens": "GTB Standard Time",
  "Europe/Bucharest": "GTB Standard Time",
  "America/New_York": "Eastern Standard Time",
  "America/Chicago": "Central Standard Time",
  "America/Denver": "Mountain Standard Time",
  "America/Los_Angeles": "Pacific Standard Time",
  "Asia/Tokyo": "Tokyo Standard Time",
  "Asia/Shanghai": "China Standard Time",
  "Asia/Kolkata": "India Standard Time",
  "Australia/Sydney": "AUS Eastern Standard Time",
  UTC: "UTC",
  "Etc/UTC": "UTC",
};

const WINDOWS_TO_IANA: Record<string, string> = {};
for (const [iana, windows] of Object.entries(IANA_TO_WINDOWS)) WINDOWS_TO_IANA[windows] ??= iana;

export interface ResolvedTimeZone {
  /** Name sent to Graph (payloads and the Prefer header). */
  graph: string;
  /** IANA name for local wall-clock arithmetic, when known. */
  iana?: string;
}

function isIanaZone(name: string): boolean {
  if (!name.includes("/") && name !== "UTC") return false;
  try {
    new Intl.DateTimeFormat("en-US", { timeZone: name });
    return true;
  } catch {
    return false;
  }
}

export function resolveTimeZone(timeZone: string): ResolvedTimeZone {
  const name = timeZone.trim();
  const windows = IANA_TO_WINDOWS[name];
  if (windows) return { graph: windows, iana: name };
  const iana = WINDOWS_TO_IANA[name];
  if (iana) return { graph: name, iana };
  if (isIanaZone(name)) return { graph: name, iana: name };
  return { graph: name };
}

export function preferTimeZone(tz: ResolvedTimeZone): string {
  return `outlook.timezone="${tz.graph.replace(/"/g, "")}"`;
}

interface ParsedDateTime {
  date: string; // YYYY-MM-DD
  time?: string; // HH:mm:ss
  offsetMinutes?: number;
}

const DATE_TIME_RE = /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

export function parseDateTimeInput(value: string): ParsedDateTime {
  const m = DATE_TIME_RE.exec(value.trim());
  if (!m) throw new Error(`Invalid date/time "${value}": use YYYY-MM-DD or YYYY-MM-DDTHH:mm (optionally with seconds and a UTC offset).`);
  const [, date, hh, mm, ss, offset] = m;
  if (Number.isNaN(Date.parse(`${date}T00:00:00Z`)) || new Date(`${date}T00:00:00Z`).toISOString().slice(0, 10) !== date) {
    throw new Error(`Invalid date "${value}".`);
  }
  if (hh === undefined) return { date: date! };
  if (Number(hh) > 23 || Number(mm) > 59 || Number(ss ?? 0) > 59) throw new Error(`Invalid time "${value}".`);
  const result: ParsedDateTime = { date: date!, time: `${hh}:${mm}:${ss ?? "00"}` };
  if (offset) {
    if (offset === "Z") result.offsetMinutes = 0;
    else {
      const sign = offset.startsWith("-") ? -1 : 1;
      const digits = offset.slice(1).replace(":", "");
      result.offsetMinutes = sign * (Number(digits.slice(0, 2)) * 60 + Number(digits.slice(2, 4)));
    }
  }
  return result;
}

export function addDays(date: string, days: number): string {
  return new Date(Date.parse(`${date}T00:00:00Z`) + days * 86_400_000).toISOString().slice(0, 10);
}

/** Offset (ms) of `iana` from UTC at the instant `utcMs`. */
function zoneOffsetMs(utcMs: number, iana: string): number {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: iana,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).formatToParts(new Date(utcMs));
  const get = (type: string) => Number(parts.find((p) => p.type === type)?.value);
  const asUtc = Date.UTC(get("year"), get("month") - 1, get("day"), get("hour") % 24, get("minute"), get("second"));
  return asUtc - Math.floor(utcMs / 1000) * 1000;
}

/** UTC instant for a wall-clock time in `iana` (wall clock given as "YYYY-MM-DDTHH:mm:ss"). */
function wallClockToUtcMs(wall: string, iana: string): number {
  const wallMs = Date.parse(`${wall}Z`);
  const first = zoneOffsetMs(wallMs, iana);
  const second = zoneOffsetMs(wallMs - first, iana);
  return wallMs - (first === second ? first : second);
}

function utcMsToWallClock(utcMs: number, iana: string): string {
  return new Date(utcMs + zoneOffsetMs(utcMs, iana)).toISOString().slice(0, 19);
}

function instantOf(p: ParsedDateTime): number {
  return Date.parse(`${p.date}T${p.time ?? "00:00:00"}Z`) - (p.offsetMinutes ?? 0) * 60_000;
}

export interface GraphDateTime {
  dateTime: string;
  timeZone: string;
}

/**
 * A timed value for an event payload. Without an offset the value is local
 * wall-clock time in `tz`; with an offset it is converted to wall-clock time
 * in `tz` (or sent as UTC when `tz` is not a known IANA zone).
 */
export function toEventDateTime(value: string, tz: ResolvedTimeZone): GraphDateTime {
  const p = parseDateTimeInput(value);
  if (!p.time) {
    throw new Error(`"${value}" is a date without a time: set all_day=true for all-day events, or pass YYYY-MM-DDTHH:mm.`);
  }
  if (p.offsetMinutes === undefined) return { dateTime: `${p.date}T${p.time}`, timeZone: tz.graph };
  const utcMs = instantOf(p);
  if (tz.iana) return { dateTime: utcMsToWallClock(utcMs, tz.iana), timeZone: tz.graph };
  return { dateTime: new Date(utcMs).toISOString().slice(0, 19), timeZone: "UTC" };
}

/**
 * A UTC instant (ISO, with Z) for calendarView ranges, which Graph reads as
 * UTC when no offset is given. A date alone means local midnight; as an end
 * it means the end of that day (midnight of the next day).
 */
export function toRangeInstant(value: string, tz: ResolvedTimeZone, isEnd: boolean): string {
  const p = parseDateTimeInput(value);
  if (p.offsetMinutes !== undefined) return new Date(instantOf(p)).toISOString();
  const date = !p.time && isEnd ? addDays(p.date, 1) : p.date;
  const wall = `${date}T${p.time ?? "00:00:00"}`;
  if (!tz.iana) {
    throw new Error(`Cannot convert "${value}" from time zone "${tz.graph}": pass a UTC offset (e.g. ${wall}+01:00) or an IANA time_zone.`);
  }
  return new Date(wallClockToUtcMs(wall, tz.iana)).toISOString();
}

/** A window boundary for getSchedule/findMeetingTimes; a date alone means local midnight (end: end of that day). */
function toWindowDateTime(value: string, tz: ResolvedTimeZone, isEnd: boolean): GraphDateTime {
  const p = parseDateTimeInput(value);
  if (!p.time) return { dateTime: `${isEnd ? addDays(p.date, 1) : p.date}T00:00:00`, timeZone: tz.graph };
  return toEventDateTime(value, tz);
}

function windowMs(start: string, end: string, tz: ResolvedTimeZone): number | undefined {
  try {
    return Date.parse(toRangeInstant(end, tz, true)) - Date.parse(toRangeInstant(start, tz, false));
  } catch {
    return undefined;
  }
}

// ---------------------------------------------------------------------------
// Output shaping
// ---------------------------------------------------------------------------

type GraphObject = Record<string, unknown>;
interface GraphEmail {
  name?: string;
  address?: string;
}
interface GraphAttendee {
  type?: string;
  emailAddress?: GraphEmail;
  status?: { response?: string; time?: string };
}

/** "2026-11-16T09:00:00.0000000" -> "2026-11-16T09:00". */
function shortDateTime(value: unknown): string | undefined {
  const dt = (value as GraphDateTime | undefined)?.dateTime;
  return typeof dt === "string" ? dt.slice(0, 16) : undefined;
}

function shapeTimes(e: GraphObject): { start?: string; end?: string } {
  const start = shortDateTime(e.start);
  const end = shortDateTime(e.end);
  if (e.isAllDay === true && start && end) {
    // Graph's all-day end is exclusive midnight; report the inclusive last day.
    const lastDay = end.endsWith("T00:00") ? addDays(end.slice(0, 10), -1) : end.slice(0, 10);
    return { start: start.slice(0, 10), end: lastDay < start.slice(0, 10) ? start.slice(0, 10) : lastDay };
  }
  return { start, end };
}

function email(e: GraphEmail | undefined): { name?: string; address?: string } | undefined {
  if (!e) return undefined;
  return { name: e.name, address: e.address };
}

export function shapeEventSummary(e: GraphObject) {
  const location = (e.location as { displayName?: string } | undefined)?.displayName;
  return {
    id: e.id as string,
    subject: e.subject as string | undefined,
    ...shapeTimes(e),
    is_all_day: e.isAllDay === true,
    location: location || undefined,
    organizer: email((e.organizer as { emailAddress?: GraphEmail } | undefined)?.emailAddress),
    is_organizer: e.isOrganizer as boolean | undefined,
    my_response: (e.responseStatus as { response?: string } | undefined)?.response,
    show_as: e.showAs as string | undefined,
    is_cancelled: e.isCancelled === true ? true : undefined,
    is_online_meeting: e.isOnlineMeeting === true ? true : undefined,
    join_url: (e.onlineMeeting as { joinUrl?: string } | undefined)?.joinUrl ?? undefined,
    type: e.type as string | undefined,
    categories: Array.isArray(e.categories) && e.categories.length ? (e.categories as string[]) : undefined,
  };
}

function shapeAttendees(attendees: unknown) {
  return ((attendees as GraphAttendee[] | undefined) ?? []).map((a) => ({
    name: a.emailAddress?.name,
    address: a.emailAddress?.address,
    type: a.type,
    response: a.status?.response,
  }));
}

export function shapeEventDetail(e: GraphObject, maxBodyChars: number) {
  const body = e.body as { contentType?: string; content?: string } | undefined;
  const rawText = body?.content ? (body.contentType === "html" ? htmlToText(body.content) : body.content.trim()) : "";
  const { text, truncated } = truncateText(rawText, maxBodyChars);
  return {
    ...shapeEventSummary(e),
    time_zone: (e.start as GraphDateTime | undefined)?.timeZone,
    body: text || undefined,
    body_truncated: truncated || undefined,
    attendees: shapeAttendees(e.attendees),
    online_meeting_provider: e.onlineMeetingProvider as string | undefined,
    reminder_minutes: e.isReminderOn === false ? null : (e.reminderMinutesBeforeStart as number | undefined),
    importance: e.importance as string | undefined,
    sensitivity: e.sensitivity as string | undefined,
    response_requested: e.responseRequested as boolean | undefined,
    recurrence: e.recurrence ?? undefined,
    series_master_id: e.seriesMasterId as string | undefined,
    web_link: e.webLink as string | undefined,
  };
}

const EVENT_SUMMARY_SELECT = [
  "id",
  "subject",
  "start",
  "end",
  "isAllDay",
  "location",
  "organizer",
  "isOrganizer",
  "responseStatus",
  "showAs",
  "isCancelled",
  "isOnlineMeeting",
  "onlineMeeting",
  "type",
  "categories",
].join(",");

const EVENT_DETAIL_SELECT = [
  EVENT_SUMMARY_SELECT,
  "body",
  "attendees",
  "onlineMeetingProvider",
  "isReminderOn",
  "reminderMinutesBeforeStart",
  "importance",
  "sensitivity",
  "responseRequested",
  "recurrence",
  "seriesMasterId",
  "webLink",
].join(",");

const enc = encodeURIComponent;

function clampTop(top: number | undefined): number {
  return Math.min(Math.max(Math.trunc(top ?? DEFAULT_PAGE_SIZE), 1), MAX_PAGE_SIZE);
}

// ---------------------------------------------------------------------------
// Reads
// ---------------------------------------------------------------------------

export async function listCalendars(client: Client, opts: { top?: number; pageToken?: string } = {}) {
  const request = opts.pageToken
    ? nextPageRequest(client, opts.pageToken)
    : client
        .api("/me/calendars")
        .select("id,name,color,isDefaultCalendar,canEdit,owner,defaultOnlineMeetingProvider,allowedOnlineMeetingProviders")
        .top(clampTop(opts.top ?? MAX_PAGE_SIZE));
  const page = pageOf<GraphObject>(await request.get());
  return {
    calendars: page.items.map((c) => ({
      id: c.id as string,
      name: c.name as string | undefined,
      is_default: c.isDefaultCalendar === true,
      can_edit: c.canEdit as boolean | undefined,
      owner: email(c.owner as GraphEmail | undefined),
      color: c.color as string | undefined,
      default_online_meeting_provider: c.defaultOnlineMeetingProvider as string | undefined,
      allowed_online_meeting_providers: c.allowedOnlineMeetingProviders as string[] | undefined,
    })),
    next_page_token: page.next_page_token,
  };
}

export interface ListEventsOptions {
  start: string;
  end: string;
  timeZone: string;
  calendarId?: string;
  top?: number;
  pageToken?: string;
}

/** Events (recurring occurrences expanded) overlapping [start, end) via calendarView. */
export async function listEvents(client: Client, opts: ListEventsOptions) {
  const tz = resolveTimeZone(opts.timeZone);
  let request;
  let range: { start: string; end: string } | undefined;
  if (opts.pageToken) {
    request = nextPageRequest(client, opts.pageToken);
  } else {
    range = { start: toRangeInstant(opts.start, tz, false), end: toRangeInstant(opts.end, tz, true) };
    if (Date.parse(range.end) <= Date.parse(range.start)) throw new Error("end must be after start.");
    const base = opts.calendarId ? `/me/calendars/${enc(opts.calendarId)}/calendarView` : "/me/calendarView";
    request = client
      .api(base)
      .query({ startDateTime: range.start, endDateTime: range.end })
      .select(EVENT_SUMMARY_SELECT)
      .orderby("start/dateTime")
      .top(clampTop(opts.top));
  }
  const page: Page<GraphObject> = pageOf(await request.header("Prefer", preferTimeZone(tz)).get());
  return {
    time_zone: tz.graph,
    range_utc: range,
    events: page.items.map(shapeEventSummary),
    next_page_token: page.next_page_token,
  };
}

export async function getEvent(client: Client, eventId: string, opts: { timeZone: string; maxBodyChars: number }) {
  const tz = resolveTimeZone(opts.timeZone);
  const event = (await client
    .api(`/me/events/${enc(eventId)}`)
    .select(EVENT_DETAIL_SELECT)
    .header("Prefer", preferTimeZone(tz))
    .get()) as GraphObject;
  return shapeEventDetail(event, opts.maxBodyChars);
}

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

export interface EventAttendeeInput {
  email: string;
  name?: string;
  type?: "required" | "optional" | "resource";
}

export type OnlineMeetingProvider = "teamsForBusiness" | "skypeForBusiness" | "skypeForConsumer";
export type ShowAs = "free" | "tentative" | "busy" | "oof" | "workingElsewhere";

export interface EventInput {
  subject?: string;
  /** Timed: YYYY-MM-DDTHH:mm (local in the time zone) or with offset. All-day: YYYY-MM-DD. */
  start?: string;
  /** Timed: exclusive end time. All-day: the inclusive last day (defaults to start). */
  end?: string;
  allDay?: boolean;
  location?: string | null;
  body?: string | null;
  attendees?: EventAttendeeInput[];
  onlineMeeting?: boolean;
  onlineMeetingProvider?: OnlineMeetingProvider;
  /** Minutes before start; null turns the reminder off. */
  reminderMinutes?: number | null;
  showAs?: ShowAs;
  importance?: "low" | "normal" | "high";
}

function timesPayload(input: EventInput, tz: ResolvedTimeZone, mode: "create" | "update"): GraphObject {
  if (input.allDay === true) {
    if (!input.start) throw new Error("All-day events need start (YYYY-MM-DD).");
    const start = parseDateTimeInput(input.start);
    const end = parseDateTimeInput(input.end ?? input.start);
    if (start.time || end.time) throw new Error("All-day events take dates only (YYYY-MM-DD), not times.");
    if (end.date < start.date) throw new Error("end (the last day) must not be before start.");
    return {
      isAllDay: true,
      start: { dateTime: `${start.date}T00:00:00`, timeZone: tz.graph },
      end: { dateTime: `${addDays(end.date, 1)}T00:00:00`, timeZone: tz.graph },
    };
  }
  const payload: GraphObject = {};
  if (input.allDay === false || mode === "create") {
    if (!input.start || !input.end) throw new Error("Timed events need both start and end (YYYY-MM-DDTHH:mm).");
    if (input.allDay === false) payload.isAllDay = false;
  }
  const start = input.start ? toEventDateTime(input.start, tz) : undefined;
  const end = input.end ? toEventDateTime(input.end, tz) : undefined;
  if (start && end && start.timeZone === end.timeZone && end.dateTime <= start.dateTime) {
    throw new Error("end must be after start.");
  }
  if (start) payload.start = start;
  if (end) payload.end = end;
  return payload;
}

/** Graph event JSON for create (`mode: "create"`) or a PATCH with only the given fields. */
export function buildEventPayload(input: EventInput, tz: ResolvedTimeZone, mode: "create" | "update"): GraphObject {
  if (mode === "create" && !input.subject) throw new Error("subject is required.");
  const payload: GraphObject = { ...timesPayload(input, tz, mode) };
  if (input.subject !== undefined) payload.subject = input.subject;
  if (input.location !== undefined) payload.location = { displayName: input.location ?? "" };
  if (input.body !== undefined && (input.body !== null || mode === "update")) {
    payload.body = { contentType: "text", content: input.body ?? "" };
  }
  if (input.attendees !== undefined) {
    payload.attendees = input.attendees.map((a) => ({
      emailAddress: a.name ? { address: a.email, name: a.name } : { address: a.email },
      type: a.type ?? "required",
    }));
  }
  if (input.onlineMeeting === true) {
    payload.isOnlineMeeting = true;
    payload.onlineMeetingProvider = input.onlineMeetingProvider ?? "teamsForBusiness";
  } else if (input.onlineMeeting === false) {
    payload.isOnlineMeeting = false;
  }
  if (input.reminderMinutes === null) payload.isReminderOn = false;
  else if (input.reminderMinutes !== undefined) {
    payload.isReminderOn = true;
    payload.reminderMinutesBeforeStart = input.reminderMinutes;
  }
  if (input.showAs !== undefined) payload.showAs = input.showAs;
  if (input.importance !== undefined) payload.importance = input.importance;
  return payload;
}

export async function createEvent(
  client: Client,
  input: EventInput,
  opts: { timeZone: string; allowSend: boolean; calendarId?: string }
) {
  const tz = resolveTimeZone(opts.timeZone);
  const attendees = input.attendees?.length ?? 0;
  if (attendees > 0 && !opts.allowSend) throw new SendNotAllowedError("Creating an event with attendees (it sends them invitations)");
  const payload = buildEventPayload(input, tz, "create");
  const path = opts.calendarId ? `/me/calendars/${enc(opts.calendarId)}/events` : "/me/events";
  const created = (await client.api(path).header("Prefer", preferTimeZone(tz)).post(payload)) as GraphObject;
  return { event: shapeEventSummary(created), time_zone: tz.graph, invitations_sent: attendees };
}

/** Fields that only affect the signed-in user's copy, so changing them never notifies attendees. */
const PERSONAL_FIELDS = new Set(["isReminderOn", "reminderMinutesBeforeStart"]);

interface ExistingEvent {
  isOrganizer?: boolean;
  isCancelled?: boolean;
  attendees?: GraphAttendee[];
}

async function getExisting(client: Client, eventId: string): Promise<ExistingEvent> {
  return (await client.api(`/me/events/${enc(eventId)}`).select("id,isOrganizer,isCancelled,attendees").get()) as ExistingEvent;
}

function addressSet(attendees: Array<{ emailAddress?: GraphEmail }> | undefined): Set<string> {
  return new Set((attendees ?? []).map((a) => (a.emailAddress?.address ?? "").toLowerCase()).filter(Boolean));
}

export async function updateEvent(
  client: Client,
  eventId: string,
  input: EventInput,
  opts: { timeZone: string; allowSend: boolean }
) {
  const tz = resolveTimeZone(opts.timeZone);
  const payload = buildEventPayload(input, tz, "update");
  if (Object.keys(payload).length === 0) throw new Error("Nothing to update: pass at least one field to change.");
  const existing = await getExisting(client, eventId);
  const before = addressSet(existing.attendees);
  const after = input.attendees ? new Set(input.attendees.map((a) => a.email.toLowerCase())) : before;
  const affected = new Set([...before, ...after]).size;
  const personalOnly = Object.keys(payload).every((k) => PERSONAL_FIELDS.has(k));
  // isOrganizer missing is treated as organizer, so the check errs on the side of "this sends".
  const notifies = existing.isOrganizer !== false && affected > 0 && !personalOnly;
  if (notifies && !opts.allowSend) throw new SendNotAllowedError("Updating a meeting you organize that has attendees (it sends them updates)");
  const updated = (await client
    .api(`/me/events/${enc(eventId)}`)
    .header("Prefer", preferTimeZone(tz))
    .patch(payload)) as GraphObject | undefined;
  return {
    event: updated ? shapeEventSummary(updated) : { id: eventId },
    time_zone: tz.graph,
    attendees_notified: notifies ? affected : 0,
  };
}

/**
 * Deletes an event. A meeting you organize that has attendees is cancelled
 * instead (POST /cancel, optional comment), which notifies them and moves it
 * to Deleted Items; anything else is deleted without notifying anyone.
 */
export async function deleteEvent(client: Client, eventId: string, opts: { allowSend: boolean; comment?: string }) {
  const existing = await getExisting(client, eventId);
  const attendees = addressSet(existing.attendees).size;
  if (existing.isOrganizer !== false && attendees > 0 && existing.isCancelled !== true) {
    if (!opts.allowSend) throw new SendNotAllowedError("Deleting a meeting you organize that has attendees (it sends them a cancellation)");
    await client.api(`/me/events/${enc(eventId)}/cancel`).post(opts.comment ? { Comment: opts.comment } : {});
    return { deleted: true, cancelled: true, event_id: eventId, attendees_notified: attendees };
  }
  await client.api(`/me/events/${enc(eventId)}`).delete();
  return { deleted: true, cancelled: false, event_id: eventId, attendees_notified: 0 };
}

export type EventResponse = "accept" | "tentative" | "decline";
const RESPONSE_ACTION: Record<EventResponse, string> = { accept: "accept", tentative: "tentativelyAccept", decline: "decline" };

export async function respondToEvent(
  client: Client,
  eventId: string,
  response: EventResponse,
  opts: { sendResponse: boolean; comment?: string; allowSend: boolean }
) {
  if (opts.sendResponse && !opts.allowSend) throw new SendNotAllowedError("Replying to the organizer (send_response=true)");
  const body: GraphObject = { sendResponse: opts.sendResponse };
  if (opts.comment && opts.sendResponse) body.comment = opts.comment;
  await client.api(`/me/events/${enc(eventId)}/${RESPONSE_ACTION[response]}`).post(body);
  return {
    event_id: eventId,
    response,
    organizer_notified: opts.sendResponse,
    comment_dropped: opts.comment && !opts.sendResponse ? true : undefined,
  };
}

// ---------------------------------------------------------------------------
// Availability
// ---------------------------------------------------------------------------

export interface ScheduleOptions {
  schedules: string[];
  start: string;
  end: string;
  timeZone: string;
  intervalMinutes?: number;
}

const MAX_SCHEDULE_WINDOW_MS = 62 * 86_400_000;

/** Free/busy for up to 20 people/rooms over a window shorter than 62 days (work/school accounts only). */
export async function getSchedule(client: Client, opts: ScheduleOptions) {
  const tz = resolveTimeZone(opts.timeZone);
  if (opts.schedules.length === 0 || opts.schedules.length > 20) throw new Error("Pass between 1 and 20 schedules (email addresses).");
  const span = windowMs(opts.start, opts.end, tz);
  if (span !== undefined && span <= 0) throw new Error("end must be after start.");
  if (span !== undefined && span >= MAX_SCHEDULE_WINDOW_MS) throw new Error("The window must be shorter than 62 days.");
  const interval = Math.min(Math.max(Math.trunc(opts.intervalMinutes ?? 30), 5), 1440);
  const result = (await client
    .api("/me/calendar/getSchedule")
    .header("Prefer", preferTimeZone(tz))
    .post({
      schedules: opts.schedules,
      startTime: toWindowDateTime(opts.start, tz, false),
      endTime: toWindowDateTime(opts.end, tz, true),
      availabilityViewInterval: interval,
    })) as { value?: GraphObject[] };
  return {
    time_zone: tz.graph,
    interval_minutes: interval,
    legend: "availability_view: one digit per interval — 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere",
    schedules: (result.value ?? []).map((s) => ({
      schedule: s.scheduleId as string,
      availability_view: s.availabilityView as string | undefined,
      items: ((s.scheduleItems as GraphObject[] | undefined) ?? []).map((i) => ({
        status: i.status as string | undefined,
        start: shortDateTime(i.start),
        end: shortDateTime(i.end),
        subject: (i.subject as string | undefined) || undefined,
        location: (i.location as string | undefined) || undefined,
        is_private: i.isPrivate === true ? true : undefined,
      })),
      working_hours: s.workingHours ?? undefined,
      error: (s.error as { message?: string } | undefined)?.message,
    })),
  };
}

export interface FindMeetingTimesOptions {
  attendees: Array<{ email: string; optional?: boolean }>;
  durationMinutes: number;
  start: string;
  end: string;
  timeZone: string;
  maxCandidates?: number;
  activityDomain?: "work" | "personal" | "unrestricted";
  minimumAttendeePercentage?: number;
  isOrganizerOptional?: boolean;
}

export const FIND_MEETING_TIMES_SCOPE_NOTE =
  "findMeetingTimes needs the Calendars.Read.Shared permission (work/school accounts only, granted by the shared-calendars " +
  "module). Use calendar_get_schedule (free/busy) instead.";

export async function findMeetingTimes(client: Client, opts: FindMeetingTimesOptions) {
  const tz = resolveTimeZone(opts.timeZone);
  const span = windowMs(opts.start, opts.end, tz);
  if (span !== undefined && span <= 0) throw new Error("end must be after start.");
  const body = {
    attendees: opts.attendees.map((a) => ({ type: a.optional ? "optional" : "required", emailAddress: { address: a.email } })),
    timeConstraint: {
      activityDomain: opts.activityDomain ?? "work",
      timeSlots: [{ start: toWindowDateTime(opts.start, tz, false), end: toWindowDateTime(opts.end, tz, true) }],
    },
    meetingDuration: `PT${Math.trunc(opts.durationMinutes)}M`,
    maxCandidates: Math.min(Math.max(Math.trunc(opts.maxCandidates ?? 10), 1), 50),
    isOrganizerOptional: opts.isOrganizerOptional ?? false,
    returnSuggestionReasons: true,
    ...(opts.minimumAttendeePercentage !== undefined ? { minimumAttendeePercentage: opts.minimumAttendeePercentage } : {}),
  };
  let result: GraphObject;
  try {
    result = (await client.api("/me/findMeetingTimes").header("Prefer", preferTimeZone(tz)).post(body)) as GraphObject;
  } catch (err) {
    if ((err as { statusCode?: number }).statusCode === 403) throw new Error(FIND_MEETING_TIMES_SCOPE_NOTE);
    throw err;
  }
  return {
    time_zone: tz.graph,
    empty_suggestions_reason: (result.emptySuggestionsReason as string | undefined) || undefined,
    suggestions: ((result.meetingTimeSuggestions as GraphObject[] | undefined) ?? []).map((s) => {
      const slot = s.meetingTimeSlot as { start?: GraphDateTime; end?: GraphDateTime } | undefined;
      return {
        start: shortDateTime(slot?.start),
        end: shortDateTime(slot?.end),
        confidence: s.confidence as number | undefined,
        organizer_availability: s.organizerAvailability as string | undefined,
        attendees: ((s.attendeeAvailability as GraphObject[] | undefined) ?? []).map((a) => ({
          address: (a.attendee as { emailAddress?: GraphEmail } | undefined)?.emailAddress?.address,
          availability: a.availability as string | undefined,
        })),
        reason: (s.suggestionReason as string | undefined) || undefined,
      };
    }),
  };
}
