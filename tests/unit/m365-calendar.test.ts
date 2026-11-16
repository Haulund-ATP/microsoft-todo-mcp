import { describe, expect, it, vi, beforeEach } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

const shared = vi.hoisted(() => ({
  audits: [] as Array<{ action: string; connectionId?: string; detail?: string }>,
  client: undefined as unknown,
}));

vi.mock("../../src/config/env.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../src/config/env.js")>()),
  loadEnv: () => ({ DEFAULT_TIMEZONE: "Europe/Copenhagen", LOG_LEVEL: "silent" }),
}));
vi.mock("../../src/logging/audit.js", () => ({
  audit: (event: { action: string; connectionId?: string; detail?: string }) => shared.audits.push(event),
}));
vi.mock("../../src/graph/client.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../../src/graph/client.js")>()),
  createGraphClientForConnection: () => shared.client,
}));

import * as cal from "../../src/graph/m365/calendar.js";
import { registerCalendarTools } from "../../src/tools/m365/calendarTools.js";
import type { ToolContext } from "../../src/tools/context.js";

interface Call {
  path: string;
  method?: string;
  query: Record<string, unknown>;
  headers: Record<string, string>;
  select?: string;
  top?: number;
  orderby?: string;
  filter?: string;
  search?: string;
  body?: unknown;
}

/** Fake Graph client: every request is recorded; `respond` decides the result (throw for errors). */
function fakeGraph(respond: (call: Call) => unknown = () => ({})) {
  const calls: Call[] = [];
  const client = {
    api(path: string) {
      const call: Call = { path, query: {}, headers: {} };
      const finish = async (method: string, body?: unknown) => {
        call.method = method;
        if (body !== undefined) call.body = body;
        calls.push(call);
        return respond(call);
      };
      const builder = {
        select: (s: string | string[]) => ((call.select = Array.isArray(s) ? s.join(",") : s), builder),
        top: (n: number) => ((call.top = n), builder),
        orderby: (o: string) => ((call.orderby = o), builder),
        filter: (f: string) => ((call.filter = f), builder),
        search: (s: string) => ((call.search = s), builder),
        query: (q: Record<string, unknown>) => (Object.assign(call.query, q), builder),
        header: (k: string, v: string) => ((call.headers[k] = v), builder),
        get: () => finish("GET"),
        post: (body: unknown) => finish("POST", body),
        patch: (body: unknown) => finish("PATCH", body),
        delete: () => finish("DELETE"),
      };
      return builder;
    },
  } as unknown as Client;
  return { client, calls };
}

const CPH = cal.resolveTimeZone("Europe/Copenhagen");
const PREFER_CPH = 'outlook.timezone="Romance Standard Time"';

describe("time zones", () => {
  it("maps IANA zones Graph may not accept to Windows names, and back", () => {
    expect(CPH).toEqual({ graph: "Romance Standard Time", iana: "Europe/Copenhagen" });
    expect(cal.resolveTimeZone("Pacific Standard Time")).toEqual({ graph: "Pacific Standard Time", iana: "America/Los_Angeles" });
    expect(cal.resolveTimeZone("Asia/Singapore")).toEqual({ graph: "Asia/Singapore", iana: "Asia/Singapore" });
    expect(cal.resolveTimeZone("Not A Zone")).toEqual({ graph: "Not A Zone" });
  });

  it("turns local dates into UTC range instants across DST", () => {
    expect(cal.toRangeInstant("2026-11-16", CPH, false)).toBe("2026-11-15T23:00:00.000Z");
    expect(cal.toRangeInstant("2026-11-16", CPH, true)).toBe("2026-11-16T23:00:00.000Z");
    expect(cal.toRangeInstant("2026-07-01T09:00", CPH, false)).toBe("2026-07-01T07:00:00.000Z");
    expect(cal.toRangeInstant("2026-07-01T09:00+02:00", cal.resolveTimeZone("Not A Zone"), false)).toBe("2026-07-01T07:00:00.000Z");
    expect(() => cal.toRangeInstant("2026-07-01", cal.resolveTimeZone("Not A Zone"), false)).toThrow(/UTC offset/);
  });

  it("keeps wall-clock times local and converts offset times into the zone", () => {
    expect(cal.toEventDateTime("2026-11-20T10:00", CPH)).toEqual({ dateTime: "2026-11-20T10:00:00", timeZone: "Romance Standard Time" });
    expect(cal.toEventDateTime("2026-11-20T09:00:00Z", CPH)).toEqual({ dateTime: "2026-11-20T10:00:00", timeZone: "Romance Standard Time" });
    expect(() => cal.toEventDateTime("2026-11-20", CPH)).toThrow(/all_day/);
    expect(() => cal.parseDateTimeInput("2026-02-30")).toThrow();
  });
});

describe("calendar reads", () => {
  it("lists events through calendarView with the range, Prefer header, $select and default page size", async () => {
    const g = fakeGraph(() => ({
      value: [
        {
          id: "e1",
          subject: "Standup",
          start: { dateTime: "2026-11-16T09:00:00.0000000", timeZone: "Romance Standard Time" },
          end: { dateTime: "2026-11-16T09:15:00.0000000", timeZone: "Romance Standard Time" },
          isAllDay: false,
          isOnlineMeeting: true,
          onlineMeeting: { joinUrl: "https://teams.example.com/join/1" },
          organizer: { emailAddress: { name: "Org", address: "org@example.com" } },
          responseStatus: { response: "accepted" },
          location: { displayName: "" },
        },
        {
          id: "e2",
          subject: "Trip",
          start: { dateTime: "2026-11-17T00:00:00.0000000", timeZone: "Romance Standard Time" },
          end: { dateTime: "2026-11-19T00:00:00.0000000", timeZone: "Romance Standard Time" },
          isAllDay: true,
        },
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skip=25",
    }));
    const result = await cal.listEvents(g.client, { start: "2026-11-16", end: "2026-11-20", timeZone: "Europe/Copenhagen" });
    const call = g.calls[0]!;
    expect(call.path).toBe("/me/calendarView");
    expect(call.query).toEqual({ startDateTime: "2026-11-15T23:00:00.000Z", endDateTime: "2026-11-20T23:00:00.000Z" });
    expect(call.headers.Prefer).toBe(PREFER_CPH);
    expect(call.top).toBe(25);
    expect(call.orderby).toBe("start/dateTime");
    expect(call.select).toContain("onlineMeeting");
    expect(call.select).not.toContain("body");
    expect(result.events[0]).toMatchObject({ id: "e1", start: "2026-11-16T09:00", end: "2026-11-16T09:15", join_url: "https://teams.example.com/join/1", my_response: "accepted" });
    expect(result.events[0]!.location).toBeUndefined();
    expect(result.events[1]).toMatchObject({ is_all_day: true, start: "2026-11-17", end: "2026-11-18" });
    expect(result.next_page_token).toBe("https://graph.microsoft.com/v1.0/me/calendarView?$skip=25");
  });

  it("uses the calendar's own calendarView, caps the page size, and pages only to Graph URLs", async () => {
    const g = fakeGraph(() => ({ value: [] }));
    await cal.listEvents(g.client, { start: "2026-11-16", end: "2026-11-16", timeZone: "UTC", calendarId: "a/b", top: 500 });
    expect(g.calls[0]!.path).toBe("/me/calendars/a%2Fb/calendarView");
    expect(g.calls[0]!.top).toBe(100);
    expect(g.calls[0]!.headers.Prefer).toBe('outlook.timezone="UTC"');

    const token = "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=x&$skip=25";
    await cal.listEvents(g.client, { start: "2026-11-16", end: "2026-11-16", timeZone: "Europe/Copenhagen", pageToken: token });
    expect(g.calls[1]!.path).toBe(token);
    expect(g.calls[1]!.headers.Prefer).toBe(PREFER_CPH);
    await expect(
      cal.listEvents(g.client, { start: "2026-11-16", end: "2026-11-16", timeZone: "UTC", pageToken: "https://evil.example.com/x" })
    ).rejects.toThrow(/next_page_token/);
  });

  it("returns an event's body as text with attendees and their responses", async () => {
    const g = fakeGraph(() => ({
      id: "e1",
      subject: "Review",
      start: { dateTime: "2026-11-16T13:00:00.0000000", timeZone: "Romance Standard Time" },
      end: { dateTime: "2026-11-16T14:00:00.0000000", timeZone: "Romance Standard Time" },
      body: { contentType: "html", content: "<html><body><p>Agenda &amp; notes</p><p>" + "x".repeat(50) + "</p></body></html>" },
      attendees: [{ type: "required", emailAddress: { name: "A", address: "a@example.com" }, status: { response: "tentativelyAccepted" } }],
      isReminderOn: false,
    }));
    const event = await cal.getEvent(g.client, "e1", { timeZone: "Europe/Copenhagen", maxBodyChars: 20 });
    expect(g.calls[0]!.headers.Prefer).toBe(PREFER_CPH);
    expect(g.calls[0]!.select).toContain("attendees");
    expect(event.body).toBe("Agenda & notes\nxxxxx");
    expect(event.body_truncated).toBe(true);
    expect(event.attendees).toEqual([{ name: "A", address: "a@example.com", type: "required", response: "tentativelyAccepted" }]);
    expect(event.reminder_minutes).toBeNull();
  });
});

describe("event payloads", () => {
  it("builds a timed event with location, body, reminder and a Teams meeting", () => {
    const payload = cal.buildEventPayload(
      { subject: "Sync", start: "2026-11-20T10:00", end: "2026-11-20T10:30", location: "Room 1", body: "Hi", onlineMeeting: true, reminderMinutes: 10 },
      CPH,
      "create"
    );
    expect(payload).toEqual({
      subject: "Sync",
      start: { dateTime: "2026-11-20T10:00:00", timeZone: "Romance Standard Time" },
      end: { dateTime: "2026-11-20T10:30:00", timeZone: "Romance Standard Time" },
      location: { displayName: "Room 1" },
      body: { contentType: "text", content: "Hi" },
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
      isReminderOn: true,
      reminderMinutesBeforeStart: 10,
    });
  });

  it("builds all-day events with an exclusive end the day after the inclusive last day", () => {
    expect(cal.buildEventPayload({ subject: "Off", start: "2026-12-24", allDay: true }, CPH, "create")).toMatchObject({
      isAllDay: true,
      start: { dateTime: "2026-12-24T00:00:00", timeZone: "Romance Standard Time" },
      end: { dateTime: "2026-12-25T00:00:00", timeZone: "Romance Standard Time" },
    });
    expect(cal.buildEventPayload({ subject: "Off", start: "2026-12-30", end: "2027-01-01", allDay: true }, CPH, "create").end).toEqual({
      dateTime: "2027-01-02T00:00:00",
      timeZone: "Romance Standard Time",
    });
    expect(() => cal.buildEventPayload({ subject: "Off", start: "2026-12-24T10:00", allDay: true }, CPH, "create")).toThrow(/dates only/);
  });

  it("rejects incomplete or inverted times and builds partial updates", () => {
    expect(() => cal.buildEventPayload({ subject: "X", start: "2026-11-20T10:00" }, CPH, "create")).toThrow(/start and end/);
    expect(() => cal.buildEventPayload({ subject: "X", start: "2026-11-20T10:00", end: "2026-11-20T09:00" }, CPH, "create")).toThrow(/after start/);
    expect(cal.buildEventPayload({ reminderMinutes: null, location: null }, CPH, "update")).toEqual({
      isReminderOn: false,
      location: { displayName: "" },
    });
    expect(cal.buildEventPayload({ attendees: [{ email: "b@example.com", type: "optional" }] }, CPH, "update")).toEqual({
      attendees: [{ emailAddress: { address: "b@example.com" }, type: "optional" }],
    });
  });
});

describe("sending rules", () => {
  const attendees = [{ email: "a@example.com" }, { email: "b@example.com", name: "B" }];
  const createdEvent = { id: "new", subject: "Sync", start: { dateTime: "2026-11-20T10:00:00" }, end: { dateTime: "2026-11-20T10:30:00" } };
  const base = { subject: "Sync", start: "2026-11-20T10:00", end: "2026-11-20T10:30" };

  it("rejects creating an event with attendees without allowSend, before any Graph call", async () => {
    const g = fakeGraph(() => createdEvent);
    await expect(cal.createEvent(g.client, { ...base, attendees }, { timeZone: "Europe/Copenhagen", allowSend: false })).rejects.toBeInstanceOf(
      cal.SendNotAllowedError
    );
    expect(g.calls).toEqual([]);
  });

  it("creates attendee-free events without allowSend, and invitations with it", async () => {
    const g = fakeGraph(() => createdEvent);
    const solo = await cal.createEvent(g.client, base, { timeZone: "Europe/Copenhagen", allowSend: false, calendarId: "cal1" });
    expect(solo.invitations_sent).toBe(0);
    expect(g.calls[0]).toMatchObject({ method: "POST", path: "/me/calendars/cal1/events", headers: { Prefer: PREFER_CPH } });

    const meeting = await cal.createEvent(g.client, { ...base, attendees }, { timeZone: "Europe/Copenhagen", allowSend: true });
    expect(meeting.invitations_sent).toBe(2);
    expect(g.calls[1]!.path).toBe("/me/events");
    expect((g.calls[1]!.body as { attendees: unknown }).attendees).toEqual([
      { emailAddress: { address: "a@example.com" }, type: "required" },
      { emailAddress: { address: "b@example.com", name: "B" }, type: "required" },
    ]);
  });

  const organizedMeeting = (call: Call) =>
    call.method === "GET"
      ? { id: "m1", isOrganizer: true, attendees: [{ emailAddress: { address: "a@example.com" } }] }
      : { id: "m1", subject: "Moved" };

  it("blocks meeting updates that notify attendees without allowSend, but allows reminder-only changes", async () => {
    const g = fakeGraph(organizedMeeting);
    await expect(cal.updateEvent(g.client, "m1", { subject: "Moved" }, { timeZone: "UTC", allowSend: false })).rejects.toBeInstanceOf(
      cal.SendNotAllowedError
    );
    expect(g.calls.filter((c) => c.method === "PATCH")).toEqual([]);

    const reminder = await cal.updateEvent(g.client, "m1", { reminderMinutes: 5 }, { timeZone: "UTC", allowSend: false });
    expect(reminder.attendees_notified).toBe(0);
    expect(g.calls.at(-1)).toMatchObject({ method: "PATCH", body: { isReminderOn: true, reminderMinutesBeforeStart: 5 } });
  });

  it("counts everyone affected when a meeting update is sent, and lets invitees edit their own copy", async () => {
    const g = fakeGraph(organizedMeeting);
    const result = await cal.updateEvent(g.client, "m1", { attendees: [{ email: "b@example.com" }] }, { timeZone: "UTC", allowSend: true });
    expect(result.attendees_notified).toBe(2); // a removed (cancellation) + b added (invitation)

    const invitee = fakeGraph((call) => (call.method === "GET" ? { isOrganizer: false, attendees: [{ emailAddress: { address: "me@example.com" } }] } : { id: "m2" }));
    const own = await cal.updateEvent(invitee.client, "m2", { subject: "My note" }, { timeZone: "UTC", allowSend: false });
    expect(own.attendees_notified).toBe(0);
  });

  it("cancels organized meetings (with comment) only with allowSend and deletes everything else", async () => {
    const g = fakeGraph(organizedMeeting);
    await expect(cal.deleteEvent(g.client, "m1", { allowSend: false })).rejects.toBeInstanceOf(cal.SendNotAllowedError);
    expect(g.calls.some((c) => c.method !== "GET")).toBe(false);

    const cancelled = await cal.deleteEvent(g.client, "m1", { allowSend: true, comment: "Sorry" });
    expect(cancelled).toMatchObject({ cancelled: true, attendees_notified: 1 });
    expect(g.calls.at(-1)).toMatchObject({ method: "POST", path: "/me/events/m1/cancel", body: { Comment: "Sorry" } });

    const solo = fakeGraph((call) => (call.method === "GET" ? { isOrganizer: true, attendees: [] } : undefined));
    const deleted = await cal.deleteEvent(solo.client, "s1", { allowSend: false });
    expect(deleted).toMatchObject({ cancelled: false, attendees_notified: 0 });
    expect(solo.calls.at(-1)).toMatchObject({ method: "DELETE", path: "/me/events/s1" });
  });

  it("only replies to the organizer with allowSend", async () => {
    const g = fakeGraph(() => undefined);
    await expect(cal.respondToEvent(g.client, "i1", "accept", { sendResponse: true, allowSend: false })).rejects.toBeInstanceOf(cal.SendNotAllowedError);
    const quiet = await cal.respondToEvent(g.client, "i1", "tentative", { sendResponse: false, comment: "maybe", allowSend: false });
    expect(g.calls[0]).toMatchObject({ method: "POST", path: "/me/events/i1/tentativelyAccept", body: { sendResponse: false } });
    expect(quiet).toMatchObject({ organizer_notified: false, comment_dropped: true });
    await cal.respondToEvent(g.client, "i1", "decline", { sendResponse: true, comment: "Away", allowSend: true });
    expect(g.calls[1]).toMatchObject({ path: "/me/events/i1/decline", body: { sendResponse: true, comment: "Away" } });
  });
});

describe("availability", () => {
  it("posts getSchedule with local window boundaries and shapes the result", async () => {
    const g = fakeGraph(() => ({
      value: [
        {
          scheduleId: "a@example.com",
          availabilityView: "0220",
          scheduleItems: [{ status: "busy", start: { dateTime: "2026-11-16T09:00:00.0000000" }, end: { dateTime: "2026-11-16T10:00:00.0000000" } }],
        },
      ],
    }));
    const result = await cal.getSchedule(g.client, { schedules: ["a@example.com"], start: "2026-11-16", end: "2026-11-16", timeZone: "Europe/Copenhagen", intervalMinutes: 1 });
    expect(g.calls[0]).toMatchObject({
      method: "POST",
      path: "/me/calendar/getSchedule",
      headers: { Prefer: PREFER_CPH },
      body: {
        schedules: ["a@example.com"],
        startTime: { dateTime: "2026-11-16T00:00:00", timeZone: "Romance Standard Time" },
        endTime: { dateTime: "2026-11-17T00:00:00", timeZone: "Romance Standard Time" },
        availabilityViewInterval: 5,
      },
    });
    expect(result.schedules[0]).toMatchObject({ schedule: "a@example.com", availability_view: "0220", items: [{ status: "busy", start: "2026-11-16T09:00", end: "2026-11-16T10:00" }] });
    await expect(cal.getSchedule(g.client, { schedules: ["a@example.com"], start: "2026-01-01", end: "2026-03-15", timeZone: "UTC" })).rejects.toThrow(/62 days/);
  });

  it("asks findMeetingTimes for a duration and explains a refused permission", async () => {
    const g = fakeGraph(() => ({
      emptySuggestionsReason: "",
      meetingTimeSuggestions: [
        { confidence: 100, organizerAvailability: "free", meetingTimeSlot: { start: { dateTime: "2026-11-17T10:00:00.0000000" }, end: { dateTime: "2026-11-17T10:45:00.0000000" } } },
      ],
    }));
    const result = await cal.findMeetingTimes(g.client, { attendees: [{ email: "a@example.com" }, { email: "b@example.com", optional: true }], durationMinutes: 45, start: "2026-11-17", end: "2026-11-18", timeZone: "Europe/Copenhagen" });
    expect(g.calls[0]!.body).toMatchObject({
      meetingDuration: "PT45M",
      attendees: [
        { type: "required", emailAddress: { address: "a@example.com" } },
        { type: "optional", emailAddress: { address: "b@example.com" } },
      ],
      timeConstraint: { activityDomain: "work", timeSlots: [{ start: { dateTime: "2026-11-17T00:00:00" }, end: { dateTime: "2026-11-19T00:00:00" } }] },
    });
    expect(result.suggestions[0]).toMatchObject({ start: "2026-11-17T10:00", end: "2026-11-17T10:45", confidence: 100 });

    const denied = fakeGraph(() => {
      throw Object.assign(new Error("Access is denied."), { statusCode: 403, code: "ErrorAccessDenied" });
    });
    await expect(cal.findMeetingTimes(denied.client, { attendees: [], durationMinutes: 30, start: "2026-11-17", end: "2026-11-17", timeZone: "UTC" })).rejects.toThrow(/Calendars.Read.Shared/);
  });
});

describe("calendar tools", () => {
  type Handler = (args: Record<string, unknown>) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;
  function register(allowSend: boolean, modules: ToolContext["modules"] = ["calendar", "calendar_shared"]) {
    const tools = new Map<string, { config: { description: string }; handler: Handler }>();
    const server = {
      registerTool: (name: string, config: { description: string }, handler: Handler) => tools.set(name, { config, handler }),
    } as unknown as McpServer;
    const ctx: ToolContext = { profileAlias: "work", boundConnectionId: "conn-1", correlationId: "corr", modules, allowSend };
    registerCalendarTools(server, ctx);
    return tools;
  }

  beforeEach(() => {
    shared.audits.length = 0;
  });

  it("registers the calendar_ tool family", () => {
    expect([...register(false).keys()].sort()).toEqual([
      "calendar_create_event",
      "calendar_delete_event",
      "calendar_find_meeting_times",
      "calendar_get_event",
      "calendar_get_schedule",
      "calendar_list_calendars",
      "calendar_list_events",
      "calendar_respond_to_event",
      "calendar_update_event",
    ]);
  });

  it("only offers the meeting-time finder with the shared-calendars module (personal accounts cannot have it)", () => {
    expect([...register(false, ["calendar"]).keys()]).not.toContain("calendar_find_meeting_times");
  });

  it("rejects invitations without allowSend as a tool error and audits nothing", async () => {
    const g = fakeGraph(() => ({ id: "x" }));
    shared.client = g.client;
    const result = await register(false).get("calendar_create_event")!.handler({
      subject: "Sync",
      start: "2026-11-20T10:00",
      end: "2026-11-20T11:00",
      attendees: [{ email: "a@example.com" }],
    });
    expect(result.isError).toBe(true);
    expect(result.content[0]!.text).toContain("allowSend");
    expect(g.calls).toEqual([]);
    expect(shared.audits).toEqual([]);
  });

  it("audits sent invitations by count, using the server's default time zone", async () => {
    const g = fakeGraph(() => ({ id: "x", subject: "Sync" }));
    shared.client = g.client;
    const result = await register(true).get("calendar_create_event")!.handler({
      subject: "Sync",
      start: "2026-11-20T10:00",
      end: "2026-11-20T11:00",
      attendees: [{ email: "a@example.com" }, { email: "b@example.com" }],
    });
    expect(result.isError).toBeUndefined();
    expect(g.calls[0]!.headers.Prefer).toBe(PREFER_CPH);
    expect(shared.audits).toEqual([expect.objectContaining({ action: "m365.send", connectionId: "conn-1", detail: "event attendees=2" })]);
  });

  it("audits deletes, and cancellations as both a send and a delete", async () => {
    shared.client = fakeGraph((call) => (call.method === "GET" ? { isOrganizer: true, attendees: [] } : undefined)).client;
    await register(false).get("calendar_delete_event")!.handler({ event_id: "s1" });
    expect(shared.audits).toEqual([expect.objectContaining({ action: "m365.delete", connectionId: "conn-1", detail: "event" })]);

    shared.audits.length = 0;
    shared.client = fakeGraph((call) =>
      call.method === "GET" ? { isOrganizer: true, attendees: [{ emailAddress: { address: "a@example.com" } }, { emailAddress: { address: "b@example.com" } }] } : undefined
    ).client;
    await register(true).get("calendar_delete_event")!.handler({ event_id: "m1", comment: "Off" });
    expect(shared.audits.map((a) => [a.action, a.detail])).toEqual([
      ["m365.send", "event cancel attendees=2"],
      ["m365.delete", "event cancelled"],
    ]);
  });

  it("defaults send_response to the endpoint's allowSend", async () => {
    const quiet = fakeGraph(() => undefined);
    shared.client = quiet.client;
    await register(false).get("calendar_respond_to_event")!.handler({ event_id: "i1", response: "accept" });
    expect(quiet.calls[0]!.body).toEqual({ sendResponse: false });
    expect(shared.audits).toEqual([]);

    const loud = fakeGraph(() => undefined);
    shared.client = loud.client;
    await register(true).get("calendar_respond_to_event")!.handler({ event_id: "i1", response: "accept" });
    expect(loud.calls[0]!.body).toEqual({ sendResponse: true });
    expect(shared.audits).toEqual([expect.objectContaining({ action: "m365.send", detail: "event response=accept" })]);
  });
});
