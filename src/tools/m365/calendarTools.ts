import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { hasModule, type ToolContext } from "../context.js";
import { connectionIdInputShape } from "../accounts.js";
import { auditM365, runM365Tool, UNTRUSTED_CONTENT_NOTE } from "./common.js";
import { defaultTimezone } from "../../graph/time.js";
import * as cal from "../../graph/m365/calendar.js";

const dateOrDateTime = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/, "Use YYYY-MM-DD or YYYY-MM-DDTHH:mm[:ss][offset].");

const timeZoneField = z
  .string()
  .min(1)
  .max(64)
  .optional()
  .describe("IANA (e.g. Europe/Copenhagen) or Windows time zone name. Defaults to the server's time zone. Times without an offset are read, and returned, in this zone.");

const pageTokenField = z.string().optional().describe("next_page_token from the previous result, passed back unchanged.");
const topField = z.number().int().min(1).max(cal.MAX_PAGE_SIZE).optional().describe(`Page size (default ${cal.DEFAULT_PAGE_SIZE}, max ${cal.MAX_PAGE_SIZE}).`);

const attendeeShape = z.object({
  email: z.string().email(),
  name: z.string().max(255).optional(),
  type: z.enum(["required", "optional", "resource"]).optional().describe("Default required. Use resource for rooms/equipment."),
});

const eventFields = {
  subject: z.string().min(1).max(255),
  start: dateOrDateTime.describe("Timed: YYYY-MM-DDTHH:mm local time in time_zone (or with a UTC offset). All-day: YYYY-MM-DD."),
  end: dateOrDateTime.describe("Timed: end time (after start). All-day: the last day, inclusive (defaults to start)."),
  all_day: z.boolean().describe("All-day event: start/end are dates."),
  location: z.string().max(255),
  body: z.string().max(20_000).describe("Plain-text description."),
  attendees: z.array(attendeeShape).max(100),
  online_meeting: z.boolean().describe("Add an online meeting (Teams by default) with a join link."),
  online_meeting_provider: z
    .enum(["teamsForBusiness", "skypeForBusiness", "skypeForConsumer"])
    .describe("Default teamsForBusiness. See allowed_online_meeting_providers in calendar_list_calendars."),
  reminder_minutes: z.number().int().min(0).max(40_320).describe("Reminder minutes before start."),
  show_as: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]),
  importance: z.enum(["low", "normal", "high"]),
};

export const createEventInputShape = {
  subject: eventFields.subject,
  start: eventFields.start,
  end: eventFields.end.optional(),
  all_day: eventFields.all_day.optional(),
  location: eventFields.location.optional(),
  body: eventFields.body.optional(),
  attendees: eventFields.attendees.optional(),
  online_meeting: eventFields.online_meeting.optional(),
  online_meeting_provider: eventFields.online_meeting_provider.optional(),
  reminder_minutes: eventFields.reminder_minutes.optional(),
  show_as: eventFields.show_as.optional(),
  importance: eventFields.importance.optional(),
  calendar_id: z.string().min(1).optional().describe("Calendar to create it in (default calendar if omitted)."),
  time_zone: timeZoneField,
  ...connectionIdInputShape,
};

export const updateEventInputShape = {
  event_id: z.string().min(1),
  subject: eventFields.subject.optional(),
  start: eventFields.start.optional(),
  end: eventFields.end.optional(),
  all_day: eventFields.all_day
    .optional()
    .describe("true: make it all-day (pass start, and end for multi-day). false: make it timed (pass start and end)."),
  location: eventFields.location.nullable().optional().describe("null clears it."),
  body: eventFields.body.nullable().optional().describe("Replaces the description (plain text); null clears it."),
  attendees: eventFields.attendees.optional().describe("Replaces the whole attendee list."),
  online_meeting: eventFields.online_meeting.optional(),
  online_meeting_provider: eventFields.online_meeting_provider.optional(),
  reminder_minutes: eventFields.reminder_minutes.nullable().optional().describe("Minutes before start; null turns the reminder off."),
  show_as: eventFields.show_as.optional(),
  importance: eventFields.importance.optional(),
  time_zone: timeZoneField,
  ...connectionIdInputShape,
};

export function registerCalendarTools(server: McpServer, ctx: ToolContext): void {
  const allowSend = ctx.allowSend === true;
  const tz = (value: string | undefined) => value ?? defaultTimezone();
  const sendRule = allowSend
    ? " Sending is enabled on this endpoint: attendees are emailed."
    : " Sending is not enabled on this endpoint, so calls that would email attendees are rejected.";

  server.registerTool(
    "calendar_list_calendars",
    {
      title: "List calendars",
      description: "Lists the signed-in user's calendars (id, name, default flag, edit right, owner, online meeting providers).",
      inputSchema: { top: topField, page_token: pageTokenField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ top, page_token, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_list_calendars", connection_id, (client) =>
        cal.listCalendars(client, { top, pageToken: page_token })
      )
  );

  server.registerTool(
    "calendar_list_events",
    {
      title: "List calendar events in a date range",
      description:
        "Lists events overlapping a date range (recurring meetings expanded into occurrences), ordered by start, in the " +
        "default calendar or calendar_id. start/end may be dates (end date inclusive) or local date-times; times come back " +
        "in time_zone. All-day events show inclusive dates. Use calendar_get_event for body and attendees." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        start: dateOrDateTime.describe("Range start: YYYY-MM-DD (local midnight) or YYYY-MM-DDTHH:mm."),
        end: dateOrDateTime.describe("Range end: YYYY-MM-DD (includes that whole day) or YYYY-MM-DDTHH:mm (exclusive)."),
        calendar_id: z.string().min(1).optional(),
        time_zone: timeZoneField,
        top: topField,
        page_token: pageTokenField.describe("next_page_token from the previous result; start/end are then taken from the token."),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ start, end, calendar_id, time_zone, top, page_token, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_list_events", connection_id, (client) =>
        cal.listEvents(client, { start, end, calendarId: calendar_id, timeZone: tz(time_zone), top, pageToken: page_token })
      )
  );

  server.registerTool(
    "calendar_get_event",
    {
      title: "Get calendar event",
      description:
        "Gets one event: times, location, body as plain text, organizer, attendees with their responses, online meeting " +
        "join URL, reminder and recurrence." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        event_id: z.string().min(1),
        time_zone: timeZoneField,
        max_body_chars: z.number().int().min(0).max(100_000).optional().describe("Body is cut at this length (default 8000)."),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ event_id, time_zone, max_body_chars, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_get_event", connection_id, (client) =>
        cal.getEvent(client, event_id, { timeZone: tz(time_zone), maxBodyChars: max_body_chars ?? 8000 })
      )
  );

  server.registerTool(
    "calendar_create_event",
    {
      title: "Create calendar event",
      description:
        "Creates an event (timed or all-day) with optional location, plain-text body, reminder, Teams online meeting and " +
        "attendees. With attendees it becomes a meeting and invitations are emailed to them." +
        sendRule,
      inputSchema: createEventInputShape,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: allowSend },
    },
    async (args) =>
      runM365Tool(ctx, "calendar", "calendar_create_event", args.connection_id, async (client, connectionId) => {
        const result = await cal.createEvent(
          client,
          {
            subject: args.subject,
            start: args.start,
            end: args.end,
            allDay: args.all_day,
            location: args.location,
            body: args.body,
            attendees: args.attendees,
            onlineMeeting: args.online_meeting,
            onlineMeetingProvider: args.online_meeting_provider,
            reminderMinutes: args.reminder_minutes,
            showAs: args.show_as,
            importance: args.importance,
          },
          { timeZone: tz(args.time_zone), allowSend, calendarId: args.calendar_id }
        );
        if (result.invitations_sent > 0) auditM365(ctx, "m365.send", connectionId, `event attendees=${result.invitations_sent}`);
        return result;
      })
  );

  server.registerTool(
    "calendar_update_event",
    {
      title: "Update calendar event",
      description:
        "Changes only the given fields of an event. On a meeting you organize that has attendees, any change other than " +
        "the reminder emails an update to them (and adding/removing attendees invites/cancels them)." +
        sendRule,
      inputSchema: updateEventInputShape,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: allowSend },
    },
    async (args) =>
      runM365Tool(ctx, "calendar", "calendar_update_event", args.connection_id, async (client, connectionId) => {
        const result = await cal.updateEvent(
          client,
          args.event_id,
          {
            subject: args.subject,
            start: args.start,
            end: args.end,
            allDay: args.all_day,
            location: args.location,
            body: args.body,
            attendees: args.attendees,
            onlineMeeting: args.online_meeting,
            onlineMeetingProvider: args.online_meeting_provider,
            reminderMinutes: args.reminder_minutes,
            showAs: args.show_as,
            importance: args.importance,
          },
          { timeZone: tz(args.time_zone), allowSend }
        );
        if (result.attendees_notified > 0) auditM365(ctx, "m365.send", connectionId, `event update attendees=${result.attendees_notified}`);
        return result;
      })
  );

  server.registerTool(
    "calendar_delete_event",
    {
      title: "Delete calendar event",
      description:
        "Deletes an event (it goes to Deleted Items). A meeting you organize that has attendees is cancelled instead, which " +
        "emails them a cancellation with the optional comment. Deleting a meeting you were invited to does not tell the " +
        "organizer — use calendar_respond_to_event with decline for that." +
        sendRule,
      inputSchema: {
        event_id: z.string().min(1),
        comment: z.string().max(2000).optional().describe("Cancellation message to attendees (meetings you organize only)."),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: allowSend },
    },
    async ({ event_id, comment, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_delete_event", connection_id, async (client, connectionId) => {
        const result = await cal.deleteEvent(client, event_id, { allowSend, comment });
        if (result.cancelled) auditM365(ctx, "m365.send", connectionId, `event cancel attendees=${result.attendees_notified}`);
        auditM365(ctx, "m365.delete", connectionId, result.cancelled ? "event cancelled" : "event");
        return result;
      })
  );

  server.registerTool(
    "calendar_respond_to_event",
    {
      title: "Respond to meeting invitation",
      description:
        "Accepts, tentatively accepts or declines a meeting invitation. send_response=true emails the reply (and comment) " +
        `to the organizer; false only updates your calendar. Default: ${allowSend ? "true" : "false"}.` +
        sendRule,
      inputSchema: {
        event_id: z.string().min(1),
        response: z.enum(["accept", "tentative", "decline"]),
        comment: z.string().max(2000).optional().describe("Message to the organizer; only sent with send_response=true."),
        send_response: z.boolean().optional(),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: allowSend },
    },
    async ({ event_id, response, comment, send_response, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_respond_to_event", connection_id, async (client, connectionId) => {
        const result = await cal.respondToEvent(client, event_id, response, {
          sendResponse: send_response ?? allowSend,
          comment,
          allowSend,
        });
        if (result.organizer_notified) auditM365(ctx, "m365.send", connectionId, `event response=${response}`);
        return result;
      })
  );

  server.registerTool(
    "calendar_get_schedule",
    {
      title: "Get free/busy schedule",
      description:
        "Free/busy for up to 20 people, distribution lists or rooms (email addresses; include your own to see yourself) " +
        "over a window shorter than 62 days: busy blocks plus an availability string per interval. Work/school accounts " +
        "only; how much detail you see of others depends on their sharing settings." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        schedules: z.array(z.string().email()).min(1).max(20),
        start: dateOrDateTime.describe("Window start: YYYY-MM-DD (local midnight) or YYYY-MM-DDTHH:mm."),
        end: dateOrDateTime.describe("Window end: YYYY-MM-DD (includes that day) or YYYY-MM-DDTHH:mm."),
        interval_minutes: z.number().int().min(5).max(1440).optional().describe("Slot length for the availability string (default 30)."),
        time_zone: timeZoneField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ schedules, start, end, interval_minutes, time_zone, connection_id }) =>
      runM365Tool(ctx, "calendar", "calendar_get_schedule", connection_id, (client) =>
        cal.getSchedule(client, { schedules, start, end, intervalMinutes: interval_minutes, timeZone: tz(time_zone) })
      )
  );

  // findMeetingTimes needs Calendars.Read.Shared, which comes with the
  // work/school-only shared-calendars module, not the base calendar module.
  if (!hasModule(ctx, "calendar_shared")) return;
  server.registerTool(
    "calendar_find_meeting_times",
    {
      title: "Find meeting times",
      description:
        "Suggests meeting slots where you and the attendees are free within a window (Graph findMeetingTimes), with a " +
        "confidence score per slot. Work/school accounts only (shared-calendars module); if it is refused, use " +
        "calendar_get_schedule instead." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        attendees: z
          .array(z.object({ email: z.string().email(), optional: z.boolean().optional() }))
          .max(20)
          .describe("People to meet (empty: just find free time for yourself)."),
        duration_minutes: z.number().int().min(5).max(1440),
        start: dateOrDateTime.describe("Window start: YYYY-MM-DD or YYYY-MM-DDTHH:mm."),
        end: dateOrDateTime.describe("Window end: YYYY-MM-DD (includes that day) or YYYY-MM-DDTHH:mm."),
        max_candidates: z.number().int().min(1).max(50).optional().describe("Default 10."),
        activity_domain: z
          .enum(["work", "personal", "unrestricted"])
          .optional()
          .describe("work (default): working hours; personal: also weekends; unrestricted: any hour."),
        minimum_attendee_percentage: z.number().min(0).max(100).optional(),
        is_organizer_optional: z.boolean().optional(),
        time_zone: timeZoneField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "calendar_shared", "calendar_find_meeting_times", args.connection_id, (client) =>
        cal.findMeetingTimes(client, {
          attendees: args.attendees,
          durationMinutes: args.duration_minutes,
          start: args.start,
          end: args.end,
          maxCandidates: args.max_candidates,
          activityDomain: args.activity_domain,
          minimumAttendeePercentage: args.minimum_attendee_percentage,
          isOrganizerOptional: args.is_organizer_optional,
          timeZone: tz(args.time_zone),
        })
      )
  );
}
