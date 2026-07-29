import { loadEnv } from "../config/env.js";

/**
 * Date/timezone helpers. Default timezone is Europe/Copenhagen, but any
 * explicit timezone passed by the caller (tool input) or present on a Graph
 * dateTimeTimeZone value is respected instead. Tool outputs always include
 * both a machine-readable ISO 8601 (UTC) string and a human-readable local
 * representation.
 */

export function defaultTimezone(): string {
  return loadEnv().DEFAULT_TIMEZONE;
}

export interface DisplayableDateTime {
  iso: string; // ISO 8601, UTC
  local: string; // human-readable, in the resolved timezone
  timeZone: string;
}

/** Graph's To Do API represents due/reminder dates as { dateTime, timeZone }. */
export interface GraphDateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export function toGraphDateTimeTimeZone(isoOrDate: string, timeZone?: string): GraphDateTimeTimeZone {
  const tz = timeZone ?? defaultTimezone();
  const date = new Date(isoOrDate);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Invalid date/time: "${isoOrDate}"`);
  }
  // Microsoft Graph expects a local wall-clock time paired with an IANA
  // (or Windows) timeZone name, without a trailing "Z".
  return { dateTime: date.toISOString().replace("Z", ""), timeZone: tz };
}

export function fromGraphDateTimeTimeZone(value: GraphDateTimeTimeZone | undefined | null): DisplayableDateTime | undefined {
  if (!value?.dateTime) return undefined;
  const iso = value.dateTime.endsWith("Z") ? value.dateTime : `${value.dateTime}Z`;
  const date = new Date(iso);
  const local = new Intl.DateTimeFormat("en-DK", {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone: normalizeTimeZone(value.timeZone),
  }).format(date);
  return { iso: date.toISOString(), local, timeZone: value.timeZone };
}

/** Graph sometimes returns Windows timezone names (e.g. "Romance Standard Time"); Intl needs IANA. */
function normalizeTimeZone(tz: string): string {
  const map: Record<string, string> = {
    "Romance Standard Time": "Europe/Copenhagen",
    UTC: "UTC",
  };
  return map[tz] ?? tz;
}
