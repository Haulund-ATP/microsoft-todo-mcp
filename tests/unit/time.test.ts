import { describe, expect, it } from "vitest";
import { fromGraphDateTimeTimeZone, toGraphDateTimeTimeZone } from "../../src/graph/time.js";

describe("toGraphDateTimeTimeZone", () => {
  it("strips the trailing Z and defaults the timezone", () => {
    const result = toGraphDateTimeTimeZone("2026-08-01T10:00:00.000Z", "Europe/Copenhagen");
    expect(result.timeZone).toBe("Europe/Copenhagen");
    expect(result.dateTime.endsWith("Z")).toBe(false);
  });

  it("throws on an invalid date string", () => {
    expect(() => toGraphDateTimeTimeZone("not-a-date")).toThrow();
  });
});

describe("fromGraphDateTimeTimeZone", () => {
  it("returns undefined for a missing value", () => {
    expect(fromGraphDateTimeTimeZone(undefined)).toBeUndefined();
    expect(fromGraphDateTimeTimeZone(null)).toBeUndefined();
  });

  it("produces both an ISO and a human-readable local string", () => {
    const result = fromGraphDateTimeTimeZone({ dateTime: "2026-08-01T10:00:00.000", timeZone: "UTC" });
    expect(result).toBeDefined();
    expect(result?.iso).toContain("2026-08-01T10:00:00");
    expect(result?.local.length).toBeGreaterThan(0);
  });
});
