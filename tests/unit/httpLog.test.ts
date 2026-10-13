import { describe, expect, it } from "vitest";
import { httpLogSerializers, isProbeRequest, redactUrl } from "../../src/logging/httpLog.js";
import { categorizeBearerError } from "../../src/oauth/requireBearer.js";
import type { IncomingMessage } from "node:http";

describe("redactUrl", () => {
  it("redacts the Microsoft sign-in code and account identifiers, keeping other params", () => {
    const url = "/oauth/owner/callback?code=1.ATsACT772&client_info=eyJ1aWQi&state=abc&session_state=00aa";
    const redacted = redactUrl(url)!;
    expect(redacted).not.toContain("1.ATsACT772");
    expect(redacted).not.toContain("eyJ1aWQi");
    expect(redacted).not.toContain("00aa");
    expect(redacted).toContain("state=abc");
    expect(redacted.startsWith("/oauth/owner/callback?")).toBe(true);
  });

  it("leaves URLs without sensitive params untouched", () => {
    expect(redactUrl("/oauth/authorize?client_id=x&scope=mcp")).toBe("/oauth/authorize?client_id=x&scope=mcp");
    expect(redactUrl("/mcp/nadia")).toBe("/mcp/nadia");
    expect(redactUrl(undefined)).toBeUndefined();
  });
});

describe("request log shape", () => {
  it("skips health and readiness probes only", () => {
    const req = (url: string) => ({ url }) as IncomingMessage;
    expect(isProbeRequest(req("/health"))).toBe(true);
    expect(isProbeRequest(req("/ready?x=1"))).toBe(true);
    expect(isProbeRequest(req("/mcp/health"))).toBe(false);
  });

  it("logs method, redacted url, user agent and client ip — never headers or cookies", () => {
    const serialized = httpLogSerializers.req({
      id: "c1",
      method: "GET",
      url: "/oauth/owner/callback?code=secret",
      headers: { "user-agent": "ua", cookie: "s=1", authorization: "Bearer t", "x-envoy-external-address": "1.2.3.4" },
    });
    expect(serialized).toEqual({ id: "c1", method: "GET", url: "/oauth/owner/callback?code=%5Bredacted%5D", userAgent: "ua", clientIp: "1.2.3.4" });
    expect(httpLogSerializers.res({ statusCode: 302, getHeaders: () => ({ "set-cookie": "x" }) } as never)).toEqual({ statusCode: 302 });
  });
});

describe("categorizeBearerError", () => {
  it.each([
    [{ code: "ERR_JWT_EXPIRED" }, "bearer_expired"],
    [{ code: "ERR_JWT_CLAIM_VALIDATION_FAILED", claim: "aud" }, "bearer_audience_mismatch"],
    [{ code: "ERR_JWT_CLAIM_VALIDATION_FAILED", claim: "iss" }, "bearer_issuer_mismatch"],
    [{ code: "ERR_JWS_SIGNATURE_VERIFICATION_FAILED" }, "bearer_bad_signature"],
    [{ code: "ERR_JWS_INVALID" }, "bearer_malformed"],
    [new Error("key vault down"), "bearer_invalid"],
  ])("%o -> %s", (err, category) => {
    expect(categorizeBearerError(err)).toBe(category);
  });
});
