import { describe, expect, it } from "vitest";

// loadEnv() validates the full env schema; only PUBLIC_BASE_URL matters here.
process.env.NODE_ENV ??= "test";
process.env.PUBLIC_BASE_URL ??= "https://example.test";
process.env.AZURE_TENANT_ID ??= "test-tenant";
process.env.AZURE_CLIENT_ID ??= "test-client";
process.env.AZURE_KEY_VAULT_URI ??= "https://example-kv.test";
process.env.AZURE_STORAGE_CONNECTION_STRING ??= "UseDevelopmentStorage=true";
process.env.MCP_OAUTH_SIGNING_KEY_NAME ??= "test-signing-key";
process.env.MCP_OAUTH_ISSUER ??= "https://example.test";
process.env.ADMIN_OWNER_CLAIM_VALUE ??= "test-owner";
process.env.SESSION_COOKIE_SECRET ??= "test-session-secret-0123456789ab";
process.env.CSRF_COOKIE_SECRET ??= "test-csrf-secret-0123456789abcdef";

const { canonicalResource, protectedResourceMetadataUrl, resourceUrl } = await import("../../src/oauth/resources.js");
const { buildProtectedResourceMetadata } = await import("../../src/oauth/metadata.js");
const { renderConsentPage } = await import("../../src/oauth/authorize.js");

const base = process.env.PUBLIC_BASE_URL!.replace(/\/$/, "");

describe("protected resources", () => {
  it("keeps the To Do resource identical to what tokens were always issued for", () => {
    expect(resourceUrl("todo")).toBe(base);
    expect(buildProtectedResourceMetadata().resource).toBe(base);
    expect(protectedResourceMetadataUrl("todo")).toBe(`${base}/.well-known/oauth-protected-resource`);
  });

  it("gives Microsoft 365 its own resource, audience and metadata URL", () => {
    expect(resourceUrl("m365")).toBe(`${base}/m365`);
    expect(buildProtectedResourceMetadata("m365")).toMatchObject({ resource: `${base}/m365`, authorization_servers: [base] });
    expect(protectedResourceMetadataUrl("m365")).toBe(`${base}/.well-known/oauth-protected-resource/m365`);
  });

  it("maps client-supplied resources to the canonical one", () => {
    expect(canonicalResource(base)).toEqual({ kind: "todo", url: base });
    expect(canonicalResource(`${base}/`)).toEqual({ kind: "todo", url: base });
    expect(canonicalResource(`${base}/mcp/personal`)).toEqual({ kind: "todo", url: base });
    expect(canonicalResource(`${base}/m365`)).toEqual({ kind: "m365", url: `${base}/m365` });
    expect(canonicalResource(`${base}/m365/work/`)).toEqual({ kind: "m365", url: `${base}/m365` });
  });

  it("rejects other hosts and look-alike paths", () => {
    expect(canonicalResource("https://evil.test")).toBeUndefined();
    expect(canonicalResource(`${base}.evil.test`)).toBeUndefined();
    expect(canonicalResource(`${base}/m365evil`)).toBeUndefined();
    expect(canonicalResource(`${base}/accounts`)).toBeUndefined();
  });
});

describe("consent page per resource", () => {
  const params = {
    response_type: "code" as const, client_id: "c", redirect_uri: "https://claude.ai/cb", state: "s",
    code_challenge: "x".repeat(43), code_challenge_method: "S256" as const, resource: `${base}/m365`, scope: "mcp",
  };

  it("names Microsoft 365 and sending on the m365 resource", () => {
    const html = renderConsentPage("Claude", [], params, "csrf", "m365");
    expect(html).toContain("<strong>Microsoft 365</strong>");
    expect(html).toContain("sende");
  });

  it("still says To Do on the To Do resource", () => {
    const html = renderConsentPage("Claude", [], { ...params, resource: base }, "csrf");
    expect(html).toContain("Microsoft To Do-server");
    expect(html).not.toContain("Microsoft 365");
  });
});
