import { describe, expect, it } from "vitest";

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

const { htmlToText, truncateText, graphSearchValue, pageOf, nextPageRequest } = await import("../../src/tools/m365/common.js");
const { adminConsentScopes, ALL_PERSONAL_MODULES, ALL_ORGANIZATIONAL_MODULES, MODULE_IDS } = await import("../../src/graph/modules.js");
const { buildTenantAdminConsentUrl, buildAdminConsentError, ADMIN_CONSENT_STATE_PREFIX } = await import("../../src/graph/upstreamOAuth.js");

describe("htmlToText", () => {
  it("keeps line structure and drops markup, scripts, styles and comments", () => {
    const html =
      "<html><head><style>p{color:red}</style></head><body><p>Hej <b>Thomas</b>,</p><!-- tracking --><div>Linje 2<br>Linje 3</div><script>alert(1)</script></body></html>";
    expect(htmlToText(html)).toBe("Hej Thomas,\nLinje 2\nLinje 3");
  });

  it("decodes entities and strips nested tag fragments until none are left", () => {
    expect(htmlToText("A&nbsp;&amp;&nbsp;B &lt;ok&gt; &#248;")).toBe("A & B <ok> ø");
    expect(htmlToText("<<b>script>x")).not.toContain("<b>");
  });
});

describe("text helpers", () => {
  it("truncates with a flag", () => {
    expect(truncateText("abcdef", 3)).toEqual({ text: "abc", truncated: true });
    expect(truncateText("abc", 3)).toEqual({ text: "abc", truncated: false });
  });

  it("quotes $search values safely", () => {
    expect(graphSearchValue(' budget "2027" ')).toBe('"budget 2027"');
  });
});

describe("paging", () => {
  it("exposes nextLink as an opaque token only when present", () => {
    expect(pageOf({ value: [1, 2] })).toEqual({ items: [1, 2] });
    expect(pageOf({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=25" })).toEqual({
      items: [],
      next_page_token: "https://graph.microsoft.com/v1.0/me/messages?$skip=25",
    });
  });

  it("only follows Microsoft Graph URLs, so a model-supplied token can never leak the bearer token", () => {
    const calls: string[] = [];
    const client = { api: (u: string) => (calls.push(u), {}) } as never;
    nextPageRequest(client, "https://graph.microsoft.com/v1.0/me/messages?$skip=25");
    expect(calls).toHaveLength(1);
    for (const bad of [
      "https://evil.example/steal",
      "https://graph.microsoft.com.evil.example/v1.0",
      "http://graph.microsoft.com/v1.0/me",
      "/me/messages",
      "https://graph.microsoft.com/v1.0/me messages",
    ]) {
      expect(() => nextPageRequest(client, bad)).toThrow(/next_page_token/);
    }
    expect(calls).toHaveLength(1);
  });
});

describe("tenant admin consent", () => {
  it("covers every module's Graph permission, fully qualified, without OIDC scopes", () => {
    const scopes = adminConsentScopes();
    expect(scopes).toContain("https://graph.microsoft.com/User.Read");
    expect(scopes).toContain("https://graph.microsoft.com/ChannelMessage.Read.All");
    expect(scopes).toContain("https://graph.microsoft.com/Sites.ReadWrite.All");
    expect(scopes.some((s) => /openid|offline_access|profile$/.test(s))).toBe(false);
    expect(new Set(scopes).size).toBe(scopes.length);
  });

  it("has presets for everything each account type supports", () => {
    expect(ALL_ORGANIZATIONAL_MODULES).toEqual([...MODULE_IDS]);
    expect(ALL_PERSONAL_MODULES).toEqual(["todo", "mail", "calendar", "contacts", "files"]);
  });

  it("builds a single v2 admin consent URL for the tenant", () => {
    const url = new URL(buildTenantAdminConsentUrl("contoso.com", "ac.state1"));
    expect(url.origin + url.pathname).toBe("https://login.microsoftonline.com/contoso.com/v2.0/adminconsent");
    expect(url.searchParams.get("client_id")).toBe(process.env.AZURE_CLIENT_ID);
    expect(url.searchParams.get("redirect_uri")).toBe("https://example.test/oauth/microsoft/callback");
    expect(url.searchParams.get("state")).toBe("ac.state1");
    expect(url.searchParams.get("scope")!.split(" ")).toEqual(adminConsentScopes());
  });

  it("points consent-required errors at the same all-modules URL, recognisable by its state prefix", () => {
    const info = buildAdminConsentError("11111111-2222-3333-4444-555555555555", "AADSTS65001");
    const url = new URL(info.adminConsentUrl);
    expect(url.pathname).toBe("/11111111-2222-3333-4444-555555555555/v2.0/adminconsent");
    expect(url.searchParams.get("state")!.startsWith(ADMIN_CONSENT_STATE_PREFIX)).toBe(true);
  });
});
