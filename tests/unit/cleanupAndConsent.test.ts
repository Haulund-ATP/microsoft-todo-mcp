import { describe, expect, it, vi } from "vitest";

const tables = new Map<string, Array<{ partitionKey: string; rowKey: string; expiresAt: string }>>();
const filters: string[] = [];

vi.mock("../../src/storage/clients.js", () => ({
  TABLE_NAMES: { oauthCodes: "oauthcodes", oauthTokens: "oauthtokens", pendingAuth: "pendingauth" },
  getTableClient: (name: string) => ({
    async *listEntities({ queryOptions }: { queryOptions: { filter: string } }) {
      filters.push(queryOptions.filter);
      const cutoff = /expiresAt lt '([^']+)'/.exec(queryOptions.filter)![1]!;
      for (const row of [...(tables.get(name) ?? [])]) if (row.expiresAt < cutoff) yield row;
    },
    async deleteEntity(pk: string, rk: string) {
      tables.set(name, (tables.get(name) ?? []).filter((r) => !(r.partitionKey === pk && r.rowKey === rk)));
    },
  }),
}));

const { deleteExpiredOAuthRows } = await import("../../src/storage/cleanup.js");
const { renderConsentPage, renderRedirectPage } = await import("../../src/oauth/authorize.js");

describe("deleteExpiredOAuthRows", () => {
  it("deletes rows expired more than a day ago from every one-time OAuth table", async () => {
    const now = Date.parse("2026-10-13T12:00:00.000Z");
    const row = (rowKey: string, expiresAt: string) => ({ partitionKey: "p", rowKey, expiresAt });
    tables.set("oauthcodes", [row("old", "2026-10-11T00:00:00.000Z"), row("recent", "2026-10-13T00:00:00.000Z")]);
    tables.set("oauthtokens", [row("old", "2026-09-01T00:00:00.000Z"), row("live", "2026-11-12T00:00:00.000Z")]);
    tables.set("pendingauth", [row("old", "2026-10-01T00:00:00.000Z")]);

    const deleted = await deleteExpiredOAuthRows(now);

    expect(deleted).toEqual({ oauthcodes: 1, oauthtokens: 1, pendingauth: 1 });
    expect(tables.get("oauthcodes")!.map((r) => r.rowKey)).toEqual(["recent"]);
    expect(tables.get("oauthtokens")!.map((r) => r.rowKey)).toEqual(["live"]);
    expect(filters[0]).toBe("expiresAt lt '2026-10-12T12:00:00.000Z'");
  });
});

describe("consent pages", () => {
  const params = {
    response_type: "code" as const, client_id: "c", redirect_uri: "https://chatgpt.com/cb", state: "s",
    code_challenge: "x".repeat(43), code_challenge_method: "S256" as const, resource: "https://todo.example", scope: "mcp",
  };

  it("names the client and the profiles the grant covers, escaping both", () => {
    const html = renderConsentPage("Chat<GPT>", ["nadia", "per<sonal"], params, "csrf");
    expect(html).toContain('lang="da"');
    expect(html).toContain("<strong>Chat&lt;GPT&gt;</strong>");
    expect(html).toContain("<strong>nadia</strong>, <strong>per&lt;sonal</strong>");
    expect(html).toContain(">Giv adgang</button>");
  });

  it("omits the profile line when none are bound", () => {
    expect(renderConsentPage("ChatGPT", [], params, "csrf")).not.toContain("profiler");
  });

  it("redirects via meta refresh with an escaped fallback link", () => {
    const html = renderRedirectPage('https://chatgpt.com/cb?code=a&state="b"');
    expect(html).toContain('content="0;url=https://chatgpt.com/cb?code=a&amp;state=&quot;b&quot;"');
    expect(html).toContain("kan du lukke denne fane");
  });
});
