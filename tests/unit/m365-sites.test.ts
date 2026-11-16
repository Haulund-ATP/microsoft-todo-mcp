import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";

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

const state = vi.hoisted(() => ({
  client: undefined as unknown,
  audits: [] as Array<Record<string, unknown>>,
  modules: [] as string[],
}));

vi.mock("../../src/tools/helpers.js", () => ({
  runGraphTool: async (_ctx: unknown, _name: string, _cid: unknown, fn: (c: unknown, id: string) => Promise<unknown>, module: string) => {
    state.modules.push(module);
    try {
      return { content: [{ type: "text", text: JSON.stringify(await fn(state.client, "conn-1")) }] };
    } catch (err) {
      return { content: [{ type: "text", text: JSON.stringify({ message: (err as Error).message }) }], isError: true };
    }
  },
}));
vi.mock("../../src/logging/audit.js", () => ({ audit: (e: Record<string, unknown>) => state.audits.push(e) }));

const sites = await import("../../src/graph/m365/sites.js");
const { registerSitesTools } = await import("../../src/tools/m365/sitesTools.js");

interface Call {
  path: string;
  method: string;
  query: Record<string, unknown>;
  headers: Record<string, string>;
  select?: string;
  expand?: string;
  filter?: string;
  orderby?: string;
  top?: number;
  body?: unknown;
}

function fakeGraph(respond: (call: Call) => unknown) {
  const calls: Call[] = [];
  const client = {
    api(path: string) {
      const call: Call = { path, method: "", query: {}, headers: {} };
      const send = async (method: string, body?: unknown) => {
        call.method = method;
        if (body !== undefined) call.body = body;
        calls.push(call);
        return respond(call);
      };
      const b = {
        select: (s: string) => ((call.select = s), b),
        expand: (s: string) => ((call.expand = s), b),
        filter: (s: string) => ((call.filter = s), b),
        orderby: (s: string) => ((call.orderby = s), b),
        top: (n: number) => ((call.top = n), b),
        query: (q: Record<string, unknown>) => (Object.assign(call.query, q), b),
        header: (k: string, v: string) => ((call.headers[k] = v), b),
        get: () => send("GET"),
        post: (body: unknown) => send("POST", body),
        patch: (body: unknown) => send("PATCH", body),
        delete: () => send("DELETE"),
      };
      return b;
    },
  };
  return { client: client as unknown as Client, calls };
}

const SITE_ID = "example.sharepoint.com,11111111-2222-3333-4444-555555555555,66666666-7777-8888-9999-000000000000";

/** Answers site-path lookups with SITE_ID and everything else with `rest`. */
function siteAware(rest: (call: Call) => unknown) {
  return (call: Call) => (call.path.includes(":/") && call.select === "id" ? { id: SITE_ID } : rest(call));
}

beforeEach(() => {
  state.audits.length = 0;
  state.modules.length = 0;
});

describe("site addressing", () => {
  it("accepts ids (commas kept), root, hostnames, host:/path and site URLs", () => {
    expect(sites.siteUrl(SITE_ID)).toEqual({ url: `/sites/${SITE_ID}`, byPath: false });
    expect(sites.siteUrl("root")).toEqual({ url: "/sites/root", byPath: false });
    expect(sites.siteUrl("example.sharepoint.com")).toEqual({ url: "/sites/example.sharepoint.com", byPath: false });
    expect(sites.siteUrl("example.sharepoint.com:/sites/Team Site")).toEqual({ url: "/sites/example.sharepoint.com:/sites/Team%20Site", byPath: true });
    expect(sites.siteUrl("https://example.sharepoint.com/sites/R%26D")).toEqual({ url: "/sites/example.sharepoint.com:/sites/R%26D", byPath: true });
    expect(sites.siteUrl("https://example.sharepoint.com/")).toEqual({ url: "/sites/example.sharepoint.com", byPath: false });
  });

  it("rejects bad hostnames, traversal and slashes in ids", () => {
    expect(() => sites.siteUrl("bad host:/sites/x")).toThrow(/hostname/);
    expect(() => sites.siteUrl("example.sharepoint.com:/sites/../x")).toThrow(/'\.\.'/);
    expect(() => sites.siteUrl("a/b")).toThrow(/Invalid id/);
    expect(() => sites.siteUrl(" ")).toThrow(/empty/);
  });

  it("resolves a site path to its id before addressing sub-resources", async () => {
    const { client, calls } = fakeGraph(siteAware(() => ({ value: [{ id: "b!lib", name: "Documents", driveType: "documentLibrary", webUrl: "https://example.test/d" }] })));
    const page = await sites.listSiteDrives(client, "example.sharepoint.com:/sites/team");
    expect(calls.map((c) => c.path)).toEqual(["/sites/example.sharepoint.com:/sites/team", `/sites/${SITE_ID}/drives`]);
    expect(page.items).toEqual([{ drive_id: "b!lib", name: "Documents", drive_type: "documentLibrary", web_url: "https://example.test/d" }]);
  });
});

describe("sites", () => {
  it("searches sites with the search parameter and a clamped page size", async () => {
    const { client, calls } = fakeGraph(() => ({ value: [{ id: SITE_ID, displayName: "Team", webUrl: "https://example.test/t", siteCollection: { hostname: "example.sharepoint.com" } }] }));
    const page = await sites.searchSites(client, { query: "team", top: 1000 });
    expect(calls[0]).toMatchObject({ path: "/sites", query: { search: "team" }, top: 100 });
    expect(page.items[0]).toEqual({ id: SITE_ID, display_name: "Team", web_url: "https://example.test/t", hostname: "example.sharepoint.com" });
  });

  it("gets a site by path directly and lists followed sites", async () => {
    const { client, calls } = fakeGraph(() => ({ id: SITE_ID, name: "team", value: [] }));
    await sites.getSite(client, "example.sharepoint.com:/sites/team");
    await sites.followedSites(client);
    expect(calls.map((c) => c.path)).toEqual(["/sites/example.sharepoint.com:/sites/team", "/me/followedSites"]);
  });

  it("hides hidden lists unless asked", async () => {
    const respond = () => ({ value: [{ id: "l1", displayName: "Opgaver", list: { template: "genericList" } }, { id: "l2", displayName: "System", list: { hidden: true } }] });
    const { client } = fakeGraph(respond);
    expect((await sites.listLists(client, SITE_ID)).items.map((l) => l.id)).toEqual(["l1"]);
    expect((await sites.listLists(client, SITE_ID, { includeHidden: true })).items.map((l) => l.id)).toEqual(["l1", "l2"]);
  });
});

describe("list items", () => {
  it("expands selected fields, filters with the non-indexed Prefer header and cleans OData noise", async () => {
    const { client, calls } = fakeGraph(() => ({
      value: [{ id: "7", fields: { "@odata.etag": "\"1\"", Title: "Kickoff", Status: "Open", LinkTitle: "Kickoff" }, lastModifiedBy: { user: { displayName: "A. Person" } } }],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/x/lists/y/items?$filter=fields/Status&$skiptoken=1",
    }));
    const page = await sites.listListItems(client, SITE_ID, "Project Tasks", {
      fields: ["Title", "Status"],
      filter: "fields/Status eq 'Open'",
      orderBy: "fields/Title desc",
      top: 10,
    });
    expect(calls[0]).toMatchObject({
      path: `/sites/${SITE_ID}/lists/Project%20Tasks/items`,
      expand: "fields($select=Title,Status)",
      filter: "fields/Status eq 'Open'",
      orderby: "fields/Title desc",
      top: 10,
      headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" },
    });
    expect(page.items[0]).toEqual({ id: "7", last_modified_by: "A. Person", fields: { Title: "Kickoff", Status: "Open" } });

    await sites.listListItems(client, SITE_ID, "x", { pageToken: page.next_page_token });
    expect(calls[1]!.headers.Prefer).toBe("HonorNonIndexedQueriesWarningMayFailRandomly");
  });

  it("rejects invalid column names and order_by values before calling Graph", async () => {
    const { client, calls } = fakeGraph(() => ({ value: [] }));
    await expect(sites.listListItems(client, SITE_ID, "l", { fields: ["Title)"] })).rejects.toThrow(/column name/);
    await expect(sites.listListItems(client, SITE_ID, "l", { orderBy: "fields/Title; drop" })).rejects.toThrow(/order_by/);
    await expect(sites.listListItems(client, SITE_ID, "a/b")).rejects.toThrow(/list/);
    expect(calls).toHaveLength(0);
  });

  it("creates, updates and deletes items with the documented shapes", async () => {
    const { client, calls } = fakeGraph((call) => (call.method === "POST" ? { id: "9", fields: { Title: "Ny" } } : call.method === "PATCH" ? { Title: "Rettet", "@odata.etag": "x" } : undefined));
    const created = await sites.createListItem(client, SITE_ID, "l1", { Title: "Ny", "Tags@odata.type": "Collection(Edm.String)", Tags: ["a", "b"] });
    const updated = await sites.updateListItem(client, SITE_ID, "l1", "9", { Title: "Rettet" });
    const deleted = await sites.deleteListItem(client, SITE_ID, "l1", "9");
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual([
      `POST /sites/${SITE_ID}/lists/l1/items`,
      `PATCH /sites/${SITE_ID}/lists/l1/items/9/fields`,
      `DELETE /sites/${SITE_ID}/lists/l1/items/9`,
    ]);
    expect(calls[0]!.body).toEqual({ fields: { Title: "Ny", "Tags@odata.type": "Collection(Edm.String)", Tags: ["a", "b"] } });
    expect(calls[1]!.body).toEqual({ Title: "Rettet" });
    expect(created).toMatchObject({ id: "9", fields: { Title: "Ny" } });
    expect(updated).toEqual({ id: "9", fields: { Title: "Rettet" } });
    expect(deleted).toMatchObject({ deleted: true, id: "9" });
    await expect(sites.createListItem(client, SITE_ID, "l1", {})).rejects.toThrow(/at least one/);
    await expect(sites.updateListItem(client, SITE_ID, "l1", "9", { "bad name": 1 })).rejects.toThrow(/column name/);
  });
});

describe("tool registration", () => {
  function register() {
    const tools = new Map<string, { handler: (args: Record<string, unknown>) => Promise<{ isError?: boolean }> }>();
    registerSitesTools({ registerTool: (name: string, _config: unknown, handler: never) => tools.set(name, { handler }) } as never, {
      profileAlias: "test",
      correlationId: "corr-1",
      boundConnectionId: "conn-1",
    });
    return tools;
  }

  it("registers the sites family on the sites module", async () => {
    const tools = register();
    expect([...tools.keys()].sort()).toEqual(
      [
        "sites_create_list_item",
        "sites_delete_list_item",
        "sites_followed",
        "sites_get",
        "sites_get_list_item",
        "sites_list_drives",
        "sites_list_items",
        "sites_list_lists",
        "sites_search",
        "sites_update_list_item",
      ].sort()
    );
    state.client = fakeGraph(() => ({ id: SITE_ID })).client;
    await tools.get("sites_get")!.handler({ site: "root" });
    expect(state.modules).toEqual(["sites"]);
  });

  it("audits list item deletes, and not failed ones", async () => {
    const tools = register();
    state.client = fakeGraph(() => undefined).client;
    await tools.get("sites_delete_list_item")!.handler({ site: SITE_ID, list: "l1", item_id: "9" });
    expect(state.audits).toEqual([expect.objectContaining({ action: "m365.delete", connectionId: "conn-1", detail: "list item delete count=1" })]);

    state.audits.length = 0;
    state.client = fakeGraph(() => {
      throw Object.assign(new Error("denied"), { statusCode: 403 });
    }).client;
    expect((await tools.get("sites_delete_list_item")!.handler({ site: SITE_ID, list: "l1", item_id: "9" })).isError).toBe(true);
    expect(state.audits).toEqual([]);
  });
});
