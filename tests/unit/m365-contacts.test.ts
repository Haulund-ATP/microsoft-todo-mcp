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

import * as contacts from "../../src/graph/m365/contacts.js";
import { registerContactsTools } from "../../src/tools/m365/contactsTools.js";
import type { ToolContext } from "../../src/tools/context.js";

interface Call {
  path: string;
  method?: string;
  headers: Record<string, string>;
  select?: string;
  top?: number;
  orderby?: string;
  filter?: string;
  search?: string;
  body?: unknown;
}

function fakeGraph(respond: (call: Call) => unknown = () => ({})) {
  const calls: Call[] = [];
  const client = {
    api(path: string) {
      const call: Call = { path, headers: {} };
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
        query: () => builder,
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

const anna = {
  id: "c1",
  displayName: "Anna Jensen",
  givenName: "Anna",
  surname: "Jensen",
  companyName: "Example A/S",
  emailAddresses: [{ address: "anna@example.com", name: "Anna Jensen" }],
  mobilePhone: "+45 11 22 33 44",
  businessPhones: [],
};
const bo = { id: "c2", displayName: "Bo Hansen", givenName: "Bo", surname: "Hansen", emailAddresses: [{ address: "bo@example.com" }] };

describe("contacts reads", () => {
  it("lists contacts with $select, name order and the default page size", async () => {
    const g = fakeGraph(() => ({ value: [anna], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$skip=25" }));
    const result = await contacts.listContacts(g.client);
    expect(g.calls[0]).toMatchObject({ path: "/me/contacts", orderby: "displayName", top: 25 });
    expect(g.calls[0]!.select).toContain("emailAddresses");
    expect(g.calls[0]!.select).not.toContain("personalNotes");
    expect(result.contacts[0]).toEqual({
      id: "c1",
      display_name: "Anna Jensen",
      given_name: "Anna",
      surname: "Jensen",
      company: "Example A/S",
      job_title: undefined,
      emails: ["anna@example.com"],
      mobile_phone: "+45 11 22 33 44",
      business_phones: undefined,
    });
    expect(result.next_page_token).toContain("$skip=25");
  });

  it("searches a full email address with Graph's supported $filter", async () => {
    const g = fakeGraph(() => ({ value: [anna] }));
    const result = await contacts.searchContacts(g.client, "Anna@Example.com");
    expect(g.calls).toHaveLength(1);
    expect(g.calls[0]!.filter).toBe("emailAddresses/any(a:a/address eq 'anna@example.com')");
    expect(result.method).toBe("email_filter");
    expect(result.contacts.map((c) => c.id)).toEqual(["c1"]);
  });

  it("matches every word of a name query client-side across pages", async () => {
    const next = "https://graph.microsoft.com/v1.0/me/contacts?$skip=100";
    const g = fakeGraph((call) => (call.path === next ? { value: [anna] } : { value: [bo], "@odata.nextLink": next }));
    const result = await contacts.searchContacts(g.client, "jensen ANNA");
    expect(g.calls.map((c) => c.path)).toEqual(["/me/contacts", next]);
    expect(g.calls[0]!.top).toBe(100);
    expect(result).toMatchObject({ method: "name_scan", scanned: 2 });
    expect(result.contacts.map((c) => c.id)).toEqual(["c1"]);
  });

  it("stops scanning once enough contacts matched", async () => {
    const g = fakeGraph(() => ({ value: [anna, { ...anna, id: "c3" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$skip=100" }));
    const result = await contacts.searchContacts(g.client, "anna", { limit: 1 });
    expect(g.calls).toHaveLength(1);
    expect(result.contacts).toHaveLength(1);
  });

  it("returns a contact's details with dates trimmed", async () => {
    const g = fakeGraph(() => ({ ...anna, birthday: "1990-05-17T11:59:00Z", personalNotes: "  Met at the fair  ", homeAddress: { city: "Aarhus", street: "" } }));
    const detail = await contacts.getContact(g.client, "c/1");
    expect(g.calls[0]!.path).toBe("/me/contacts/c%2F1");
    expect(detail).toMatchObject({ birthday: "1990-05-17", notes: "Met at the fair", home_address: { city: "Aarhus" } });
  });
});

describe("contact writes", () => {
  it("builds payloads with only the given fields, null clearing them", () => {
    expect(
      contacts.buildContactPayload({
        givenName: "Anna",
        emails: [{ address: "anna@example.com" }],
        businessAddress: { city: "Aarhus", postal_code: "8000", country: "Denmark" },
        birthday: "1990-05-17",
        jobTitle: null,
        notes: null,
      })
    ).toEqual({
      givenName: "Anna",
      emailAddresses: [{ address: "anna@example.com" }],
      businessAddress: { city: "Aarhus", postalCode: "8000", countryOrRegion: "Denmark" },
      birthday: "1990-05-17T12:00:00Z",
      jobTitle: "",
      personalNotes: "",
    });
    expect(() => contacts.buildContactPayload({ emails: [1, 2, 3, 4].map((n) => ({ address: `p${n}@example.com` })) })).toThrow(/at most 3/);
    expect(() => contacts.buildContactPayload({ birthday: "17-05-1990" })).toThrow(/YYYY-MM-DD/);
  });

  it("creates contacts and refuses empty ones", async () => {
    const g = fakeGraph((call) => ({ id: "new", ...(call.body as object) }));
    await expect(contacts.createContact(g.client, { notes: "x" })).rejects.toThrow(/name, a company or an email/);
    const created = await contacts.createContact(g.client, { givenName: "Bo", surname: "Hansen" });
    expect(g.calls[0]).toMatchObject({ method: "POST", path: "/me/contacts", body: { givenName: "Bo", surname: "Hansen" } });
    expect(created.id).toBe("new");
  });

  it("keeps the display name on updates that do not touch the name", async () => {
    const g = fakeGraph((call) => (call.method === "GET" ? { id: "c1", displayName: "Anna J." } : { id: "c1" }));
    await contacts.updateContact(g.client, "c1", { mobilePhone: "+45 99" });
    expect(g.calls.map((c) => c.method)).toEqual(["GET", "PATCH"]);
    expect(g.calls[1]!.body).toEqual({ mobilePhone: "+45 99", displayName: "Anna J." });

    const renamed = fakeGraph(() => ({ id: "c1" }));
    await contacts.updateContact(renamed.client, "c1", { surname: "Holm" });
    expect(renamed.calls.map((c) => c.method)).toEqual(["PATCH"]);
    await expect(contacts.updateContact(renamed.client, "c1", {})).rejects.toThrow(/Nothing to update/);
  });
});

describe("people", () => {
  it("searches /me/people with a quoted $search and returns scored addresses", async () => {
    const g = fakeGraph(() => ({
      value: [
        {
          id: "p1",
          displayName: "Anna Jensen",
          scoredEmailAddresses: [{ address: "anna@example.com", relevanceScore: 8 }, { relevanceScore: 1 }],
          personType: { class: "Person", subclass: "OrganizationUser" },
        },
      ],
    }));
    const result = await contacts.searchPeople(g.client, { query: 'Anna "J"', top: 5 });
    expect(g.calls[0]).toMatchObject({ path: "/me/people", search: '"Anna J"', top: 5 });
    expect(g.calls[0]!.select).toContain("scoredEmailAddresses");
    expect(result.people).toEqual([
      { id: "p1", display_name: "Anna Jensen", emails: [{ address: "anna@example.com", score: 8 }], job_title: undefined, company: undefined, department: undefined, kind: "OrganizationUser" },
    ]);
  });

  it("lists the most relevant people when no query is given", async () => {
    const g = fakeGraph(() => ({ value: [] }));
    await contacts.searchPeople(g.client);
    expect(g.calls[0]!.search).toBeUndefined();
    expect(g.calls[0]!.top).toBe(25);
  });
});

describe("contacts tools", () => {
  type Handler = (args: Record<string, unknown>) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;
  function register() {
    const tools = new Map<string, Handler>();
    const server = { registerTool: (name: string, _config: unknown, handler: Handler) => tools.set(name, handler) } as unknown as McpServer;
    const ctx: ToolContext = { profileAlias: "work", boundConnectionId: "conn-1", correlationId: "corr", modules: ["contacts"], allowSend: false };
    registerContactsTools(server, ctx);
    return tools;
  }

  beforeEach(() => {
    shared.audits.length = 0;
  });

  it("registers the contacts_ and people_ tools", () => {
    expect([...register().keys()].sort()).toEqual([
      "contacts_create",
      "contacts_delete",
      "contacts_get",
      "contacts_list",
      "contacts_search",
      "contacts_update",
      "people_search",
    ]);
  });

  it("audits a contact delete without its content", async () => {
    const g = fakeGraph(() => undefined);
    shared.client = g.client;
    const result = await register().get("contacts_delete")!({ contact_id: "c1" });
    expect(result.isError).toBeUndefined();
    expect(g.calls[0]).toMatchObject({ method: "DELETE", path: "/me/contacts/c1" });
    expect(shared.audits).toEqual([expect.objectContaining({ action: "m365.delete", connectionId: "conn-1", detail: "contact" })]);
  });

  it("does not audit a failed delete", async () => {
    shared.client = fakeGraph(() => {
      throw Object.assign(new Error("gone"), { statusCode: 404 });
    }).client;
    const result = await register().get("contacts_delete")!({ contact_id: "c1" });
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0]!.text).errorType).toBe("not_found");
    expect(shared.audits).toEqual([]);
  });

  it("maps snake_case update arguments onto the Graph contact", async () => {
    const g = fakeGraph(() => ({ id: "c1", displayName: "Anna" }));
    shared.client = g.client;
    await register().get("contacts_update")!({ contact_id: "c1", given_name: "Anna", company: null });
    expect(g.calls[0]).toMatchObject({ method: "PATCH", body: { givenName: "Anna", companyName: "" } });
  });
});
