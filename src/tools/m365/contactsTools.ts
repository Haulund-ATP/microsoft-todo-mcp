import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "../context.js";
import { connectionIdInputShape } from "../accounts.js";
import { auditM365, runM365Tool, UNTRUSTED_CONTENT_NOTE } from "./common.js";
import * as contacts from "../../graph/m365/contacts.js";

const pageTokenField = z.string().optional().describe("next_page_token from the previous result, passed back unchanged.");
const topField = z
  .number()
  .int()
  .min(1)
  .max(contacts.MAX_PAGE_SIZE)
  .optional()
  .describe(`Page size (default ${contacts.DEFAULT_PAGE_SIZE}, max ${contacts.MAX_PAGE_SIZE}).`);

const addressShape = z.object({
  street: z.string().max(255).optional(),
  city: z.string().max(128).optional(),
  state: z.string().max(128).optional(),
  postal_code: z.string().max(32).optional(),
  country: z.string().max(128).optional(),
});

const contactFields = {
  given_name: z.string().max(255),
  surname: z.string().max(255),
  display_name: z.string().max(255).describe("Defaults to given name + surname."),
  emails: z
    .array(z.object({ address: z.string().email(), name: z.string().max(255).optional() }))
    .max(3)
    .describe("Up to 3 email addresses (replaces the existing list on update)."),
  mobile_phone: z.string().max(64),
  business_phones: z.array(z.string().max(64)).max(5),
  home_phones: z.array(z.string().max(64)).max(5),
  company: z.string().max(255),
  job_title: z.string().max(255),
  department: z.string().max(255),
  business_address: addressShape,
  home_address: addressShape,
  birthday: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "Use YYYY-MM-DD."),
  notes: z.string().max(10_000),
};

export const createContactInputShape = {
  given_name: contactFields.given_name.optional(),
  surname: contactFields.surname.optional(),
  display_name: contactFields.display_name.optional(),
  emails: contactFields.emails.optional(),
  mobile_phone: contactFields.mobile_phone.optional(),
  business_phones: contactFields.business_phones.optional(),
  home_phones: contactFields.home_phones.optional(),
  company: contactFields.company.optional(),
  job_title: contactFields.job_title.optional(),
  department: contactFields.department.optional(),
  business_address: contactFields.business_address.optional(),
  home_address: contactFields.home_address.optional(),
  birthday: contactFields.birthday.optional(),
  notes: contactFields.notes.optional(),
  ...connectionIdInputShape,
};

export const updateContactInputShape = {
  contact_id: z.string().min(1),
  given_name: contactFields.given_name.optional(),
  surname: contactFields.surname.optional(),
  display_name: contactFields.display_name.optional(),
  emails: contactFields.emails.optional(),
  mobile_phone: contactFields.mobile_phone.nullable().optional(),
  business_phones: contactFields.business_phones.optional(),
  home_phones: contactFields.home_phones.optional(),
  company: contactFields.company.nullable().optional(),
  job_title: contactFields.job_title.nullable().optional(),
  department: contactFields.department.nullable().optional(),
  business_address: contactFields.business_address.nullable().optional(),
  home_address: contactFields.home_address.nullable().optional(),
  birthday: contactFields.birthday.nullable().optional(),
  notes: contactFields.notes.nullable().optional(),
  ...connectionIdInputShape,
};

type ContactArgs = {
  given_name?: string;
  surname?: string;
  display_name?: string;
  emails?: Array<{ address: string; name?: string }>;
  mobile_phone?: string | null;
  business_phones?: string[];
  home_phones?: string[];
  company?: string | null;
  job_title?: string | null;
  department?: string | null;
  business_address?: contacts.ContactAddressInput | null;
  home_address?: contacts.ContactAddressInput | null;
  birthday?: string | null;
  notes?: string | null;
};

function toContactInput(a: ContactArgs): contacts.ContactInput {
  return {
    givenName: a.given_name,
    surname: a.surname,
    displayName: a.display_name,
    emails: a.emails,
    mobilePhone: a.mobile_phone,
    businessPhones: a.business_phones,
    homePhones: a.home_phones,
    companyName: a.company,
    jobTitle: a.job_title,
    department: a.department,
    businessAddress: a.business_address,
    homeAddress: a.home_address,
    birthday: a.birthday,
    notes: a.notes,
  };
}

export function registerContactsTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "contacts_list",
    {
      title: "List contacts",
      description: "Lists the signed-in user's personal Outlook contacts (default folder), sorted by name, with emails and phones." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { top: topField, page_token: pageTokenField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ top, page_token, connection_id }) =>
      runM365Tool(ctx, "contacts", "contacts_list", connection_id, (client) => contacts.listContacts(client, { top, pageToken: page_token }))
  );

  server.registerTool(
    "contacts_search",
    {
      title: "Search contacts",
      description:
        "Finds personal contacts by full email address (exact match) or by words in the name, company or email " +
        `(case-insensitive; scans up to ${contacts.MAX_SEARCH_SCAN} contacts). For colleagues and people you correspond ` +
        "with who are not saved as contacts, use people_search." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        query: z.string().min(1).max(255),
        limit: topField.describe(`Max results (default ${contacts.DEFAULT_PAGE_SIZE}, max ${contacts.MAX_PAGE_SIZE}).`),
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ query, limit, connection_id }) =>
      runM365Tool(ctx, "contacts", "contacts_search", connection_id, (client) => contacts.searchContacts(client, query, { limit }))
  );

  server.registerTool(
    "contacts_get",
    {
      title: "Get contact",
      description: "Gets one personal contact with all details (phones, addresses, birthday, notes)." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { contact_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ contact_id, connection_id }) =>
      runM365Tool(ctx, "contacts", "contacts_get", connection_id, (client) => contacts.getContact(client, contact_id))
  );

  server.registerTool(
    "contacts_create",
    {
      title: "Create contact",
      description: "Creates a personal contact in the default contacts folder. Needs at least a name, company or email.",
      inputSchema: createContactInputShape,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "contacts", "contacts_create", args.connection_id, (client) => contacts.createContact(client, toContactInput(args)))
  );

  server.registerTool(
    "contacts_update",
    {
      title: "Update contact",
      description: "Changes only the given fields of a personal contact; null clears a field. Lists (emails, phones) are replaced as a whole.",
      inputSchema: updateContactInputShape,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async ({ contact_id, connection_id, ...fields }) =>
      runM365Tool(ctx, "contacts", "contacts_update", connection_id, (client) =>
        contacts.updateContact(client, contact_id, toContactInput(fields))
      )
  );

  server.registerTool(
    "contacts_delete",
    {
      title: "Delete contact",
      description: "Deletes a personal contact (it goes to Deleted Items).",
      inputSchema: { contact_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
    },
    async ({ contact_id, connection_id }) =>
      runM365Tool(ctx, "contacts", "contacts_delete", connection_id, async (client, connectionId) => {
        const result = await contacts.deleteContact(client, contact_id);
        auditM365(ctx, "m365.delete", connectionId, "contact");
        return result;
      })
  );

  server.registerTool(
    "people_search",
    {
      title: "Search people",
      description:
        "People relevant to the signed-in user (colleagues, frequent correspondents, contacts), ranked by how much you " +
        "work with them; with a query, matched on name and email. Returns names and email addresses with relevance " +
        "scores — useful to resolve a name to an address before inviting or mailing someone." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        query: z.string().max(255).optional().describe("Name or email fragment; omit for the most relevant people."),
        top: topField,
        page_token: pageTokenField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async ({ query, top, page_token, connection_id }) =>
      runM365Tool(ctx, "contacts", "people_search", connection_id, (client) =>
        contacts.searchPeople(client, { query, top, pageToken: page_token })
      )
  );
}
