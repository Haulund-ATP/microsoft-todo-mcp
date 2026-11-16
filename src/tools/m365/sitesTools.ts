import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "../context.js";
import { connectionIdInputShape } from "../accounts.js";
import { auditM365, runM365Tool, UNTRUSTED_CONTENT_NOTE } from "./common.js";
import * as sites from "../../graph/m365/sites.js";

/**
 * SharePoint tool family (module "sites", Sites.ReadWrite.All, work/school
 * accounts only): sites, document libraries and lists. Files inside a
 * library are handled by the files_* tools with the library's drive_id.
 */

const siteField = z
  .string()
  .min(1)
  .max(1024)
  .describe(
    "The site: its id (from sites_search), 'root', a hostname such as 'contoso.sharepoint.com', 'hostname:/sites/path', or the site's https URL."
  );
const listField = z.string().min(1).max(255).describe("The list's id (from sites_list_lists) or its display name.");
const topField = z.number().int().min(1).max(sites.MAX_PAGE_SIZE).optional().describe(`Page size (default ${sites.DEFAULT_PAGE_SIZE}, max ${sites.MAX_PAGE_SIZE}).`);
const pageTokenField = z.string().optional().describe("next_page_token from the previous result, passed back unchanged.");
const fieldValue = z.union([z.string(), z.number(), z.boolean(), z.null(), z.array(z.string()), z.array(z.number())]);
const fieldsField = z
  .record(fieldValue)
  .describe(
    "Column values keyed by internal column name, e.g. {\"Title\": \"Kickoff\", \"Status\": \"Open\"}. Lookup and person columns use '<Name>LookupId'; multi-value columns also need '<Name>@odata.type': 'Collection(Edm.String)'."
  );

export function registerSitesTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "sites_search",
    {
      title: "Search SharePoint sites",
      description: "Searches SharePoint sites you can access by name or description ('*' lists all). Work/school accounts only." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { query: z.string().min(1).max(200), top: topField, page_token: pageTokenField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_search", args.connection_id, (client) =>
        sites.searchSites(client, { query: args.query, top: args.top, pageToken: args.page_token })
      )
  );

  server.registerTool(
    "sites_followed",
    {
      title: "Followed SharePoint sites",
      description: "Lists the SharePoint sites you follow." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { page_token: pageTokenField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) => runM365Tool(ctx, "sites", "sites_followed", args.connection_id, (client) => sites.followedSites(client, { pageToken: args.page_token }))
  );

  server.registerTool(
    "sites_get",
    {
      title: "Get a SharePoint site",
      description: "Returns a site's id, name, description and web_url, addressed by id, hostname, 'hostname:/sites/path' or URL." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { site: siteField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) => runM365Tool(ctx, "sites", "sites_get", args.connection_id, (client) => sites.getSite(client, args.site))
  );

  server.registerTool(
    "sites_list_drives",
    {
      title: "List document libraries",
      description:
        "Lists a site's document libraries. Pass a library's drive_id to files_list_children, files_search, files_read_content and the other " +
        "files_* tools to work with its files.",
      inputSchema: { site: siteField, page_token: pageTokenField, ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_list_drives", args.connection_id, (client) => sites.listSiteDrives(client, args.site, { pageToken: args.page_token }))
  );

  server.registerTool(
    "sites_list_lists",
    {
      title: "List SharePoint lists",
      description: "Lists a site's lists (and libraries, which are lists too). Hidden system lists are left out unless include_hidden is true." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        site: siteField,
        include_hidden: z.boolean().default(false),
        top: topField,
        page_token: pageTokenField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_list_lists", args.connection_id, (client) =>
        sites.listLists(client, args.site, { includeHidden: args.include_hidden, top: args.top, pageToken: args.page_token })
      )
  );

  server.registerTool(
    "sites_list_items",
    {
      title: "List SharePoint list items",
      description:
        "Lists items in a SharePoint list with their column values. fields limits which columns are returned (internal names). filter is an " +
        "OData filter on columns, e.g. \"fields/Status eq 'Open'\" (filtering non-indexed columns can fail on very large lists). order_by e.g. " +
        "'fields/Title' or 'lastModifiedDateTime desc'." +
        UNTRUSTED_CONTENT_NOTE,
      inputSchema: {
        site: siteField,
        list: listField,
        fields: z.array(z.string().min(1).max(255)).max(50).optional().describe("Internal column names to return; all columns when omitted."),
        filter: z.string().max(1000).optional(),
        order_by: z.string().max(200).optional(),
        top: topField,
        page_token: pageTokenField,
        ...connectionIdInputShape,
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_list_items", args.connection_id, (client) =>
        sites.listListItems(client, args.site, args.list, {
          fields: args.fields,
          filter: args.filter,
          orderBy: args.order_by,
          top: args.top,
          pageToken: args.page_token,
        })
      )
  );

  server.registerTool(
    "sites_get_list_item",
    {
      title: "Get a SharePoint list item",
      description: "Returns one list item with all its column values." + UNTRUSTED_CONTENT_NOTE,
      inputSchema: { site: siteField, list: listField, item_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_get_list_item", args.connection_id, (client) => sites.getListItem(client, args.site, args.list, args.item_id))
  );

  server.registerTool(
    "sites_create_list_item",
    {
      title: "Create a SharePoint list item",
      description: "Adds an item to a SharePoint list with the given column values.",
      inputSchema: { site: siteField, list: listField, fields: fieldsField, ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_create_list_item", args.connection_id, (client) =>
        sites.createListItem(client, args.site, args.list, args.fields)
      )
  );

  server.registerTool(
    "sites_update_list_item",
    {
      title: "Update a SharePoint list item",
      description: "Updates the given column values of a list item; other columns are left unchanged.",
      inputSchema: { site: siteField, list: listField, item_id: z.string().min(1), fields: fieldsField, ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_update_list_item", args.connection_id, (client) =>
        sites.updateListItem(client, args.site, args.list, args.item_id, args.fields)
      )
  );

  server.registerTool(
    "sites_delete_list_item",
    {
      title: "Delete a SharePoint list item",
      description: "Deletes a list item. It goes to the site recycle bin, from which a site owner can restore it. Confirm with the user before deleting.",
      inputSchema: { site: siteField, list: listField, item_id: z.string().min(1), ...connectionIdInputShape },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    },
    async (args) =>
      runM365Tool(ctx, "sites", "sites_delete_list_item", args.connection_id, async (client, connectionId) => {
        const result = await sites.deleteListItem(client, args.site, args.list, args.item_id);
        auditM365(ctx, "m365.delete", connectionId, "list item delete count=1");
        return result;
      })
  );
}
