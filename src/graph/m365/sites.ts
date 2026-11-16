import type { Client } from "@microsoft/microsoft-graph-client";
import { nextPageRequest, pageOf, type Page } from "../../tools/m365/common.js";
import { encodeId } from "./files.js";

/**
 * SharePoint sites and lists through Graph v1.0, delegated
 * (`Sites.ReadWrite.All`, work/school accounts only). Endpoint notes
 * (verified against Microsoft Learn):
 *
 * - `GET /sites?search={q}` searches sites; `GET /me/followedSites` lists the
 *   sites the user follows.
 * - A site is addressed by id (`hostname,siteCollectionId,webId`), by
 *   `/sites/root`, by hostname (that host's root site) or by server-relative
 *   path `/sites/{hostname}:/sites/{path}`. Sub-resources are addressed via
 *   the id, so a path is resolved to an id first.
 * - `GET /sites/{id}/drives` lists document libraries; their ids work with
 *   every files_* tool as drive_id.
 * - `GET /sites/{id}/lists/{list}/items?$expand=fields(...)` lists items;
 *   `$filter` on fields that are not indexed needs the header
 *   `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`.
 * - `POST .../items` with `{ fields }` creates, `PATCH .../items/{id}/fields`
 *   updates and `DELETE .../items/{id}` deletes (to the site recycle bin).
 */

export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;

const SITE_SELECT = "id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,siteCollection";
const DRIVE_SELECT = "id,name,description,webUrl,driveType,createdDateTime,lastModifiedDateTime";
const LIST_SELECT = "id,name,displayName,description,webUrl,list,createdDateTime,lastModifiedDateTime";
const ITEM_SELECT = "id,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,contentType";

function clampTop(top: number | undefined): number {
  return Math.min(Math.max(top ?? DEFAULT_PAGE_SIZE, 1), MAX_PAGE_SIZE);
}

function dropUndefined<T extends Record<string, unknown>>(obj: T): T {
  for (const k of Object.keys(obj)) if (obj[k] === undefined) delete obj[k];
  return obj;
}

// ---------------------------------------------------------------------------
// Site addressing

const HOST = /^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$/;

function encodePathSegments(path: string): string {
  const segments = path.split("/").filter((s) => s !== "");
  for (const s of segments) if (s === "." || s === "..") throw new Error("Site paths may not contain '.' or '..' segments.");
  return segments.map((s) => encodeURIComponent(s)).join("/");
}

/**
 * Turns the model's site reference into a Graph URL. Accepts a site id,
 * "root", a hostname, "hostname:/sites/path" or a full https URL of the site.
 * `byPath` is true when Graph must resolve a server-relative path.
 */
export function siteUrl(site: string): { url: string; byPath: boolean } {
  const s = site.trim();
  if (s === "") throw new Error("site must not be empty.");
  if (s === "root") return { url: "/sites/root", byPath: false };
  let host: string | undefined;
  let path: string | undefined;
  if (/^https:\/\//i.test(s)) {
    const parsed = new URL(s);
    host = parsed.hostname;
    path = decodeURIComponent(parsed.pathname);
  } else {
    const colon = s.indexOf(":");
    if (colon > 0 && !s.includes(",")) {
      host = s.slice(0, colon);
      path = s.slice(colon + 1);
    } else if (HOST.test(s)) {
      host = s;
      path = "";
    }
  }
  if (host !== undefined) {
    if (!HOST.test(host)) throw new Error(`Invalid SharePoint hostname: ${JSON.stringify(host)}`);
    const encoded = encodePathSegments(path ?? "");
    return encoded === "" ? { url: `/sites/${host}`, byPath: false } : { url: `/sites/${host}:/${encoded}`, byPath: true };
  }
  return { url: `/sites/${encodeId(s)}`, byPath: false };
}

/** Base URL for a site's sub-resources; path references are resolved to the site id first. */
export async function siteBase(client: Client, site: string): Promise<string> {
  const { url, byPath } = siteUrl(site);
  if (!byPath) return url;
  const resolved = (await client.api(url).select("id").get()) as { id?: string };
  if (!resolved.id) throw new Error("The site could not be resolved.");
  return `/sites/${encodeId(resolved.id)}`;
}

// ---------------------------------------------------------------------------
// Shapes

interface GraphIdentitySet {
  user?: { displayName?: string | null } | null;
  application?: { displayName?: string | null } | null;
}

interface GraphSite {
  id?: string;
  name?: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  siteCollection?: { hostname?: string } | null;
}

export interface SiteSummary {
  id?: string;
  name?: string;
  display_name?: string;
  description?: string;
  web_url?: string;
  hostname?: string;
  last_modified?: string;
}

export function summarizeSite(s: GraphSite): SiteSummary {
  return dropUndefined({
    id: s.id,
    name: s.name,
    display_name: s.displayName,
    description: s.description || undefined,
    web_url: s.webUrl,
    hostname: s.siteCollection?.hostname ?? undefined,
    last_modified: s.lastModifiedDateTime,
  });
}

function sitePage(response: { value?: GraphSite[]; "@odata.nextLink"?: string }): Page<SiteSummary> {
  const page = pageOf(response);
  return { ...page, items: page.items.map(summarizeSite) };
}

function who(set: GraphIdentitySet | undefined | null): string | undefined {
  return set?.user?.displayName ?? set?.application?.displayName ?? undefined;
}

// ---------------------------------------------------------------------------
// Sites

export async function searchSites(client: Client, opts: { query: string; top?: number; pageToken?: string }): Promise<Page<SiteSummary>> {
  if (opts.pageToken) return sitePage(await nextPageRequest(client, opts.pageToken).get());
  const query = opts.query.trim() || "*";
  return sitePage(await client.api("/sites").query({ search: query }).select(SITE_SELECT).top(clampTop(opts.top)).get());
}

export async function followedSites(client: Client, opts: { pageToken?: string } = {}): Promise<Page<SiteSummary>> {
  if (opts.pageToken) return sitePage(await nextPageRequest(client, opts.pageToken).get());
  return sitePage(await client.api("/me/followedSites").select(SITE_SELECT).get());
}

export async function getSite(client: Client, site: string): Promise<SiteSummary> {
  return summarizeSite(await client.api(siteUrl(site).url).select(SITE_SELECT).get());
}

export interface DriveSummary {
  drive_id?: string;
  name?: string;
  description?: string;
  web_url?: string;
  drive_type?: string;
  last_modified?: string;
}

export async function listSiteDrives(client: Client, site: string, opts: { pageToken?: string } = {}): Promise<Page<DriveSummary>> {
  const response = opts.pageToken
    ? await nextPageRequest(client, opts.pageToken).get()
    : await client.api(`${await siteBase(client, site)}/drives`).select(DRIVE_SELECT).get();
  const page = pageOf<{ id?: string; name?: string; description?: string; webUrl?: string; driveType?: string; lastModifiedDateTime?: string }>(response);
  return {
    ...page,
    items: page.items.map((d) =>
      dropUndefined({
        drive_id: d.id,
        name: d.name,
        description: d.description || undefined,
        web_url: d.webUrl,
        drive_type: d.driveType,
        last_modified: d.lastModifiedDateTime,
      })
    ),
  };
}

// ---------------------------------------------------------------------------
// Lists

export interface ListSummary {
  id?: string;
  name?: string;
  display_name?: string;
  description?: string;
  web_url?: string;
  template?: string;
  hidden?: boolean;
  last_modified?: string;
}

interface GraphList {
  id?: string;
  name?: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  list?: { template?: string; hidden?: boolean } | null;
}

export async function listLists(
  client: Client,
  site: string,
  opts: { includeHidden?: boolean; top?: number; pageToken?: string } = {}
): Promise<Page<ListSummary>> {
  const response = opts.pageToken
    ? await nextPageRequest(client, opts.pageToken).get()
    : await client.api(`${await siteBase(client, site)}/lists`).select(LIST_SELECT).top(clampTop(opts.top)).get();
  const page = pageOf<GraphList>(response);
  return {
    ...page,
    items: page.items
      .filter((l) => opts.includeHidden || !l.list?.hidden)
      .map((l) =>
        dropUndefined({
          id: l.id,
          name: l.name,
          display_name: l.displayName,
          description: l.description || undefined,
          web_url: l.webUrl,
          template: l.list?.template,
          hidden: l.list?.hidden || undefined,
          last_modified: l.lastModifiedDateTime,
        })
      ),
  };
}

const FIELD_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;

export function validateFieldNames(fields: readonly string[]): string[] {
  for (const f of fields) if (!FIELD_NAME.test(f)) throw new Error(`Invalid field (column) name: ${JSON.stringify(f)}. Use the internal column name.`);
  return [...fields];
}

export interface ListItemSummary {
  id?: string;
  web_url?: string;
  content_type?: string;
  created?: string;
  last_modified?: string;
  created_by?: string;
  last_modified_by?: string;
  fields?: Record<string, unknown>;
}

interface GraphListItem {
  id?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: GraphIdentitySet;
  lastModifiedBy?: GraphIdentitySet;
  contentType?: { name?: string } | null;
  fields?: Record<string, unknown> | null;
}

/** Drops OData annotations and SharePoint bookkeeping columns that only add noise. */
const NOISE_FIELDS = new Set(["_UIVersionString", "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId", "AppAuthorLookupId", "AppEditorLookupId", "ItemChildCount", "FolderChildCount", "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "ContentType"]);

export function cleanFields(fields: Record<string, unknown> | null | undefined): Record<string, unknown> | undefined {
  if (!fields) return undefined;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(fields)) {
    if (k.startsWith("@odata") || k.includes("@odata.") || NOISE_FIELDS.has(k)) continue;
    out[k] = v;
  }
  return out;
}

export function summarizeListItem(i: GraphListItem): ListItemSummary {
  return dropUndefined({
    id: i.id,
    web_url: i.webUrl,
    content_type: i.contentType?.name ?? undefined,
    created: i.createdDateTime,
    last_modified: i.lastModifiedDateTime,
    created_by: who(i.createdBy),
    last_modified_by: who(i.lastModifiedBy),
    fields: cleanFields(i.fields),
  });
}

/** A list is addressed by id (GUID) or display name; names may contain spaces, so only slashes are refused. */
export function encodeListRef(list: string): string {
  const trimmed = list.trim();
  if (trimmed === "" || /[/\\]/.test(trimmed)) throw new Error(`Invalid list id or name: ${JSON.stringify(list)}`);
  return encodeURIComponent(trimmed);
}

function listBase(siteBaseUrl: string, list: string): string {
  return `${siteBaseUrl}/lists/${encodeListRef(list)}`;
}

export async function listListItems(
  client: Client,
  site: string,
  list: string,
  opts: { fields?: string[]; filter?: string; orderBy?: string; top?: number; pageToken?: string } = {}
): Promise<Page<ListItemSummary>> {
  let response;
  if (opts.pageToken) {
    let req = nextPageRequest(client, opts.pageToken);
    if (opts.filter || /[?&](\$|%24)filter=/i.test(opts.pageToken)) req = req.header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
    response = await req.get();
  } else {
    const fields = opts.fields?.length ? validateFieldNames(opts.fields) : undefined;
    let req = client
      .api(`${listBase(await siteBase(client, site), list)}/items`)
      .select(ITEM_SELECT)
      .expand(fields ? `fields($select=${fields.join(",")})` : "fields")
      .top(clampTop(opts.top));
    if (opts.filter) {
      if (opts.filter.length > 1000) throw new Error("filter is too long.");
      req = req.filter(opts.filter).header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
    }
    if (opts.orderBy) {
      if (!/^fields\/[A-Za-z_][A-Za-z0-9_]*( (asc|desc))?$|^(createdDateTime|lastModifiedDateTime)( (asc|desc))?$/.test(opts.orderBy)) {
        throw new Error("order_by must look like 'fields/Title desc' or 'lastModifiedDateTime desc'.");
      }
      req = req.orderby(opts.orderBy);
    }
    response = await req.get();
  }
  const page = pageOf<GraphListItem>(response);
  return { ...page, items: page.items.map(summarizeListItem) };
}

export async function getListItem(client: Client, site: string, list: string, itemId: string): Promise<ListItemSummary> {
  const item = await client
    .api(`${listBase(await siteBase(client, site), list)}/items/${encodeId(itemId)}`)
    .select(ITEM_SELECT)
    .expand("fields")
    .get();
  return summarizeListItem(item as GraphListItem);
}

export type FieldValue = string | number | boolean | null | string[] | number[];

function validateFieldValues(fields: Record<string, FieldValue>): Record<string, FieldValue> {
  const keys = Object.keys(fields);
  if (keys.length === 0) throw new Error("fields must contain at least one column.");
  // "Name@odata.type" annotations are allowed for multi-value columns, e.g. "Tags@odata.type": "Collection(Edm.String)".
  for (const k of keys) {
    const base = k.endsWith("@odata.type") ? k.slice(0, -"@odata.type".length) : k;
    validateFieldNames([base]);
  }
  return fields;
}

export async function createListItem(
  client: Client,
  site: string,
  list: string,
  fields: Record<string, FieldValue>
): Promise<ListItemSummary> {
  const created = await client.api(`${listBase(await siteBase(client, site), list)}/items`).post({ fields: validateFieldValues(fields) });
  return summarizeListItem(created as GraphListItem);
}

export async function updateListItem(
  client: Client,
  site: string,
  list: string,
  itemId: string,
  fields: Record<string, FieldValue>
): Promise<{ id: string; fields?: Record<string, unknown> }> {
  const updated = await client
    .api(`${listBase(await siteBase(client, site), list)}/items/${encodeId(itemId)}/fields`)
    .patch(validateFieldValues(fields));
  return dropUndefined({ id: itemId, fields: cleanFields(updated as Record<string, unknown>) });
}

export async function deleteListItem(client: Client, site: string, list: string, itemId: string): Promise<{ deleted: true; id: string; note: string }> {
  await client.api(`${listBase(await siteBase(client, site), list)}/items/${encodeId(itemId)}`).delete();
  return { deleted: true, id: itemId, note: "Moved to the site recycle bin; a site owner can restore it." };
}
