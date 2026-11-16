import { ResponseType, type Client } from "@microsoft/microsoft-graph-client";
import { htmlToText, nextPageRequest, pageOf, truncateText, type Page } from "../../tools/m365/common.js";
import { extractOfficeText, officeKindForExtension, OfficeTextError } from "./officeText.js";

/**
 * OneDrive / SharePoint document library files through Graph v1.0,
 * delegated (`Files.ReadWrite.All`, personal and work accounts).
 * Endpoint notes (verified against Microsoft Learn):
 *
 * - Items are addressed by id (`/drives/{d}/items/{id}`) or by path
 *   (`/drives/{d}/root:/{path}` and `/drives/{d}/root:/{path}:/{sub}`); the
 *   drive defaults to `/me/drive`. Path segments are percent-encoded one by
 *   one so `#`, `%`, `?` and spaces in names cannot break the URL.
 * - `GET {drive}/root/search(q='…')` searches one drive (all account types).
 *   `POST /search/query` with entityTypes driveItem searches everything the
 *   user can read across OneDrive and SharePoint (work/school accounts only;
 *   Files.ReadWrite.All is sufficient). KQL `filetype:` narrows the latter.
 * - `GET {item}/content?format=html` converts Loop (.loop, .fluid) and
 *   whiteboard files to HTML; Graph limits format conversion to source files
 *   of at most 2 MB. Loop content cannot be written through Graph.
 * - `GET /me/drive/recent` and `/me/drive/sharedWithMe` are deprecated by
 *   Microsoft (degraded until November 2026, then empty).
 * - Simple upload `PUT …/content` is limited to 4 MB; larger files need an
 *   upload session (not offered here).
 * - `POST {item}/copy` is asynchronous: 202 Accepted with a monitor URL on a
 *   non-Graph host, which this module never calls.
 * - `DELETE {item}` moves the item to the recycle bin.
 */

export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;
/** Characters of file text returned by default. */
export const DEFAULT_TEXT_CHARS = 50_000;
export const MAX_TEXT_CHARS = 100_000;
/** Bytes of a text file downloaded at most (larger files are read from the start with a Range request). */
export const TEXT_DOWNLOAD_CAP = 2 * 1024 * 1024;
/** Office files above this size are not downloaded for extraction. */
export const OFFICE_DOWNLOAD_CAP = 20 * 1024 * 1024;
/** Graph's limit for format conversion (?format=html). */
export const FORMAT_CONVERSION_CAP = 2 * 1024 * 1024;
/** Simple upload limit. */
export const UPLOAD_CAP = 4 * 1024 * 1024;

export const ITEM_SELECT =
  "id,name,size,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,parentReference,file,folder,package,remoteItem";

export type ConflictBehavior = "fail" | "replace" | "rename";

export interface ItemRef {
  item_id?: string;
  path?: string;
  drive_id?: string;
}

// ---------------------------------------------------------------------------
// Addressing

/** Percent-encodes an id for use as one URL segment; commas stay literal (SharePoint composite ids). */
export function encodeId(id: string): string {
  const trimmed = id.trim();
  if (trimmed === "" || /[\s/\\?#]/.test(trimmed)) throw new Error(`Invalid id: ${JSON.stringify(id)}`);
  return encodeURIComponent(trimmed).replace(/%2C/gi, ",");
}

export function driveBase(driveId?: string): string {
  return driveId ? `/drives/${encodeId(driveId)}` : "/me/drive";
}

/** Normalises a drive path ("/A/B c.txt", "A\\B") into encoded segments joined by "/"; "" means the root. */
export function encodeDrivePath(path: string): string {
  const segments = path
    .replace(/\\/g, "/")
    .split("/")
    .filter((s) => s !== "");
  for (const s of segments) {
    if (s === "." || s === "..") throw new Error("Paths may not contain '.' or '..' segments.");
  }
  return segments.map((s) => encodeURIComponent(s)).join("/");
}

/**
 * Builds the Graph URL for an item (optionally with a sub-resource such as
 * "/children" or "/content"). Exactly one of item_id / path is used; with
 * neither, `allowRoot` decides between the drive root and an error.
 */
export function itemUrl(ref: ItemRef, suffix = "", allowRoot = false): string {
  const base = driveBase(ref.drive_id);
  const hasId = ref.item_id !== undefined && ref.item_id !== "";
  const hasPath = ref.path !== undefined && ref.path !== "";
  if (hasId && hasPath) throw new Error("Pass either item_id or path, not both.");
  if (hasId) return `${base}/items/${encodeId(ref.item_id!)}${suffix}`;
  if (hasPath) {
    const encoded = encodeDrivePath(ref.path!);
    if (encoded === "") return `${base}/root${suffix}`;
    return suffix ? `${base}/root:/${encoded}:${suffix}` : `${base}/root:/${encoded}`;
  }
  if (allowRoot) return `${base}/root${suffix}`;
  throw new Error("Pass item_id or path to identify the file or folder.");
}

const INVALID_NAME = /[\\/:*?"<>|]/;

export function validateFileName(name: string): string {
  const trimmed = name.trim();
  if (trimmed === "" || trimmed === "." || trimmed === ".." || INVALID_NAME.test(trimmed)) {
    throw new Error(`Invalid file or folder name: ${JSON.stringify(name)} (must not contain \\ / : * ? " < > |).`);
  }
  return trimmed;
}

export function extensionOf(name: string | undefined | null): string {
  if (!name) return "";
  const dot = name.lastIndexOf(".");
  return dot > 0 ? name.slice(dot + 1).toLowerCase() : "";
}

function clampTop(top: number | undefined): number {
  return Math.min(Math.max(top ?? DEFAULT_PAGE_SIZE, 1), MAX_PAGE_SIZE);
}

// ---------------------------------------------------------------------------
// Item shaping

interface GraphIdentitySet {
  user?: { displayName?: string | null; email?: string | null } | null;
  application?: { displayName?: string | null } | null;
}

export interface GraphDriveItem {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: GraphIdentitySet;
  lastModifiedBy?: GraphIdentitySet;
  parentReference?: { driveId?: string; driveType?: string; id?: string; path?: string; siteId?: string } | null;
  file?: { mimeType?: string } | null;
  folder?: { childCount?: number } | null;
  package?: { type?: string } | null;
  remoteItem?: GraphDriveItem | null;
}

export interface FileItem {
  id?: string;
  name?: string;
  drive_id?: string;
  parent_id?: string;
  parent_path?: string;
  site_id?: string;
  type: "file" | "folder" | "package";
  file_type?: string;
  mime_type?: string;
  size?: number;
  child_count?: number;
  web_url?: string;
  created?: string;
  last_modified?: string;
  last_modified_by?: string;
  created_by?: string;
  is_loop?: boolean;
}

const LOOP_EXTENSIONS = new Set(["loop", "fluid"]);
const WHITEBOARD_EXTENSIONS = new Set(["wbtx", "whiteboard"]);

function who(set: GraphIdentitySet | undefined): string | undefined {
  return set?.user?.displayName ?? set?.application?.displayName ?? undefined;
}

/** Reduces a driveItem (or a shared/recent item with a remoteItem facet) to a compact summary. */
export function summarizeItem(raw: GraphDriveItem): FileItem {
  const src: GraphDriveItem = raw.remoteItem ? { ...raw, ...raw.remoteItem } : raw;
  const ext = extensionOf(src.name);
  const out: FileItem = {
    id: src.id,
    name: src.name,
    drive_id: src.parentReference?.driveId ?? undefined,
    parent_id: src.parentReference?.id ?? undefined,
    parent_path: src.parentReference?.path ? decodeParentPath(src.parentReference.path) : undefined,
    site_id: src.parentReference?.siteId ?? undefined,
    type: src.folder ? "folder" : src.package ? "package" : "file",
    file_type: src.folder ? undefined : ext || undefined,
    mime_type: src.file?.mimeType ?? undefined,
    size: src.size,
    child_count: src.folder?.childCount,
    web_url: src.webUrl,
    created: src.createdDateTime,
    last_modified: src.lastModifiedDateTime,
    last_modified_by: who(src.lastModifiedBy),
    created_by: who(src.createdBy),
  };
  if (LOOP_EXTENSIONS.has(ext)) out.is_loop = true;
  for (const k of Object.keys(out) as Array<keyof FileItem>) if (out[k] === undefined) delete out[k];
  return out;
}

/** "/drive/root:/Docs/A%20B" -> "/Docs/A B" (the drive prefix is noise for the model). */
function decodeParentPath(path: string): string {
  const i = path.indexOf("root:");
  const rest = i >= 0 ? path.slice(i + 5) : path;
  try {
    return decodeURIComponent(rest) || "/";
  } catch {
    return rest || "/";
  }
}

function itemPage(response: { value?: GraphDriveItem[]; "@odata.nextLink"?: string }): Page<FileItem> {
  const page = pageOf(response);
  return { ...page, items: page.items.map(summarizeItem) };
}

/** Optional filter on file extension; "loop" also matches .fluid. */
export function matchesFileType(item: FileItem, fileType: string | undefined): boolean {
  if (!fileType) return true;
  const wanted = fileType.replace(/^\./, "").toLowerCase();
  if (wanted === "folder") return item.type === "folder";
  if (wanted === "loop") return item.file_type !== undefined && LOOP_EXTENSIONS.has(item.file_type);
  return item.file_type === wanted;
}

// ---------------------------------------------------------------------------
// Read operations

export async function listChildren(
  client: Client,
  ref: ItemRef,
  opts: { top?: number; pageToken?: string } = {}
): Promise<Page<FileItem>> {
  if (opts.pageToken) return itemPage(await nextPageRequest(client, opts.pageToken).get());
  const response = await client.api(itemUrl(ref, "/children", true)).select(ITEM_SELECT).top(clampTop(opts.top)).get();
  return itemPage(response);
}

export async function getItem(client: Client, ref: ItemRef): Promise<FileItem> {
  return summarizeItem(await client.api(itemUrl(ref, "", true)).select(ITEM_SELECT).get());
}

export function escapeSearchQuery(q: string): string {
  return encodeURIComponent(q.replace(/'/g, "''"));
}

export interface SearchOptions {
  query: string;
  scope?: "drive" | "organization";
  driveId?: string;
  fileType?: string;
  top?: number;
  pageToken?: string;
  offset?: number;
}

export interface SearchResult {
  items: Array<FileItem & { summary?: string }>;
  next_page_token?: string;
  next_offset?: number;
  total?: number;
  note?: string;
}

/**
 * Searches one drive (default: the user's OneDrive) or, with
 * scope "organization", everything the user can read in OneDrive and
 * SharePoint through the Microsoft Search API.
 */
export async function searchFiles(client: Client, opts: SearchOptions): Promise<SearchResult> {
  const top = clampTop(opts.top);
  if (opts.scope === "organization") {
    const fileType = opts.fileType?.replace(/^\./, "").toLowerCase();
    let queryString = opts.query.trim() || "*";
    if (fileType === "loop") queryString += " (filetype:loop OR filetype:fluid)";
    else if (fileType === "folder") queryString += " isDocument:false";
    else if (fileType) {
      if (!/^[a-z0-9]{1,10}$/.test(fileType)) throw new Error("file_type must be a plain extension such as docx, pdf or loop.");
      queryString += ` filetype:${fileType}`;
    }
    const from = Math.max(0, opts.offset ?? 0);
    const response = (await client.api("/search/query").post({
      requests: [{ entityTypes: ["driveItem"], query: { queryString }, from, size: top }],
    })) as {
      value?: Array<{
        hitsContainers?: Array<{ total?: number; moreResultsAvailable?: boolean; hits?: Array<{ summary?: string; resource?: GraphDriveItem }> }>;
      }>;
    };
    const container = response.value?.[0]?.hitsContainers?.[0];
    const hits = container?.hits ?? [];
    const items = hits
      .filter((h) => h.resource)
      .map((h) => {
        const item: FileItem & { summary?: string } = summarizeItem(h.resource!);
        if (h.summary) item.summary = htmlToText(h.summary.replace(/<\/?c\d+>/g, "")).slice(0, 500);
        return item;
      });
    const result: SearchResult = { items, total: container?.total };
    if (container?.moreResultsAvailable) result.next_offset = from + hits.length;
    return result;
  }

  let page: Page<FileItem>;
  if (opts.pageToken) {
    page = itemPage(await nextPageRequest(client, opts.pageToken).get());
  } else {
    const path = `${driveBase(opts.driveId)}/root/search(q='${escapeSearchQuery(opts.query)}')`;
    page = itemPage(await client.api(path).select(ITEM_SELECT).top(top).get());
  }
  const items = page.items.filter((i) => matchesFileType(i, opts.fileType));
  const result: SearchResult = { items };
  if (page.next_page_token) result.next_page_token = page.next_page_token;
  if (opts.fileType && items.length < page.items.length) {
    result.note = "file_type is applied to each page of drive search results; continue with next_page_token for more matches.";
  }
  return result;
}

const DEPRECATED_NOTE =
  "Microsoft has deprecated this Graph endpoint (degraded until November 2026, then it returns nothing). " +
  "If it is empty, use files_search (scope organization on work accounts) instead.";

export async function recentFiles(client: Client, opts: { top?: number } = {}): Promise<{ items: FileItem[]; note?: string }> {
  const response = (await client.api("/me/drive/recent").get()) as { value?: GraphDriveItem[] };
  const items = (response.value ?? []).slice(0, clampTop(opts.top)).map(summarizeItem);
  return items.length === 0 ? { items, note: DEPRECATED_NOTE } : { items };
}

export async function sharedWithMe(
  client: Client,
  opts: { top?: number; includeExternal?: boolean; pageToken?: string } = {}
): Promise<Page<FileItem> & { note?: string; more_available?: boolean }> {
  let page: Page<FileItem> & { more_available?: boolean };
  if (opts.pageToken) page = itemPage(await nextPageRequest(client, opts.pageToken).get());
  else {
    let req = client.api("/me/drive/sharedWithMe");
    if (opts.includeExternal) req = req.query({ allowexternal: "true" });
    page = itemPage(await req.get());
  }
  const top = clampTop(opts.top);
  if (page.items.length > top) {
    // The endpoint takes no $top; cut the page and drop the token (it would skip the cut items).
    page = { items: page.items.slice(0, top), more_available: true };
  }
  return page.items.length === 0 ? { ...page, note: DEPRECATED_NOTE } : page;
}

// ---------------------------------------------------------------------------
// Content

const TEXT_EXTENSIONS = new Set([
  "txt", "text", "md", "markdown", "csv", "tsv", "json", "jsonc", "xml", "yaml", "yml", "log", "ini", "cfg", "conf", "toml",
  "env", "properties", "html", "htm", "css", "scss", "js", "mjs", "cjs", "ts", "tsx", "jsx", "py", "rb", "php", "java", "kt",
  "cs", "vb", "go", "rs", "c", "h", "cpp", "hpp", "sql", "ps1", "psm1", "psd1", "sh", "bash", "bat", "cmd", "ics", "vcf",
  "vtt", "srt", "svg", "rtf", "tex", "bicep", "tf", "gitignore", "dockerfile",
]);
const HTML_EXTENSIONS = new Set(["html", "htm"]);

export type ContentKind = "text" | "loop" | "office" | "pdf" | "unsupported";

export function contentKind(name: string | undefined, mimeType: string | undefined): ContentKind {
  const ext = extensionOf(name);
  if (LOOP_EXTENSIONS.has(ext) || WHITEBOARD_EXTENSIONS.has(ext)) return "loop";
  if (officeKindForExtension(ext)) return "office";
  if (ext === "pdf" || mimeType === "application/pdf") return "pdf";
  if (TEXT_EXTENSIONS.has(ext)) return "text";
  if (mimeType && (mimeType.startsWith("text/") || /^application\/(json|xml|[a-z.+-]*\+(json|xml))$/.test(mimeType))) return "text";
  return "unsupported";
}

/** Converts whatever the SDK returned for an ARRAYBUFFER request into a Buffer. */
export function toBuffer(data: unknown): Buffer {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof ArrayBuffer) return Buffer.from(data);
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === "string") return Buffer.from(data, "utf8");
  throw new Error("Unexpected download response type.");
}

/** Decodes text bytes (UTF-8, or UTF-16 with BOM); undefined when the data looks binary. */
export function decodeText(buf: Buffer): string | undefined {
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return new TextDecoder("utf-16le").decode(buf.subarray(2));
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) return new TextDecoder("utf-16be").decode(buf.subarray(2));
  const sniff = buf.subarray(0, 8192);
  if (sniff.includes(0)) return undefined;
  return new TextDecoder("utf-8").decode(buf);
}

export interface ReadContentResult {
  item: FileItem;
  content_kind: ContentKind | "folder";
  text?: string;
  truncated?: boolean;
  /** True when only the start of a large text file was downloaded. */
  partial_download?: boolean;
  note?: string;
}

function metadataOnly(item: FileItem, kind: ReadContentResult["content_kind"], note: string): ReadContentResult {
  return { item, content_kind: kind, note };
}

export async function readFileContent(client: Client, ref: ItemRef, opts: { maxChars?: number } = {}): Promise<ReadContentResult> {
  const maxChars = Math.min(Math.max(opts.maxChars ?? DEFAULT_TEXT_CHARS, 100), MAX_TEXT_CHARS);
  const raw = (await client.api(itemUrl(ref)).select(ITEM_SELECT).get()) as GraphDriveItem;
  const item = summarizeItem(raw);
  if (item.type === "folder") return metadataOnly(item, "folder", "This is a folder; use files_list_children.");
  if (item.type === "package") return metadataOnly(item, "unsupported", "This is a package (for example a OneNote notebook); open web_url instead.");

  // After the metadata lookup, address the content by id on the item's own drive.
  const contentRef: ItemRef = { item_id: item.id, drive_id: item.drive_id ?? ref.drive_id };
  const size = item.size ?? 0;
  const kind = contentKind(item.name, item.mime_type);

  if (kind === "loop") {
    if (size > FORMAT_CONVERSION_CAP) {
      return metadataOnly(item, kind, "Graph only converts Loop/whiteboard files up to 2 MB to HTML; open web_url to read it.");
    }
    const html = await client.api(itemUrl(contentRef, "/content")).query({ format: "html" }).responseType(ResponseType.TEXT).get();
    const { text, truncated } = truncateText(htmlToText(typeof html === "string" ? html : toBuffer(html).toString("utf8")), maxChars);
    return { item, content_kind: kind, text, truncated, note: "Loop content is read-only through Graph; edit it in Loop/Teams via web_url." };
  }

  if (kind === "office") {
    if (size > OFFICE_DOWNLOAD_CAP) return metadataOnly(item, kind, "File is too large for text extraction here (over 20 MB); open web_url instead.");
    const data = toBuffer(await client.api(itemUrl(contentRef, "/content")).responseType(ResponseType.ARRAYBUFFER).get());
    try {
      const extracted = extractOfficeText(data, officeKindForExtension(item.file_type ?? "")!, { maxChars });
      return {
        item,
        content_kind: kind,
        text: extracted.text,
        truncated: extracted.truncated,
        note: "Plain text only (formatting, images, comments and headers/footers are left out; cells show stored values).",
      };
    } catch (err) {
      if (err instanceof OfficeTextError) return metadataOnly(item, kind, `Text could not be extracted: ${err.message} Open web_url instead.`);
      throw err;
    }
  }

  if (kind === "pdf") return metadataOnly(item, kind, "PDF text extraction is not supported; open web_url to read the file.");
  if (kind === "unsupported") {
    return metadataOnly(item, kind, "This file type is not readable as text here (binary, image, media or legacy Office format); open web_url instead.");
  }

  let req = client.api(itemUrl(contentRef, "/content")).responseType(ResponseType.ARRAYBUFFER);
  const partial = size > TEXT_DOWNLOAD_CAP;
  if (partial) req = req.header("Range", `bytes=0-${TEXT_DOWNLOAD_CAP - 1}`);
  let buf = toBuffer(await req.get());
  if (buf.length > TEXT_DOWNLOAD_CAP) buf = buf.subarray(0, TEXT_DOWNLOAD_CAP);
  const decoded = decodeText(buf);
  if (decoded === undefined) return metadataOnly(item, "unsupported", "The file looks binary, so no text is returned; open web_url instead.");
  const plain = HTML_EXTENSIONS.has(item.file_type ?? "") ? htmlToText(decoded) : decoded;
  const { text, truncated } = truncateText(plain, maxChars);
  const result: ReadContentResult = { item, content_kind: kind, text, truncated: truncated || partial };
  if (partial) {
    result.partial_download = true;
    result.note = "Only the first 2 MB of this large file were read.";
  }
  return result;
}

// ---------------------------------------------------------------------------
// Write operations

export interface UploadOptions {
  driveId?: string;
  /** Full destination path including the file name, e.g. "Reports/q3.txt". */
  path?: string;
  /** Alternative to path: parent folder id plus file name. */
  parentId?: string;
  fileName?: string;
  content: string;
  encoding: "text" | "base64";
  contentType?: string;
  conflictBehavior: ConflictBehavior;
}

export function decodeUploadContent(content: string, encoding: "text" | "base64"): Buffer {
  if (encoding === "text") return Buffer.from(content, "utf8");
  const compact = content.replace(/\s+/g, "");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(compact) || compact.length % 4 === 1) throw new Error("content is not valid base64.");
  return Buffer.from(compact, "base64");
}

export async function uploadFile(client: Client, opts: UploadOptions): Promise<FileItem> {
  const bytes = decodeUploadContent(opts.content, opts.encoding);
  if (bytes.length > UPLOAD_CAP) throw new Error(`File is ${bytes.length} bytes; simple upload is limited to 4 MB.`);
  const base = driveBase(opts.driveId);
  let url: string;
  if (opts.path && (opts.parentId || opts.fileName)) throw new Error("Pass either path, or parent_id with file_name — not both.");
  if (opts.path) {
    const encoded = encodeDrivePath(opts.path);
    if (encoded === "") throw new Error("path must include the file name.");
    const segments = opts.path.replace(/\\/g, "/").split("/").filter(Boolean);
    validateFileName(segments[segments.length - 1]!);
    url = `${base}/root:/${encoded}:/content`;
  } else if (opts.fileName) {
    const name = encodeURIComponent(validateFileName(opts.fileName));
    url = opts.parentId ? `${base}/items/${encodeId(opts.parentId)}:/${name}:/content` : `${base}/root:/${name}:/content`;
  } else {
    throw new Error("Pass path (including the file name), or file_name with an optional parent_id.");
  }
  const contentType = opts.contentType ?? (opts.encoding === "text" ? "text/plain; charset=utf-8" : "application/octet-stream");
  const created = await client
    .api(url)
    .query({ "@microsoft.graph.conflictBehavior": opts.conflictBehavior })
    .header("Content-Type", contentType)
    .put(bytes);
  return summarizeItem(created as GraphDriveItem);
}

export async function createFolder(
  client: Client,
  parent: ItemRef,
  name: string,
  conflictBehavior: ConflictBehavior
): Promise<FileItem> {
  const created = await client.api(itemUrl(parent, "/children", true)).post({
    name: validateFileName(name),
    folder: {},
    "@microsoft.graph.conflictBehavior": conflictBehavior,
  });
  return summarizeItem(created as GraphDriveItem);
}

/** Resolves a destination folder to the {driveId, id} pair Graph wants in parentReference. */
async function resolveFolder(client: Client, ref: ItemRef): Promise<{ driveId: string; id: string }> {
  const folder = (await client.api(itemUrl(ref, "", true)).select("id,folder,parentReference,root").get()) as GraphDriveItem & {
    root?: unknown;
  };
  if (!folder.folder && !folder.root) throw new Error("The destination is not a folder.");
  let driveId = folder.parentReference?.driveId ?? ref.drive_id;
  if (!driveId) driveId = ((await client.api(driveBase(ref.drive_id)).select("id").get()) as { id: string }).id;
  return { driveId, id: folder.id! };
}

export async function moveOrRenameItem(
  client: Client,
  ref: ItemRef,
  opts: { newName?: string; destination?: ItemRef }
): Promise<FileItem> {
  if (!opts.newName && !opts.destination) throw new Error("Pass new_name, a destination folder, or both.");
  const body: Record<string, unknown> = {};
  if (opts.newName) body.name = validateFileName(opts.newName);
  if (opts.destination) {
    const dest = await resolveFolder(client, { ...opts.destination, drive_id: opts.destination.drive_id ?? ref.drive_id });
    body.parentReference = { id: dest.id };
  }
  const updated = await client.api(itemUrl(ref)).patch(body);
  return summarizeItem(updated as GraphDriveItem);
}

interface RawResponse {
  status: number;
  ok?: boolean;
  headers?: { get(name: string): string | null };
  json?: () => Promise<unknown>;
}

export async function copyItem(
  client: Client,
  ref: ItemRef,
  opts: { destination?: ItemRef; newName?: string; conflictBehavior?: ConflictBehavior }
): Promise<{ status: "accepted"; note: string }> {
  const body: Record<string, unknown> = {};
  if (opts.newName) body.name = validateFileName(opts.newName);
  if (opts.destination) body.parentReference = await resolveFolder(client, { ...opts.destination, drive_id: opts.destination.drive_id ?? ref.drive_id });
  if (!body.name && !body.parentReference) throw new Error("Pass a destination folder, new_name, or both.");
  let req = client.api(itemUrl(ref, "/copy")).responseType(ResponseType.RAW);
  if (opts.conflictBehavior) req = req.query({ "@microsoft.graph.conflictBehavior": opts.conflictBehavior });
  const res = (await req.post(body)) as RawResponse;
  if (res.status !== 202 && !(res.ok ?? (res.status >= 200 && res.status < 300))) {
    let message = `Copy was rejected (HTTP ${res.status}).`;
    let code: string | undefined;
    try {
      const err = (await res.json?.()) as { error?: { code?: string; message?: string } } | undefined;
      if (err?.error?.message) message = err.error.message;
      code = err?.error?.code;
    } catch {
      /* keep the generic message */
    }
    throw Object.assign(new Error(message), { statusCode: res.status, code });
  }
  return {
    status: "accepted",
    note: "Graph copies in the background. Check the destination with files_list_children in a moment; a name conflict makes the copy fail silently unless conflict_behavior is rename or replace.",
  };
}

export async function deleteItem(client: Client, ref: ItemRef): Promise<{ deleted: true; name?: string; note: string }> {
  if (!ref.item_id && (!ref.path || encodeDrivePath(ref.path) === "")) throw new Error("Refusing to delete the drive root.");
  // Look the item up first so the result can name it and a wrong path fails before anything is deleted.
  const item = summarizeItem(await client.api(itemUrl(ref)).select("id,name,parentReference,folder,file,package").get());
  await client.api(itemUrl({ item_id: item.id, drive_id: item.drive_id ?? ref.drive_id })).delete();
  return { deleted: true, name: item.name, note: "Moved to the recycle bin; it can be restored from OneDrive/SharePoint." };
}

export type LinkScope = "organization" | "users" | "anonymous";

export async function createSharingLink(
  client: Client,
  ref: ItemRef,
  opts: { type: "view" | "edit"; scope: LinkScope; allowAnonymous: boolean; expirationDateTime?: string }
): Promise<{ web_url?: string; type?: string; scope?: string; expiration?: string; permission_id?: string }> {
  if (opts.scope === "anonymous" && !opts.allowAnonymous) {
    throw new Error("Anonymous links expose the file outside the organization and are only allowed on profiles with sending enabled.");
  }
  const body: Record<string, unknown> = { type: opts.type, scope: opts.scope };
  if (opts.expirationDateTime) body.expirationDateTime = opts.expirationDateTime;
  const permission = (await client.api(itemUrl(ref, "/createLink")).post(body)) as {
    id?: string;
    expirationDateTime?: string;
    link?: { webUrl?: string; type?: string; scope?: string };
  };
  return {
    web_url: permission.link?.webUrl,
    type: permission.link?.type,
    scope: permission.link?.scope,
    expiration: permission.expirationDateTime,
    permission_id: permission.id,
  };
}
