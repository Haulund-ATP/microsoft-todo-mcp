import { beforeEach, describe, expect, it, vi } from "vitest";
import { deflateRawSync } from "node:zlib";
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

// Tool handlers run against the fake client; the real helper would resolve a connection and acquire a token.
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

const files = await import("../../src/graph/m365/files.js");
const { registerFilesTools } = await import("../../src/tools/m365/filesTools.js");

interface Call {
  path: string;
  method: string;
  query: Record<string, unknown>;
  headers: Record<string, string>;
  select?: string;
  top?: number;
  responseType?: string;
  body?: unknown;
}

/** Fake Graph client: every request is recorded and answered by `respond`. */
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
        top: (n: number) => ((call.top = n), b),
        query: (q: Record<string, unknown>) => (Object.assign(call.query, q), b),
        header: (k: string, v: string) => ((call.headers[k] = v), b),
        responseType: (t: string) => ((call.responseType = t), b),
        get: () => send("GET"),
        post: (body: unknown) => send("POST", body),
        put: (body: unknown) => send("PUT", body),
        patch: (body: unknown) => send("PATCH", body),
        delete: () => send("DELETE"),
      };
      return b;
    },
  };
  return { client: client as unknown as Client, calls };
}

function fakeServer() {
  const tools = new Map<string, { config: { inputSchema: Record<string, { safeParse: (v: unknown) => { success: boolean } }>; description: string }; handler: (args: Record<string, unknown>) => Promise<{ content: Array<{ text: string }>; isError?: boolean }> }>();
  return {
    server: { registerTool: (name: string, config: never, handler: never) => tools.set(name, { config, handler }) } as never,
    tools,
  };
}

const ctx = (allowSend = false) => ({ profileAlias: "test", correlationId: "corr-1", boundConnectionId: "conn-1", allowSend });

beforeEach(() => {
  state.audits.length = 0;
  state.modules.length = 0;
});

describe("item addressing", () => {
  it("addresses by id, by path and on another drive", () => {
    expect(files.itemUrl({ item_id: "01ABC!12" })).toBe("/me/drive/items/01ABC!12");
    expect(files.itemUrl({ item_id: "01ABC", drive_id: "b!xYz-_1" }, "/content")).toBe("/drives/b!xYz-_1/items/01ABC/content");
    expect(files.itemUrl({ path: "/Docs/Plan Q3.docx" })).toBe("/me/drive/root:/Docs/Plan%20Q3.docx");
    expect(files.itemUrl({ path: "Docs/Plan.docx", drive_id: "b!lib" }, "/content")).toBe("/drives/b!lib/root:/Docs/Plan.docx:/content");
  });

  it("percent-encodes every path segment so #, %, ? and & cannot break the URL", () => {
    expect(files.itemUrl({ path: "A#1/50% off?/R&D.txt" })).toBe("/me/drive/root:/A%231/50%25%20off%3F/R%26D.txt");
    expect(files.itemUrl({ path: "Mappe\\Æble.txt" })).toBe("/me/drive/root:/Mappe/%C3%86ble.txt");
  });

  it("handles the root, rejects ambiguous or traversal references and odd ids", () => {
    expect(files.itemUrl({ path: "/" }, "/children")).toBe("/me/drive/root/children");
    expect(files.itemUrl({}, "/children", true)).toBe("/me/drive/root/children");
    expect(() => files.itemUrl({})).toThrow(/item_id or path/);
    expect(() => files.itemUrl({ item_id: "x", path: "y" })).toThrow(/not both/);
    expect(() => files.itemUrl({ path: "a/../b" })).toThrow(/'\.\.'/);
    expect(() => files.itemUrl({ item_id: "a/b" })).toThrow(/Invalid id/);
    expect(() => files.itemUrl({ item_id: "x", drive_id: "a?b" })).toThrow(/Invalid id/);
  });
});

describe("listing and metadata", () => {
  it("lists a path's children with $select and a clamped page size, and summarises items", async () => {
    const { client, calls } = fakeGraph(() => ({
      value: [
        {
          id: "1",
          name: "Notes.loop",
          size: 10,
          webUrl: "https://example.test/notes",
          parentReference: { driveId: "d1", id: "p1", path: "/drive/root:/Team%20Docs" },
          file: { mimeType: "application/octet-stream" },
          lastModifiedBy: { user: { displayName: "A. Person" } },
        },
        { id: "2", name: "Sub", folder: { childCount: 3 }, parentReference: { driveId: "d1" } },
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=x",
    }));
    const page = await files.listChildren(client, { path: "Team Docs" }, { top: 500 });
    expect(calls[0]).toMatchObject({ path: "/me/drive/root:/Team%20Docs:/children", top: 100, select: files.ITEM_SELECT });
    expect(page.items[0]).toEqual({
      id: "1",
      name: "Notes.loop",
      drive_id: "d1",
      parent_id: "p1",
      parent_path: "/Team Docs",
      type: "file",
      file_type: "loop",
      mime_type: "application/octet-stream",
      size: 10,
      web_url: "https://example.test/notes",
      last_modified_by: "A. Person",
      is_loop: true,
    });
    expect(page.items[1]).toMatchObject({ type: "folder", child_count: 3 });
    expect(page.next_page_token).toContain("skiptoken");
  });

  it("follows only Graph page tokens", async () => {
    const { client, calls } = fakeGraph(() => ({ value: [] }));
    await files.listChildren(client, {}, { pageToken: "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=y" });
    expect(calls[0]!.path).toContain("skiptoken=y");
    await expect(files.listChildren(client, {}, { pageToken: "https://evil.example/steal" })).rejects.toThrow(/next_page_token/);
  });

  it("maps shared items through their remoteItem facet", () => {
    const item = files.summarizeItem({ id: "local", name: "x", remoteItem: { id: "remote", name: "Budget.xlsx", parentReference: { driveId: "other" } } });
    expect(item).toMatchObject({ id: "remote", name: "Budget.xlsx", drive_id: "other", file_type: "xlsx" });
  });
});

describe("files_search", () => {
  it("searches one drive with an escaped query and filters Loop files client-side", async () => {
    const { client, calls } = fakeGraph(() => ({
      value: [{ id: "a", name: "Plan.loop" }, { id: "b", name: "Plan.docx" }, { id: "c", name: "Old.fluid" }],
    }));
    const result = await files.searchFiles(client, { query: "it's #1", driveId: "b!lib", fileType: "loop" });
    expect(calls[0]!.path).toBe("/drives/b!lib/root/search(q='it''s%20%231')");
    expect(result.items.map((i) => i.id)).toEqual(["a", "c"]);
    expect(result.note).toMatch(/each page/);
  });

  it("uses Microsoft Search for the organization scope with a KQL filetype filter and offset paging", async () => {
    const { client, calls } = fakeGraph(() => ({
      value: [
        {
          hitsContainers: [
            {
              total: 40,
              moreResultsAvailable: true,
              hits: [{ summary: "the <c0>plan</c0> for Q3", resource: { id: "x", name: "Plan.loop", parentReference: { driveId: "d9", siteId: "s1" } } }],
            },
          ],
        },
      ],
    }));
    const result = await files.searchFiles(client, { query: "plan", scope: "organization", fileType: "loop", top: 10, offset: 20 });
    expect(calls[0]).toMatchObject({ path: "/search/query", method: "POST" });
    expect(calls[0]!.body).toEqual({
      requests: [{ entityTypes: ["driveItem"], query: { queryString: "plan (filetype:loop OR filetype:fluid)" }, from: 20, size: 10 }],
    });
    expect(result.items[0]).toMatchObject({ id: "x", drive_id: "d9", site_id: "s1", is_loop: true, summary: "the plan for Q3" });
    expect(result.next_offset).toBe(21);
    expect(result.total).toBe(40);
  });

  it("rejects KQL injection through file_type", async () => {
    const { client } = fakeGraph(() => ({}));
    await expect(files.searchFiles(client, { query: "x", scope: "organization", fileType: "pdf OR author:x" })).rejects.toThrow(/plain extension/);
  });
});

describe("files_read_content", () => {
  const meta = (name: string, size: number, extra: Record<string, unknown> = {}) => ({
    id: "item-1",
    name,
    size,
    webUrl: "https://example.test/f",
    parentReference: { driveId: "drive-9" },
    file: { mimeType: "application/octet-stream" },
    ...extra,
  });

  it("converts .loop files with ?format=html, addresses content by id on the item's drive, and returns text", async () => {
    const { client, calls } = fakeGraph((call) =>
      call.path.endsWith("/content") ? "<html><body><h1>Agenda</h1><ul><li>Punkt&nbsp;1</li><li>Punkt 2</li></ul></body></html>" : meta("Møde.loop", 4000)
    );
    const result = await files.readFileContent(client, { path: "Møder/Møde.loop" });
    expect(calls[0]!.path).toBe("/me/drive/root:/M%C3%B8der/M%C3%B8de.loop");
    expect(calls[1]).toMatchObject({ path: "/drives/drive-9/items/item-1/content", query: { format: "html" }, responseType: "text" });
    expect(result.content_kind).toBe("loop");
    expect(result.text).toBe("Agenda\nPunkt 1\nPunkt 2");
    expect(result.note).toMatch(/read-only/);
  });

  it("does not ask Graph to convert Loop files over its 2 MB limit", async () => {
    const { client, calls } = fakeGraph(() => meta("Big.loop", 3 * 1024 * 1024));
    const result = await files.readFileContent(client, { item_id: "item-1" });
    expect(calls).toHaveLength(1);
    expect(result.text).toBeUndefined();
    expect(result.note).toMatch(/2 MB/);
  });

  it("truncates text at max_chars", async () => {
    const { client } = fakeGraph((call) => (call.path.endsWith("/content") ? Buffer.from("x".repeat(5000)) : meta("log.txt", 5000)));
    const result = await files.readFileContent(client, { item_id: "item-1" }, { maxChars: 1000 });
    expect(result.text).toHaveLength(1000);
    expect(result.truncated).toBe(true);
  });

  it("reads only the first 2 MB of a large text file with a Range header", async () => {
    const { client, calls } = fakeGraph((call) =>
      call.path.endsWith("/content") ? new Uint8Array(files.TEXT_DOWNLOAD_CAP + 10).fill(0x61).buffer : meta("huge.csv", 50 * 1024 * 1024)
    );
    const result = await files.readFileContent(client, { item_id: "item-1" });
    expect(calls[1]!.headers.Range).toBe(`bytes=0-${files.TEXT_DOWNLOAD_CAP - 1}`);
    expect(calls[1]!.responseType).toBe("arraybuffer");
    expect(result).toMatchObject({ partial_download: true, truncated: true });
    expect(result.text).toHaveLength(files.DEFAULT_TEXT_CHARS);
  });

  it("never returns binary data as text", async () => {
    const { client } = fakeGraph((call) => (call.path.endsWith("/content") ? Buffer.from([0x50, 0x00, 0x01, 0x02]) : meta("data.json", 4)));
    const result = await files.readFileContent(client, { item_id: "item-1" });
    expect(result.text).toBeUndefined();
    expect(result.content_kind).toBe("unsupported");
  });

  it("returns metadata and web_url for PDFs and images without downloading them", async () => {
    for (const name of ["report.pdf", "photo.jpg", "old.doc"]) {
      const { client, calls } = fakeGraph(() => meta(name, 1000));
      const result = await files.readFileContent(client, { item_id: "item-1" });
      expect(calls).toHaveLength(1);
      expect(result.text).toBeUndefined();
      expect(result.item.web_url).toBe("https://example.test/f");
    }
  });

  it("extracts text from a .docx and reports folders", async () => {
    const xml = '<w:document><w:body><w:p><w:r><w:t>Hej fra Word</w:t></w:r></w:p></w:body></w:document>';
    const docx = zipOf("word/document.xml", xml);
    const { client, calls } = fakeGraph((call) => (call.path.endsWith("/content") ? docx : meta("Brev.docx", docx.length)));
    const result = await files.readFileContent(client, { item_id: "item-1" });
    expect(calls[1]!.responseType).toBe("arraybuffer");
    expect(result).toMatchObject({ content_kind: "office", text: "Hej fra Word", truncated: false });

    const folder = fakeGraph(() => ({ id: "f", name: "Mappe", folder: { childCount: 1 } }));
    expect((await files.readFileContent(folder.client, { item_id: "f" })).content_kind).toBe("folder");
  });

  it("degrades to metadata when an Office file cannot be parsed or is too large", async () => {
    const broken = fakeGraph((call) => (call.path.endsWith("/content") ? Buffer.from("not a zip") : meta("x.xlsx", 9)));
    expect((await files.readFileContent(broken.client, { item_id: "item-1" })).note).toMatch(/could not be extracted/);
    const big = fakeGraph(() => meta("x.pptx", files.OFFICE_DOWNLOAD_CAP + 1));
    const result = await files.readFileContent(big.client, { item_id: "item-1" });
    expect(big.calls).toHaveLength(1);
    expect(result.note).toMatch(/too large/);
  });
});

function zipOf(name: string, text: string): Buffer {
  const raw = Buffer.from(text);
  const comp = deflateRawSync(raw);
  const n = Buffer.from(name);
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(8, 8);
  local.writeUInt32LE(comp.length, 18);
  local.writeUInt32LE(raw.length, 22);
  local.writeUInt16LE(n.length, 26);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(8, 10);
  central.writeUInt32LE(comp.length, 20);
  central.writeUInt32LE(raw.length, 24);
  central.writeUInt16LE(n.length, 28);
  const cdOffset = 30 + n.length + comp.length;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(1, 8);
  eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(46 + n.length, 12);
  eocd.writeUInt32LE(cdOffset, 16);
  return Buffer.concat([local, n, comp, central, n, eocd]);
}

describe("writes", () => {
  it("uploads text to an encoded path with a conflict behaviour and content type", async () => {
    const { client, calls } = fakeGraph(() => ({ id: "new", name: "Q3 #1.md", size: 5 }));
    const item = await files.uploadFile(client, { path: "Rapporter/Q3 #1.md", content: "hello", encoding: "text", conflictBehavior: "fail" });
    expect(calls[0]).toMatchObject({
      method: "PUT",
      path: "/me/drive/root:/Rapporter/Q3%20%231.md:/content",
      query: { "@microsoft.graph.conflictBehavior": "fail" },
      headers: { "Content-Type": "text/plain; charset=utf-8" },
    });
    expect(Buffer.isBuffer(calls[0]!.body) && (calls[0]!.body as Buffer).toString()).toBe("hello");
    expect(item).toMatchObject({ id: "new" });
  });

  it("uploads base64 into a folder by id on another drive", async () => {
    const { client, calls } = fakeGraph(() => ({ id: "n" }));
    await files.uploadFile(client, {
      driveId: "b!lib",
      parentId: "folder-1",
      fileName: "logo.png",
      content: Buffer.from([1, 2, 3]).toString("base64"),
      encoding: "base64",
      conflictBehavior: "rename",
    });
    expect(calls[0]!.path).toBe("/drives/b!lib/items/folder-1:/logo.png:/content");
    expect(calls[0]!.headers["Content-Type"]).toBe("application/octet-stream");
    expect([...(calls[0]!.body as Buffer)]).toEqual([1, 2, 3]);
  });

  it("refuses uploads over 4 MB, invalid base64 and invalid names before calling Graph", async () => {
    const { client, calls } = fakeGraph(() => ({}));
    await expect(files.uploadFile(client, { path: "big.txt", content: "x".repeat(files.UPLOAD_CAP + 1), encoding: "text", conflictBehavior: "fail" })).rejects.toThrow(/4 MB/);
    await expect(files.uploadFile(client, { path: "a.bin", content: "@@@", encoding: "base64", conflictBehavior: "fail" })).rejects.toThrow(/base64/);
    await expect(files.uploadFile(client, { fileName: "a:b.txt", content: "x", encoding: "text", conflictBehavior: "fail" })).rejects.toThrow(/Invalid file/);
    expect(calls).toHaveLength(0);
  });

  it("creates folders with a conflict behaviour", async () => {
    const { client, calls } = fakeGraph(() => ({ id: "f", name: "Ny", folder: {} }));
    await files.createFolder(client, { path: "Projekter" }, "Ny", "rename");
    expect(calls[0]).toMatchObject({
      method: "POST",
      path: "/me/drive/root:/Projekter:/children",
      body: { name: "Ny", folder: {}, "@microsoft.graph.conflictBehavior": "rename" },
    });
  });

  it("moves and renames with parentReference resolved from a destination path", async () => {
    const { client, calls } = fakeGraph((call) => (call.method === "GET" ? { id: "dest", folder: {}, parentReference: { driveId: "d1" } } : { id: "i", name: "B.txt" }));
    await files.moveOrRenameItem(client, { item_id: "i" }, { newName: "B.txt", destination: { path: "Arkiv/2026" } });
    expect(calls[0]!.path).toBe("/me/drive/root:/Arkiv/2026");
    expect(calls[1]).toMatchObject({ method: "PATCH", path: "/me/drive/items/i", body: { name: "B.txt", parentReference: { id: "dest" } } });
  });

  it("starts an async copy with driveId+id and reports it as accepted; surfaces rejections", async () => {
    const { client, calls } = fakeGraph((call) =>
      call.method === "GET" ? { id: "dest", folder: {}, parentReference: { driveId: "d2" } } : { status: 202, headers: { get: () => "https://example.test/monitor" } }
    );
    const result = await files.copyItem(client, { item_id: "src" }, { destination: { item_id: "dest", drive_id: "d2" }, conflictBehavior: "rename" });
    expect(calls[1]).toMatchObject({
      method: "POST",
      path: "/me/drive/items/src/copy",
      responseType: "raw",
      query: { "@microsoft.graph.conflictBehavior": "rename" },
      body: { parentReference: { driveId: "d2", id: "dest" } },
    });
    expect(result.status).toBe("accepted");
    expect(JSON.stringify(result)).not.toContain("monitor");

    const rejected = fakeGraph(() => ({ status: 409, ok: false, json: async () => ({ error: { code: "nameAlreadyExists", message: "Name already exists" } }) }));
    await expect(files.copyItem(rejected.client, { item_id: "src" }, { newName: "Kopi.txt" })).rejects.toMatchObject({ statusCode: 409, message: "Name already exists" });
  });

  it("deletes by id after looking the item up, and refuses the root", async () => {
    const { client, calls } = fakeGraph((call) => (call.method === "GET" ? { id: "abc", name: "Gammel.txt", parentReference: { driveId: "d1" } } : undefined));
    const result = await files.deleteItem(client, { path: "Gammel.txt" });
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual(["GET /me/drive/root:/Gammel.txt", "DELETE /drives/d1/items/abc"]);
    expect(result).toMatchObject({ deleted: true, name: "Gammel.txt" });
    await expect(files.deleteItem(client, { path: "/" })).rejects.toThrow(/root/);
  });
});

describe("sharing links", () => {
  it("creates organization links and refuses anonymous links unless allowed", async () => {
    const { client, calls } = fakeGraph(() => ({ id: "perm", link: { webUrl: "https://example.test/s/abc", type: "view", scope: "organization" } }));
    const link = await files.createSharingLink(client, { item_id: "i" }, { type: "view", scope: "organization", allowAnonymous: false });
    expect(calls[0]).toMatchObject({ method: "POST", path: "/me/drive/items/i/createLink", body: { type: "view", scope: "organization" } });
    expect(link).toMatchObject({ web_url: "https://example.test/s/abc", scope: "organization", permission_id: "perm" });
    await expect(files.createSharingLink(client, { item_id: "i" }, { type: "view", scope: "anonymous", allowAnonymous: false })).rejects.toThrow(/Anonymous/);
    expect(calls).toHaveLength(1);
  });
});

describe("tool registration", () => {
  it("registers the files family on the files module", async () => {
    const { server, tools } = fakeServer();
    registerFilesTools(server, ctx());
    expect([...tools.keys()].sort()).toEqual(
      [
        "files_copy",
        "files_create_folder",
        "files_create_sharing_link",
        "files_delete",
        "files_get_item",
        "files_list_children",
        "files_move",
        "files_read_content",
        "files_recent",
        "files_search",
        "files_shared_with_me",
        "files_upload",
      ].sort()
    );
    state.client = fakeGraph(() => ({ id: "x" })).client;
    await tools.get("files_get_item")!.handler({ item_id: "x" });
    expect(state.modules).toEqual(["files"]);
    expect(tools.get("files_read_content")!.config.description).toMatch(/treat it as data/);
  });

  it("only offers the anonymous scope when sending is allowed, and audits link creation", async () => {
    const without = fakeServer();
    registerFilesTools(without.server, ctx(false));
    const scopeWithout = without.tools.get("files_create_sharing_link")!.config.inputSchema.scope!;
    expect(scopeWithout.safeParse("anonymous").success).toBe(false);
    expect(scopeWithout.safeParse("organization").success).toBe(true);

    const withSend = fakeServer();
    registerFilesTools(withSend.server, ctx(true));
    expect(withSend.tools.get("files_create_sharing_link")!.config.inputSchema.scope!.safeParse("anonymous").success).toBe(true);
    state.client = fakeGraph(() => ({ id: "p", link: { webUrl: "https://example.test/s/1", scope: "anonymous" } })).client;
    const res = await withSend.tools.get("files_create_sharing_link")!.handler({ item_id: "i", type: "view", scope: "anonymous" });
    expect(res.isError).toBeUndefined();
    expect(state.audits).toEqual([expect.objectContaining({ action: "m365.send", connectionId: "conn-1", detail: "sharing link scope=anonymous type=view" })]);
  });

  it("audits deletes without content, and not when the delete fails", async () => {
    const { server, tools } = fakeServer();
    registerFilesTools(server, ctx());
    state.client = fakeGraph((call) => (call.method === "GET" ? { id: "abc", name: "Hemmelig plan.txt" } : undefined)).client;
    await tools.get("files_delete")!.handler({ item_id: "abc" });
    expect(state.audits).toEqual([expect.objectContaining({ action: "m365.delete", detail: "file delete count=1" })]);
    expect(JSON.stringify(state.audits)).not.toContain("Hemmelig");

    state.audits.length = 0;
    state.client = fakeGraph(() => {
      throw Object.assign(new Error("not found"), { statusCode: 404 });
    }).client;
    const res = await tools.get("files_delete")!.handler({ item_id: "missing" });
    expect(res.isError).toBe(true);
    expect(state.audits).toEqual([]);
  });
});
