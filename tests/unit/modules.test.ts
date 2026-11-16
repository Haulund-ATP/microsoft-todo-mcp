import { describe, expect, it, vi } from "vitest";
import {
  BASE_SCOPES,
  MODULE_IDS,
  modulesCoveredByScopes,
  parseModuleList,
  scopesForModule,
  scopesForModules,
  serializeModuleList,
  supportedModules,
} from "../../src/graph/modules.js";

describe("module scopes", () => {
  it("keeps the To Do token request exactly as it was before modules existed", () => {
    expect(new Set(scopesForModule("todo"))).toEqual(
      new Set(["Tasks.ReadWrite", "User.Read", "openid", "profile", "offline_access"])
    );
  });

  it("never mixes another module's scopes into a module's token request", () => {
    const mail = scopesForModule("mail");
    expect(mail).toContain("Mail.ReadWrite");
    expect(mail).not.toContain("Tasks.ReadWrite");
    expect(mail).not.toContain("Sites.ReadWrite.All");
  });

  it("unions scopes for a multi-module sign-in without duplicates", () => {
    const scopes = scopesForModules(["todo", "mail", "teams", "teams_channels"]);
    expect(scopes.filter((s) => s === "User.Read")).toHaveLength(1);
    for (const s of BASE_SCOPES) expect(scopes).toContain(s);
    expect(scopes).toEqual(expect.arrayContaining(["Tasks.ReadWrite", "Mail.Send", "Chat.ReadWrite", "ChannelMessage.Read.All"]));
  });
});

describe("supportedModules", () => {
  it("drops work/school-only modules for personal accounts", () => {
    expect(supportedModules([...MODULE_IDS], "personal")).toEqual(["todo", "mail", "mail_rules", "calendar", "contacts", "files"]);
  });

  it("keeps everything for organizational accounts", () => {
    expect(supportedModules([...MODULE_IDS], "organizational")).toEqual([...MODULE_IDS]);
  });
});

describe("modulesCoveredByScopes", () => {
  it("recognises fully qualified, differently cased scopes", () => {
    const covered = modulesCoveredByScopes([
      "https://graph.microsoft.com/Tasks.ReadWrite",
      "https://graph.microsoft.com/mail.readwrite",
      "https://graph.microsoft.com/Mail.Send",
      "openid",
    ]);
    expect(covered).toEqual(["todo", "mail"]);
  });

  it("does not count a module whose scopes are only partly granted", () => {
    expect(modulesCoveredByScopes(["Chat.ReadWrite", "ChatMessage.Send"])).toEqual([]);
  });
});

describe("module list storage format", () => {
  it("treats missing or empty values as the pre-modules default (To Do)", () => {
    expect(parseModuleList(undefined)).toEqual(["todo"]);
    expect(parseModuleList(null)).toEqual(["todo"]);
    expect(parseModuleList("  ")).toEqual(["todo"]);
  });

  it("drops unknown ids and duplicates", () => {
    expect(parseModuleList("mail, bogus,mail,files")).toEqual(["mail", "files"]);
  });

  it("serializes in a stable order", () => {
    expect(serializeModuleList(["teams", "todo", "mail"])).toBe("todo,mail,teams");
    expect(parseModuleList(serializeModuleList(["files", "sites"]))).toEqual(["files", "sites"]);
  });
});

const rows: Record<string, Map<string, Record<string, unknown>>> = { profiles: new Map(), connections: new Map() };
vi.mock("../../src/storage/clients.js", () => ({
  TABLE_NAMES: { profiles: "profiles", connections: "connections" },
  getTableClient: (name: string) => ({
    async getEntity(pk: string, rk: string) {
      const row = rows[name]?.get(rk);
      if (!row) throw Object.assign(new Error("not found"), { statusCode: 404 });
      return { partitionKey: pk, rowKey: rk, ...row };
    },
  }),
}));

const { loadProfileContext, hasModule } = await import("../../src/tools/context.js");

describe("loadProfileContext", () => {
  // A personal To Do connection from before modules existed, and a work
  // connection that has consented to mail, files and Teams.
  rows.connections.set("c1", { connectionId: "c1", alias: "personal", accountType: "personal" });
  rows.connections.set("c2", { connectionId: "c2", alias: "work", accountType: "organizational", grantedModules: "todo,mail,files,teams" });
  rows.profiles.set("nadia", { profileAlias: "nadia", connectionId: "c1", createdAt: "2026-10-01T00:00:00Z" });
  rows.profiles.set("haa", { profileAlias: "haa", connectionId: "c2", createdAt: "x", allowSend: true });

  it("keeps /mcp/{alias} To Do only and without send, whatever the connection has", async () => {
    expect(await loadProfileContext("haa", "todo")).toEqual({ boundConnectionId: "c2", modules: ["todo"], allowSend: false });
    expect(await loadProfileContext("nadia")).toEqual({ boundConnectionId: "c1", modules: ["todo"], allowSend: false });
  });

  it("exposes every granted module and the profile's send flag on /m365/{alias}", async () => {
    expect(await loadProfileContext("haa", "m365")).toEqual({
      boundConnectionId: "c2",
      modules: ["todo", "mail", "files", "teams"],
      allowSend: true,
    });
  });

  it("gives a pre-modules connection only To Do on /m365/{alias}", async () => {
    expect(await loadProfileContext("nadia", "m365")).toEqual({ boundConnectionId: "c1", modules: ["todo"], allowSend: false });
  });

  it("keeps the universal /mcp route To Do only", async () => {
    expect(await loadProfileContext("mcp", "todo")).toEqual({ modules: ["todo"], allowSend: false });
  });

  it("exposes every module but never send on the universal /m365 route", async () => {
    const ctx = await loadProfileContext("mcp", "m365");
    expect(ctx.modules).toEqual([...MODULE_IDS]);
    expect(ctx.allowSend).toBe(false);
    expect(ctx.boundConnectionId).toBeUndefined();
  });

  it("returns no binding for an unknown alias", async () => {
    expect((await loadProfileContext("nobody", "m365")).boundConnectionId).toBeUndefined();
  });

  it("hasModule defaults to To Do when a context carries no module list", () => {
    expect(hasModule({ profileAlias: "x", correlationId: "c" }, "todo")).toBe(true);
    expect(hasModule({ profileAlias: "x", correlationId: "c" }, "mail")).toBe(false);
  });
});
