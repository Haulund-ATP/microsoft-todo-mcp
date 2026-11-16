import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolContext } from "../../src/tools/context.js";

const h = vi.hoisted(() => ({
  client: undefined as unknown,
  audits: [] as Array<{ action: string; connectionId: string; detail: string }>,
  modules: [] as string[],
}));

vi.mock("../../src/tools/m365/common.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../src/tools/m365/common.js")>();
  return {
    ...actual,
    runM365Tool: async (_ctx: unknown, module: string, _name: string, _cid: unknown, fn: (c: unknown, id: string) => Promise<unknown>) => {
      h.modules.push(module);
      try {
        return { content: [{ type: "text", text: JSON.stringify(await fn(h.client, "conn-1")) }] };
      } catch (err) {
        return { isError: true, content: [{ type: "text", text: (err as Error).message }] };
      }
    },
    auditM365: (_ctx: unknown, action: string, connectionId: string, detail: string) => {
      h.audits.push({ action, connectionId, detail });
    },
  };
});

import * as rules from "../../src/graph/m365/mailRules.js";
import { registerMailRuleTools } from "../../src/tools/m365/mailRuleTools.js";

interface Call {
  path: string;
  method?: "get" | "post" | "patch" | "delete";
  body?: unknown;
  select?: string;
}

function fakeGraph(respond: (call: Call) => unknown = () => ({})) {
  const calls: Call[] = [];
  const client = {
    api(path: string) {
      const call: Call = { path };
      const send = async (method: Call["method"], body?: unknown) => {
        call.method = method;
        if (body !== undefined) call.body = body;
        calls.push(call);
        return respond(call);
      };
      const builder = {
        select: (v: string) => ((call.select = v), builder),
        get: () => send("get"),
        post: (body: unknown) => send("post", body),
        patch: (body: unknown) => send("patch", body),
        delete: () => send("delete"),
      };
      return builder;
    },
  } as unknown as Client;
  return { client, calls };
}

const RULES = "/me/mailFolders/inbox/messageRules";
const who = (address: string) => ({ emailAddress: { address } });

/** Answers folder lookups, rule listings and echoes posts/patches back as the saved rule. */
function mailbox(existing: Array<Record<string, unknown>> = []) {
  return fakeGraph((call) => {
    if (call.path.startsWith("/me/mailFolders/") && !call.path.startsWith(RULES)) {
      const ref = decodeURIComponent(call.path.slice("/me/mailFolders/".length));
      if (ref === "missing") throw Object.assign(new Error("not found"), { statusCode: 404 });
      return { id: ref === "archive" ? "AAMkArchiveId" : ref, displayName: ref === "archive" ? "Archive" : `Folder ${ref}` };
    }
    if (call.path === RULES && call.method === "get") return { value: existing };
    if (call.path === RULES && call.method === "post") return { id: "r-new", ...(call.body as object) };
    if (call.path.startsWith(`${RULES}/`) && call.method === "get") return existing.find((r) => call.path.endsWith(`/${r.id as string}`)) ?? { id: "r1" };
    if (call.method === "patch") return { id: "r1", ...(call.body as object) };
    return undefined;
  });
}

beforeEach(() => {
  h.client = undefined;
  h.audits.length = 0;
  h.modules.length = 0;
});

describe("createRule request shaping", () => {
  it("maps snake_case to Graph, resolves the folder and puts the rule after existing ones", async () => {
    const g = mailbox([{ id: "a", sequence: 2 }, { id: "b", sequence: 5 }]);
    const rule = await rules.createRule(
      g.client,
      {
        displayName: "Newsletters",
        conditions: { sender_contains: ["@news.example.com"], from_addresses: ["a@example.com"], within_size_range_kb: { min: 10 } },
        exceptions: { importance: "high" },
        actions: { move_to_folder: "archive", mark_as_read: true, stop_processing_rules: true },
      },
      { allowForward: false }
    );
    const post = g.calls.find((c) => c.method === "post")!;
    expect(post.path).toBe(RULES);
    expect(post.body).toEqual({
      displayName: "Newsletters",
      sequence: 6,
      isEnabled: true,
      conditions: { senderContains: ["@news.example.com"], fromAddresses: [who("a@example.com")], withinSizeRange: { minimumSize: 10 } },
      exceptions: { importance: "high" },
      actions: { moveToFolder: "AAMkArchiveId", markAsRead: true, stopProcessingRules: true },
    });
    expect(rule).toMatchObject({
      id: "r-new",
      display_name: "Newsletters",
      conditions: { sender_contains: ["@news.example.com"], from_addresses: ["a@example.com"], within_size_range_kb: { min: 10 } },
      actions: { move_to_folder: "AAMkArchiveId", mark_as_read: true, stop_processing_rules: true },
      folder_names: { AAMkArchiveId: "Archive" },
    });
  });

  it("uses sequence 1 in an empty mailbox and an explicit sequence without listing", async () => {
    const empty = mailbox();
    await rules.createRule(empty.client, { displayName: "x", actions: { mark_as_read: true } }, { allowForward: false });
    expect((empty.calls.find((c) => c.method === "post")!.body as { sequence: number }).sequence).toBe(1);
    const explicit = mailbox();
    await rules.createRule(explicit.client, { displayName: "x", sequence: 3, actions: { mark_as_read: true } }, { allowForward: false });
    expect(explicit.calls.filter((c) => c.path === RULES && c.method === "get")).toHaveLength(0);
  });

  it("fails before saving when the folder does not exist", async () => {
    const g = mailbox();
    await expect(rules.createRule(g.client, { displayName: "x", actions: { move_to_folder: "missing" } }, { allowForward: false })).rejects.toThrow(
      /Mail folder not found: missing/
    );
    expect(g.calls.some((c) => c.method === "post")).toBe(false);
  });

  it("refuses forwarding unless allowed, and accepts it when allowed", async () => {
    const g = mailbox();
    await expect(rules.createRule(g.client, { displayName: "x", actions: { forward_to: ["out@example.com"] } }, { allowForward: false })).rejects.toThrow(
      /only allowed on a profile that allows sending/
    );
    expect(g.calls.some((c) => c.method === "post")).toBe(false);
    await rules.createRule(g.client, { displayName: "x", actions: { redirect_to: ["out@example.com"] } }, { allowForward: true });
    expect(g.calls.find((c) => c.method === "post")!.body).toMatchObject({ actions: { redirectTo: [who("out@example.com")] } });
  });

  it("rejects rules without actions, permanent delete, delete plus move and bad size ranges", async () => {
    const g = mailbox();
    const opts = { allowForward: true };
    await expect(rules.buildActions(g.client, {}, opts)).rejects.toThrow(/at least one action/);
    await expect(rules.buildActions(g.client, { permanent_delete: true } as never, opts)).rejects.toThrow(/cannot delete mail permanently/);
    await expect(rules.buildActions(g.client, { delete: true, move_to_folder: "archive" }, opts)).rejects.toThrow(/either delete or move/);
    expect(() => rules.buildPredicates({ within_size_range_kb: { min: 50, max: 10 } })).toThrow(/min must not be larger/);
    expect(() => rules.buildPredicates({ within_size_range_kb: {} })).toThrow(/needs min, max/);
  });
});

describe("listRules output shaping", () => {
  it("sorts by sequence, drops empty/false values, keeps unknown true predicates and names folders", async () => {
    const g = mailbox([
      {
        id: "r2",
        displayName: "Second",
        sequence: 2,
        isEnabled: false,
        conditions: { subjectContains: ["invoice"], isEncrypted: true, isSigned: false, bodyContains: [] },
        actions: { copyToFolder: "f9", markAsRead: false, forwardTo: [who("x@example.com")] },
      },
      { id: "r1", displayName: "First", sequence: 1, isEnabled: true, hasError: true, conditions: null, actions: { delete: true } },
    ]);
    const { rules: list } = await rules.listRules(g.client);
    expect(list.map((r) => r.id)).toEqual(["r1", "r2"]);
    expect(list[0]).toEqual({ id: "r1", display_name: "First", sequence: 1, is_enabled: true, has_error: true, actions: { delete: true } });
    expect(list[1]).toEqual({
      id: "r2",
      display_name: "Second",
      sequence: 2,
      is_enabled: false,
      conditions: { subject_contains: ["invoice"], isEncrypted: true },
      actions: { copy_to_folder: "f9", forward_to: ["x@example.com"] },
      folder_names: { f9: "Folder f9" },
    });
  });
});

describe("updateRule and deleteRule", () => {
  it("patches only the given fields", async () => {
    const g = mailbox([{ id: "r1", sequence: 1, isEnabled: true, actions: { markAsRead: true } }]);
    await rules.updateRule(g.client, "r1", { isEnabled: false, conditions: { subject_contains: ["x"] } }, { allowForward: false });
    const patch = g.calls.find((c) => c.method === "patch")!;
    expect(patch.path).toBe(`${RULES}/r1`);
    expect(patch.body).toEqual({ isEnabled: false, conditions: { subjectContains: ["x"] } });
    await expect(rules.updateRule(g.client, "r1", {}, { allowForward: false })).rejects.toThrow(/Nothing to update/);
  });

  it("refuses read-only rules and switching on a forwarding rule without send permission", async () => {
    const g = mailbox([
      { id: "ro", isReadOnly: true, actions: { markAsRead: true } },
      { id: "fw", isEnabled: false, actions: { forwardTo: [who("out@example.com")] } },
    ]);
    await expect(rules.updateRule(g.client, "ro", { isEnabled: false }, { allowForward: true })).rejects.toThrow(/read-only/);
    await expect(rules.updateRule(g.client, "fw", { isEnabled: true }, { allowForward: false })).rejects.toThrow(/can only be switched on/);
    // Turning it off, or replacing its actions with non-forwarding ones, is fine.
    await rules.updateRule(g.client, "fw", { isEnabled: false }, { allowForward: false });
    await rules.updateRule(g.client, "fw", { isEnabled: true, actions: { mark_as_read: true } }, { allowForward: false });
    expect(g.calls.filter((c) => c.method === "patch")).toHaveLength(2);
    await expect(rules.deleteRule(g.client, "ro")).rejects.toThrow(/read-only/);
    expect(g.calls.some((c) => c.method === "delete")).toBe(false);
  });

  it("returns the deleted rule's definition", async () => {
    const g = mailbox([{ id: "r1", displayName: "Old", sequence: 4, isEnabled: true, actions: { moveToFolder: "f1" } }]);
    const result = await rules.deleteRule(g.client, "r1");
    expect(g.calls.map((c) => [c.method, c.path])).toEqual([
      ["get", `${RULES}/r1`],
      ["delete", `${RULES}/r1`],
    ]);
    expect(result).toEqual({ deleted: true, removed_rule: { id: "r1", display_name: "Old", sequence: 4, is_enabled: true, actions: { move_to_folder: "f1" } } });
  });
});

function fakeServer() {
  const tools = new Map<
    string,
    { config: { description: string; inputSchema: Record<string, { safeParse: (v: unknown) => { success: boolean } }>; annotations?: Record<string, boolean> }; handler: (a: Record<string, unknown>) => Promise<unknown> }
  >();
  return {
    registerTool: (name: string, config: never, handler: (a: Record<string, unknown>) => Promise<unknown>) => {
      tools.set(name, { config, handler });
    },
    tools,
  };
}

function register(allowSend: boolean) {
  const server = fakeServer();
  const ctx: ToolContext = { profileAlias: "work", boundConnectionId: "conn-1", correlationId: "corr-1", modules: ["mail_rules"], allowSend };
  registerMailRuleTools(server as never, ctx);
  return server;
}

type ToolResult = { isError?: boolean; content: Array<{ text: string }> };

describe("registration", () => {
  it("registers the four rule tools on the mail_rules module", () => {
    expect([...register(false).tools.keys()]).toEqual(["mail_list_rules", "mail_create_rule", "mail_update_rule", "mail_delete_rule"]);
    expect(register(false).tools.get("mail_delete_rule")!.config.annotations).toMatchObject({ destructiveHint: true });
  });

  it("only offers forward actions in the schema when the profile allows sending", () => {
    const forward = { forward_to: ["out@example.com"] };
    const noSend = register(false).tools.get("mail_create_rule")!.config;
    expect(noSend.inputSchema.actions!.safeParse(forward).success).toBe(false);
    expect(noSend.description).toMatch(/not available on this profile/);
    const send = register(true).tools.get("mail_create_rule")!.config;
    expect(send.inputSchema.actions!.safeParse(forward).success).toBe(true);
    expect(send.description).toMatch(/explicitly asked for forwarding/);
  });

  it("rejects unknown condition and action keys instead of silently dropping them", () => {
    const schema = register(true).tools.get("mail_create_rule")!.config.inputSchema;
    expect(schema.actions!.safeParse({ permanent_delete: true }).success).toBe(false);
    expect(schema.conditions!.safeParse({ subjectContains: ["x"] }).success).toBe(false);
  });
});

describe("tool handlers", () => {
  it("mail_create_rule audits counts only", async () => {
    const g = mailbox();
    h.client = g.client;
    const result = (await register(true).tools.get("mail_create_rule")!.handler({
      display_name: "Secret project",
      conditions: { subject_contains: ["salary"] },
      actions: { forward_to: ["a@example.com", "A@example.com", "b@example.com"] },
      is_enabled: true,
    })) as ToolResult;
    expect(result.isError).toBeUndefined();
    expect(h.modules).toEqual(["mail_rules"]);
    expect(h.audits).toEqual([{ action: "m365.rule", connectionId: "conn-1", detail: "mail_rule created forward_recipients=2" }]);
    expect(JSON.stringify(h.audits)).not.toMatch(/Secret|salary|example\.com/);
  });

  it("mail_delete_rule audits the delete with the removed rule's forward count", async () => {
    h.client = mailbox([{ id: "r1", actions: { redirectTo: [who("x@example.com")] } }]).client;
    const result = (await register(false).tools.get("mail_delete_rule")!.handler({ rule_id: "r1" })) as ToolResult;
    expect(result.isError).toBeUndefined();
    expect(h.audits).toEqual([{ action: "m365.rule", connectionId: "conn-1", detail: "mail_rule deleted forward_recipients=1" }]);
  });

  it("does not audit a refused change", async () => {
    h.client = mailbox().client;
    const result = (await register(false).tools.get("mail_create_rule")!.handler({ display_name: "x", actions: {}, is_enabled: true })) as ToolResult;
    expect(result.isError).toBe(true);
    expect(h.audits).toEqual([]);
  });
});
