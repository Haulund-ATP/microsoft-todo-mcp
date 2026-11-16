import { describe, expect, it } from "vitest";
import { captureErrorBody, summarizeJsonRpc, urlHost } from "../../src/logging/mcpDiagnostics.js";

describe("summarizeJsonRpc", () => {
  it("records initialize's protocol version, client and capabilities", () => {
    expect(
      summarizeJsonRpc({
        jsonrpc: "2.0",
        id: 0,
        method: "initialize",
        params: { protocolVersion: "2025-06-18", clientInfo: { name: "claude-ai", version: "0.1.0" }, capabilities: { roots: {}, sampling: {} } },
      })
    ).toEqual({ methods: ["initialize"], batch: false, protocolVersion: "2025-06-18", clientInfo: "claude-ai/0.1.0", clientCapabilities: ["roots", "sampling"] });
  });

  it("names the tool on tools/call but never logs its arguments", () => {
    const s = summarizeJsonRpc({ jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: "mail_search_messages", arguments: { query: "secret salary" } } });
    expect(s).toEqual({ methods: ["tools/call"], batch: false, tool: "mail_search_messages" });
    expect(JSON.stringify(s)).not.toContain("secret");
  });

  it("handles batches, notifications and client responses", () => {
    expect(summarizeJsonRpc([{ method: "notifications/initialized" }, { id: 1, result: {} }])).toEqual({
      methods: ["notifications/initialized", "response"],
      batch: true,
    });
    expect(summarizeJsonRpc(undefined)).toBeUndefined();
  });
});

describe("captureErrorBody", () => {
  function fakeRes(status: number) {
    const written: unknown[] = [];
    const res = {
      statusCode: status,
      write: (chunk: unknown) => (written.push(chunk), true),
      end: (chunk?: unknown) => (written.push(chunk), res),
    };
    return { res, written };
  }

  it("captures the body of error responses, capped", () => {
    const { res } = fakeRes(400);
    const get = captureErrorBody(res as never, 20);
    res.write(Buffer.from('{"error":"Bad Request: Unsupported protocol version"}'));
    res.end();
    expect(get()).toBe('{"error":"Bad Reques');
  });

  it("never captures successful bodies (tool results)", () => {
    const { res, written } = fakeRes(200);
    const get = captureErrorBody(res as never);
    res.end('{"result":"private mail"}');
    expect(get()).toBeUndefined();
    expect(written).toEqual(['{"result":"private mail"}']);
  });
});

describe("urlHost", () => {
  it("keeps only the host", () => {
    expect(urlHost("https://claude.ai/api/mcp/auth_callback?code=x")).toBe("claude.ai");
    expect(urlHost("not a url")).toBe("invalid");
    expect(urlHost(undefined)).toBeUndefined();
  });
});
