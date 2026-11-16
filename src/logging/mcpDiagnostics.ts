import type { Response } from "express";

/**
 * Content-free diagnostics for MCP and OAuth traffic, so a connector that
 * "never finishes connecting" can be explained from the logs alone: which
 * JSON-RPC methods a client sent, which protocol version it asked for, and
 * why the server answered 4xx. Tool arguments, results, tokens and secrets
 * are never logged.
 */

export interface JsonRpcSummary {
  /** JSON-RPC methods in the request body, in order ("response" for client replies). */
  methods: string[];
  batch: boolean;
  /** tools/call only: the tool's name. */
  tool?: string;
  /** initialize only: what the client asked for. */
  protocolVersion?: string;
  clientInfo?: string;
  clientCapabilities?: string[];
}

const MAX_METHODS = 10;

function safeString(value: unknown, max = 80): string | undefined {
  return typeof value === "string" ? value.slice(0, max) : undefined;
}

export function summarizeJsonRpc(body: unknown): JsonRpcSummary | undefined {
  if (body === undefined || body === null || typeof body !== "object") return undefined;
  const batch = Array.isArray(body);
  const messages = (batch ? body : [body]).slice(0, MAX_METHODS) as Array<Record<string, unknown>>;
  const summary: JsonRpcSummary = { methods: [], batch };
  for (const m of messages) {
    if (!m || typeof m !== "object") continue;
    const method = safeString(m.method, 60);
    summary.methods.push(method ?? ("result" in m || "error" in m ? "response" : "unknown"));
    const params = (m.params ?? {}) as Record<string, unknown>;
    if (method === "tools/call") summary.tool = safeString(params.name, 64);
    if (method === "initialize") {
      summary.protocolVersion = safeString(params.protocolVersion, 20);
      const info = (params.clientInfo ?? {}) as Record<string, unknown>;
      const name = safeString(info.name, 60);
      const version = safeString(info.version, 30);
      if (name) summary.clientInfo = version ? `${name}/${version}` : name;
      if (params.capabilities && typeof params.capabilities === "object") {
        summary.clientCapabilities = Object.keys(params.capabilities as object).slice(0, 10);
      }
    }
  }
  return summary;
}

/**
 * Captures the first `limit` characters of the response body when the
 * status is an error (>= 400), e.g. the MCP SDK's "Bad Request: …" JSON-RPC
 * error. Successful bodies (tool results) are never captured.
 */
export function captureErrorBody(res: Response, limit = 500): () => string | undefined {
  let captured = "";
  const grab = (chunk: unknown) => {
    if (res.statusCode < 400 || captured.length >= limit || chunk === undefined || chunk === null) return;
    const text = Buffer.isBuffer(chunk) ? chunk.toString("utf8") : typeof chunk === "string" ? chunk : "";
    captured += text.slice(0, limit - captured.length);
  };
  const originalWrite = res.write.bind(res) as (...args: unknown[]) => boolean;
  const originalEnd = res.end.bind(res) as (...args: unknown[]) => Response;
  (res as unknown as { write: (...args: unknown[]) => boolean }).write = (...args: unknown[]) => {
    if (typeof args[0] !== "function") grab(args[0]);
    return originalWrite(...args);
  };
  (res as unknown as { end: (...args: unknown[]) => Response }).end = (...args: unknown[]) => {
    if (typeof args[0] !== "function") grab(args[0]);
    return originalEnd(...args);
  };
  return () => (captured ? captured : undefined);
}

/** Host of a URL, for logging redirect targets without query strings. */
export function urlHost(url: string | undefined): string | undefined {
  if (!url) return undefined;
  try {
    return new URL(url).host;
  } catch {
    return "invalid";
  }
}
