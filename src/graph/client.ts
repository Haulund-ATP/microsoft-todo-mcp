import { Client, MiddlewareFactory, type AuthenticationProvider, type Middleware } from "@microsoft/microsoft-graph-client";
import { getAccessTokenForConnection } from "./upstreamOAuth.js";
import { GraphDiagnosticsMiddleware } from "./diagnostics.js";

/**
 * Builds a Graph SDK client scoped to a single connection's access token.
 * Inserts `GraphDiagnosticsMiddleware` just before the terminal HTTP handler
 * so it can tag every request with a `client-request-id` and capture the
 * Graph-assigned `request-id` + status for diagnostic logging, without
 * otherwise altering the SDK's default middleware chain (auth/retry/
 * redirect/telemetry).
 */
export function createGraphClientForConnection(connectionId: string): Client {
  const authProvider: AuthenticationProvider = {
    getAccessToken: () => getAccessTokenForConnection(connectionId),
  };
  const chain: Middleware[] = MiddlewareFactory.getDefaultMiddlewareChain(authProvider);
  chain.splice(chain.length - 1, 0, new GraphDiagnosticsMiddleware() as unknown as Middleware);
  return Client.initWithMiddleware({ middleware: chain });
}

export interface GraphErrorInfo {
  statusCode: number;
  code?: string;
  message: string;
  retryAfterSeconds?: number;
}

export function toGraphErrorInfo(err: unknown): GraphErrorInfo {
  const e = err as {
    statusCode?: number;
    code?: string;
    message?: string;
    headers?: Record<string, string>;
  };
  const retryAfterHeader = e.headers?.["retry-after"] ?? e.headers?.["Retry-After"];
  return {
    statusCode: e.statusCode ?? 500,
    code: e.code,
    message: e.message ?? "Unknown Graph error",
    retryAfterSeconds: retryAfterHeader ? Number(retryAfterHeader) : undefined,
  };
}

export function isThrottled(err: unknown): boolean {
  return toGraphErrorInfo(err).statusCode === 429;
}

export function isConsentRequired(err: unknown): boolean {
  const info = toGraphErrorInfo(err);
  return info.statusCode === 403 && (info.code === "ErrorAccessDenied" || info.message.includes("consent"));
}
