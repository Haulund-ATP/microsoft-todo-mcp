import type { IncomingMessage, ServerResponse } from "node:http";

/**
 * Request-log shaping for pino-http. The defaults logged every request and
 * response header plus the raw query string, which meant (a) Microsoft's
 * sign-in `code` from /oauth/owner/callback landed in the logs verbatim and
 * (b) each line was ~1.5 KB, mostly the same security headers repeated.
 * Health/readiness probes (every few seconds) were about 40% of all lines.
 */

/** Query parameters whose values are secrets or account identifiers. */
const SENSITIVE_QUERY_PARAMS = new Set([
  "code",
  "code_verifier",
  "client_info",
  "session_state",
  "refresh_token",
  "access_token",
  "id_token",
  "client_secret",
]);

export function redactUrl(url: string | undefined): string | undefined {
  if (!url) return url;
  const q = url.indexOf("?");
  if (q < 0) return url;
  const params = new URLSearchParams(url.slice(q + 1));
  let changed = false;
  for (const key of [...params.keys()]) {
    if (SENSITIVE_QUERY_PARAMS.has(key.toLowerCase())) {
      params.set(key, "[redacted]");
      changed = true;
    }
  }
  return changed ? `${url.slice(0, q)}?${params.toString()}` : url;
}

const PROBE_PATHS = new Set(["/health", "/ready"]);

export function isProbeRequest(req: IncomingMessage): boolean {
  return PROBE_PATHS.has((req.url ?? "").split("?")[0]!);
}

export const httpLogSerializers = {
  req(req: { id?: unknown; method?: string; url?: string; headers?: IncomingMessage["headers"]; raw?: { ip?: string } }) {
    // pino-http passes pino-std-serializers' request view: url (already
    // originalUrl), headers, and the express request on `raw`.
    return {
      id: req.id,
      method: req.method,
      url: redactUrl(req.url),
      userAgent: req.headers?.["user-agent"],
      clientIp: req.headers?.["x-envoy-external-address"] ?? req.raw?.ip,
    };
  },
  res(res: ServerResponse & { statusCode?: number }) {
    return { statusCode: res.statusCode };
  },
};
