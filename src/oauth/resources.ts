import { loadEnv } from "../config/env.js";

/**
 * This server exposes two OAuth protected resources (RFC 9728), each with its
 * own access-token audience, so a token granted to a To Do connector can never
 * be replayed against the Microsoft 365 endpoint (mail, files, Teams, send):
 *
 *   todo — https://host        → /mcp, /mcp/{alias}   (To Do only, as before)
 *   m365 — https://host/m365   → /m365, /m365/{alias} (every granted module)
 */
export type ResourceKind = "todo" | "m365";

const RESOURCE_PATHS: Record<ResourceKind, string> = { todo: "", m365: "/m365" };

function base(): string {
  return loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
}

/** The resource identifier, which is also the access token's audience. */
export function resourceUrl(kind: ResourceKind): string {
  return `${base()}${RESOURCE_PATHS[kind]}`;
}

/** Where the resource's Protected Resource Metadata is served (RFC 9728 path-insertion form). */
export function protectedResourceMetadataUrl(kind: ResourceKind): string {
  return `${base()}/.well-known/oauth-protected-resource${RESOURCE_PATHS[kind]}`;
}

/**
 * Maps a client-supplied `resource` to the canonical resource it names, or
 * undefined if it is not this server. Accepts the canonical form with or
 * without a trailing slash, and — for clients that send the MCP endpoint URL
 * itself — any path under /m365 (→ m365) or /mcp (→ todo).
 */
export function canonicalResource(resource: string): { kind: ResourceKind; url: string } | undefined {
  const b = base();
  const r = resource.replace(/\/$/, "");
  if (r === b || r === `${b}/mcp` || r.startsWith(`${b}/mcp/`)) return { kind: "todo", url: resourceUrl("todo") };
  if (r === `${b}/m365` || r.startsWith(`${b}/m365/`)) return { kind: "m365", url: resourceUrl("m365") };
  return undefined;
}
