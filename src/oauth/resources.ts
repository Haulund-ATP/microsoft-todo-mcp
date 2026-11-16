import { loadEnv } from "../config/env.js";
import { isValidAlias } from "../storage/profilesRepo.js";

/**
 * OAuth protected resources (RFC 9728) served by this deployment. Each has
 * its own access-token audience, so a token granted to one connector can
 * never be replayed against another:
 *
 *   todo — https://host                → /mcp, /mcp/{alias}  (To Do only, as before)
 *   m365 — https://host/m365/{alias}   → /m365/{alias}       (one resource per profile)
 *          https://host/m365           → /m365               (universal endpoint)
 *
 * Microsoft 365 profiles get a resource each because clients (Claude)
 * keep one set of credentials per resource: with a shared resource,
 * connecting a second /m365 connector took the credentials away from the
 * first. It also means a token for one account cannot reach another.
 */
export type ResourceKind = "todo" | "m365";

function base(): string {
  return loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
}

function m365Path(alias?: string): string {
  return alias ? `/m365/${alias}` : "/m365";
}

/** The resource identifier, which is also the access token's audience. `alias` only matters for m365. */
export function resourceUrl(kind: ResourceKind, alias?: string): string {
  return kind === "todo" ? base() : `${base()}${m365Path(alias)}`;
}

/** Where the resource's Protected Resource Metadata is served (RFC 9728 path-insertion form). */
export function protectedResourceMetadataUrl(kind: ResourceKind, alias?: string): string {
  return `${base()}/.well-known/oauth-protected-resource${kind === "todo" ? "" : m365Path(alias)}`;
}

/**
 * Maps a client-supplied `resource` to the canonical resource it names, or
 * undefined if it is not this server. Accepts the canonical form with or
 * without a trailing slash; anything under /mcp maps to the To Do resource
 * (clients may send the endpoint URL itself), /m365/{alias} to that
 * profile's resource and /m365 to the universal one.
 */
export function canonicalResource(resource: string): { kind: ResourceKind; url: string; alias?: string } | undefined {
  const b = base();
  const r = resource.replace(/\/$/, "");
  if (r === b || r === `${b}/mcp` || r.startsWith(`${b}/mcp/`)) return { kind: "todo", url: resourceUrl("todo") };
  if (r === `${b}/m365`) return { kind: "m365", url: resourceUrl("m365") };
  if (r.startsWith(`${b}/m365/`)) {
    const alias = r.slice(`${b}/m365/`.length);
    if (isValidAlias(alias)) return { kind: "m365", url: resourceUrl("m365", alias), alias };
  }
  return undefined;
}
