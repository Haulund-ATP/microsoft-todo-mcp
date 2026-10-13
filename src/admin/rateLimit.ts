import rateLimit from "express-rate-limit";
import type { Request } from "express";

/**
 * Per-client rate limits, each mounted only on the routes it is meant for
 * (see http/app.ts).
 *
 * Two bugs made these misfire (2026-10-13, Claude Code got 429 on /mcp):
 * - They were mounted with a path-less `app.use`, so every request after
 *   them — including all /mcp traffic — counted against the 20/min
 *   login limit.
 * - They keyed on `req.ip`, which with `trust proxy 1` behind Container
 *   Apps is the platform's internal hop (100.100.x.x), not the client. All
 *   clients shared one budget.
 */

/**
 * Container Apps' Envoy ingress sets `x-envoy-external-address` to the real
 * client address and overwrites any value the client sends (verified
 * against production with a spoofed header). Locally the header is absent
 * and `req.ip` is the client.
 */
export function clientKey(req: Request): string {
  const envoy = req.headers["x-envoy-external-address"];
  return (typeof envoy === "string" && envoy) || req.ip || "unknown";
}

const base = { windowMs: 60_000, standardHeaders: "draft-7", legacyHeaders: false, keyGenerator: clientKey } as const;

/** Interactive sign-in and consent: owner login, authorize/approve, client registration, Microsoft connect. */
export const authRateLimiter = rateLimit({ ...base, limit: 20 });

/**
 * Token exchange and revocation. ChatGPT refreshes once per tool call, in
 * parallel, so this needs headroom; codes and refresh tokens are 256-bit
 * random values, so the limit isn't what stops guessing them.
 */
export const tokenRateLimiter = rateLimit({ ...base, limit: 120 });

/** The /accounts admin UI. */
export const adminRateLimiter = rateLimit({ ...base, limit: 60 });

/**
 * MCP traffic (bearer-protected). An MCP client opening a session sends a
 * burst — initialize, initialized, tools/list, … per server — and a chat
 * turn can run many tool calls, so this only guards against runaway loops.
 */
export const mcpRateLimiter = rateLimit({ ...base, limit: 600 });
