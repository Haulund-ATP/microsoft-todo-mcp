import { Router } from "express";
import { randomUUID } from "node:crypto";
import { generatePkce, generateState, buildMicrosoftAuthorizeUrl, completeMicrosoftSignIn, isAdminConsentError, buildAdminConsentError } from "../graph/upstreamOAuth.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { requireOwner } from "./requireOwner.js";
import { requireCsrf } from "./csrf.js";
import { childLogger } from "../logging/logger.js";
import type { PkcePair } from "../graph/upstreamOAuth.js";

/**
 * Admin-only flow for connecting a new Microsoft account (or re-authorizing
 * an existing one) via the upstream Microsoft OAuth (Authorization Code +
 * PKCE). Distinct from src/admin/ownerLogin.ts (which authenticates the
 * admin) and from src/oauth/* (the MCP-side OAuth 2.1 authorization server
 * facing ChatGPT/Claude).
 */

const pending = new Map<string, { pkce: PkcePair; reauthConnectionId?: string; createdAt: number }>();

function prune(): void {
  const cutoff = Date.now() - 10 * 60 * 1000;
  for (const [state, entry] of pending) if (entry.createdAt < cutoff) pending.delete(state);
}

export const connectRouter = Router();

connectRouter.get("/admin/connect/start", requireOwner, async (req, res) => {
  prune();
  const pkce = generatePkce();
  const state = generateState() + "." + randomUUID();
  const reauthConnectionId = typeof req.query.reauth === "string" ? req.query.reauth : undefined;
  pending.set(state, { pkce, reauthConnectionId, createdAt: Date.now() });
  const url = await buildMicrosoftAuthorizeUrl(pkce, state);
  res.redirect(url);
});

connectRouter.get("/oauth/microsoft/callback", requireOwner, async (req, res) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/microsoft/callback" });
  const { code, state, error, error_description: errorDescription } = req.query;

  if (error) {
    log.warn({ oauthErrorCategory: String(error) }, "Microsoft connect returned error");
    res.status(400).send(`Connection failed: ${String(errorDescription ?? error)}`);
    return;
  }
  if (typeof code !== "string" || typeof state !== "string") {
    res.status(400).json({ error: "invalid_request" });
    return;
  }
  const entry = pending.get(state);
  pending.delete(state);
  if (!entry) {
    res.status(400).json({ error: "invalid_state" });
    return;
  }

  try {
    await completeMicrosoftSignIn(code, entry.pkce, req.correlationId, entry.reauthConnectionId);
    res.redirect("/accounts");
  } catch (err) {
    if (isAdminConsentError(err)) {
      const tenantId = (err as { tenantId?: string }).tenantId ?? "unknown";
      const info = buildAdminConsentError(tenantId, (err as Error).message);
      res.status(403).send(info.message);
      return;
    }
    log.error({ resultStatus: "error" }, "Microsoft sign-in completion failed");
    res.status(500).send("Failed to complete Microsoft sign-in. Please try again.");
  }
});

connectRouter.post("/admin/connect/revoke/:connectionId", requireOwner, requireCsrf, async (req, res) => {
  const connectionId = req.params.connectionId;
  if (!connectionId) {
    res.status(400).json({ error: "missing_connection_id" });
    return;
  }
  const repo = new ConnectionsRepo();
  await repo.update(connectionId, { status: "revoked" });
  res.redirect("/accounts");
});

connectRouter.post("/admin/connect/delete/:connectionId", requireOwner, requireCsrf, async (req, res) => {
  const connectionId = req.params.connectionId;
  if (!connectionId) {
    res.status(400).json({ error: "missing_connection_id" });
    return;
  }
  const repo = new ConnectionsRepo();
  await repo.delete(connectionId);
  res.redirect("/accounts");
});
