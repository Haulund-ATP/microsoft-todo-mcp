import { Router } from "express";
import { randomUUID } from "node:crypto";
import { loadEnv } from "../config/env.js";
import { createMsalAppStateless } from "../graph/msalClient.js";
import { generatePkce, generateState, type PkcePair } from "../graph/upstreamOAuth.js";
import { setOwnerSessionCookie } from "./session.js";
import { audit } from "../logging/audit.js";
import { childLogger } from "../logging/logger.js";

/**
 * Owner-identity login: a *separate*, minimal-scope (openid + profile only,
 * no Tasks.ReadWrite) Microsoft sign-in used purely to prove "you are the
 * deployment owner" for /accounts and for approving MCP-side OAuth 2.1
 * authorize requests. This is intentionally decoupled from the Graph
 * "connections" sign-in (src/graph/upstreamOAuth.ts) — connecting a
 * Microsoft To Do account and authenticating as the admin are different
 * concerns, even though both go through Microsoft Entra ID.
 */

const OWNER_LOGIN_SCOPES = ["openid", "profile"];

// In-memory PKCE/state store for the short owner-login handshake. A crash
// mid-login just means the user retries; nothing sensitive is at stake.
const pendingLogins = new Map<string, { pkce: PkcePair; returnTo: string; createdAt: number }>();

function redirectUri(): string {
  return `${loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "")}/oauth/owner/callback`;
}

function pruneExpired(): void {
  const cutoff = Date.now() - 5 * 60 * 1000;
  for (const [state, entry] of pendingLogins) {
    if (entry.createdAt < cutoff) pendingLogins.delete(state);
  }
}

export const ownerLoginRouter = Router();

ownerLoginRouter.get("/oauth/owner/login", async (req, res, next) => {
  try {
    pruneExpired();
    const returnTo = typeof req.query.returnTo === "string" ? req.query.returnTo : "/accounts";
    if (!returnTo.startsWith("/")) {
      res.status(400).json({ error: "invalid_return_to" });
      return;
    }
    const pkce = generatePkce();
    const state = generateState() + "." + randomUUID();
    pendingLogins.set(state, { pkce, returnTo, createdAt: Date.now() });

    const app = await createMsalAppStateless();
    const url = await app.getAuthCodeUrl({
      scopes: OWNER_LOGIN_SCOPES,
      redirectUri: redirectUri(),
      state,
      codeChallenge: pkce.challenge,
      codeChallengeMethod: "S256",
      prompt: "select_account",
    });
    res.redirect(url);
  } catch (err) {
    next(err);
  }
});

ownerLoginRouter.get("/oauth/owner/callback", async (req, res, next) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/owner/callback" });
  try {
    const { code, state, error, error_description: errorDescription } = req.query;
    if (error) {
      log.warn({ oauthErrorCategory: String(error) }, "owner login returned error");
      res.status(400).send(`Sign-in failed: ${String(errorDescription ?? error)}`);
      return;
    }
    if (typeof code !== "string" || typeof state !== "string") {
      res.status(400).json({ error: "invalid_request" });
      return;
    }
    const pending = pendingLogins.get(state);
    pendingLogins.delete(state);
    if (!pending) {
      res.status(400).json({ error: "invalid_state" });
      return;
    }

    const app = await createMsalAppStateless();
    const result = await app.acquireTokenByCode({
      code,
      scopes: OWNER_LOGIN_SCOPES,
      redirectUri: redirectUri(),
      codeVerifier: pending.pkce.verifier,
    });

    const env = loadEnv();
    const claims = (result.account?.idTokenClaims ?? {}) as Record<string, unknown>;
    const claimValue = claims[env.ADMIN_OWNER_CLAIM_NAME];
    const ownerClaim = typeof claimValue === "string" ? claimValue : result.account?.localAccountId;

    if (!ownerClaim || ownerClaim !== env.ADMIN_OWNER_CLAIM_VALUE) {
      audit({
        action: "admin.login_denied",
        correlationId: req.correlationId,
        outcome: "denied",
      });
      res.status(403).send("This Microsoft account is not authorized to administer this deployment.");
      return;
    }

    setOwnerSessionCookie(res, { ownerClaim, issuedAt: Date.now() });
    audit({ action: "admin.login", correlationId: req.correlationId, outcome: "success" });
    res.redirect(pending.returnTo);
  } catch (err) {
    next(err);
  }
});
