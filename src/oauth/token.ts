import { Router, type Response } from "express";
import { randomBytes } from "node:crypto";
import { z } from "zod";
import { loadEnv } from "../config/env.js";
import { OAuthClientsRepo, OAuthCodesRepo, OAuthTokensRepo } from "../storage/oauthRepo.js";
import { verifyPkceS256 } from "./pkce.js";
import { signAccessToken } from "./jwt.js";
import { rotateRefreshToken } from "./refreshRotation.js";
import { audit } from "../logging/audit.js";
import { childLogger } from "../logging/logger.js";

/**
 * Token endpoint: authorization_code grant (with mandatory PKCE
 * verification) and refresh_token grant (with rotation — every refresh
 * both issues a new refresh token AND invalidates the one just used;
 * reuse of an already-rotated refresh token outside a short grace window
 * revokes the entire descendant chain, treating it as a replay/compromise
 * signal — see refreshRotation.ts).
 *
 * Every rejection is logged with an `oauthErrorCategory` so a client that
 * suddenly asks the user to reconnect can be explained from the logs alone.
 */

const tokenRequestSchema = z.discriminatedUnion("grant_type", [
  z.object({
    grant_type: z.literal("authorization_code"),
    code: z.string().min(1),
    redirect_uri: z.string().url(),
    client_id: z.string().min(1),
    code_verifier: z.string().min(43).max(128),
    resource: z.string().url().optional(),
  }),
  z.object({
    grant_type: z.literal("refresh_token"),
    refresh_token: z.string().min(1),
    client_id: z.string().min(1),
    resource: z.string().url().optional(),
  }),
]);

export const tokenRouter = Router();

tokenRouter.post("/oauth/token", async (req, res) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/token" });
  const reject = (r: Response, category: string, error: string, description?: string) => {
    log.warn({ oauthErrorCategory: category }, "token request rejected");
    r.status(400).json(description ? { error, error_description: description } : { error });
  };

  const parsed = tokenRequestSchema.safeParse(req.body);
  if (!parsed.success) {
    reject(res, "invalid_request", "invalid_request", parsed.error.message);
    return;
  }
  const body = parsed.data;
  const env = loadEnv();
  const clientsRepo = new OAuthClientsRepo();
  const client = await clientsRepo.get(body.client_id);
  if (!client) {
    reject(res, "unknown_client", "invalid_client");
    return;
  }

  if (body.grant_type === "authorization_code") {
    const codesRepo = new OAuthCodesRepo();
    const consumed = await codesRepo.consume(body.code);
    if ("reason" in consumed) {
      reject(res, `code_${consumed.reason}`, "invalid_grant", "Unknown, expired, or already-used code.");
      return;
    }
    const { record } = consumed;
    if (record.clientId !== body.client_id || record.redirectUri !== body.redirect_uri) {
      reject(res, record.clientId !== body.client_id ? "code_client_mismatch" : "code_redirect_mismatch", "invalid_grant");
      return;
    }
    if (!verifyPkceS256(body.code_verifier, record.codeChallenge)) {
      reject(res, "pkce_mismatch", "invalid_grant", "PKCE verification failed.");
      return;
    }

    const accessToken = await signAccessToken({
      subject: record.subject,
      clientId: record.clientId,
      scope: record.scope,
      resource: record.resource,
    });
    const refreshToken = randomBytes(32).toString("base64url");
    const tokensRepo = new OAuthTokensRepo();
    const now = Date.now();
    await tokensRepo.store(refreshToken, {
      clientId: record.clientId,
      subject: record.subject,
      resource: record.resource,
      scope: record.scope,
      createdAt: new Date(now).toISOString(),
      expiresAt: new Date(now + env.MCP_OAUTH_REFRESH_TOKEN_TTL_SECONDS * 1000).toISOString(),
      revoked: false,
      used: false,
    });

    audit({ action: "oauth.token_issued", correlationId: req.correlationId, outcome: "success", detail: "access_token" });
    res.json({
      access_token: accessToken,
      token_type: "Bearer",
      expires_in: env.MCP_OAUTH_ACCESS_TOKEN_TTL_SECONDS,
      refresh_token: refreshToken,
      scope: record.scope,
    });
    return;
  }

  // refresh_token grant
  const outcome = await rotateRefreshToken(new OAuthTokensRepo(), body.refresh_token, body.client_id, {
    now: Date.now(),
    graceSeconds: env.MCP_OAUTH_REFRESH_REUSE_GRACE_SECONDS,
    ttlSeconds: env.MCP_OAUTH_REFRESH_TOKEN_TTL_SECONDS,
    newToken: () => randomBytes(32).toString("base64url"),
  });
  if (!outcome.ok) {
    const category = `refresh_${outcome.category}`;
    if (outcome.category === "reuse") {
      audit({ action: "oauth.refresh_reused", correlationId: req.correlationId, outcome: "denied" });
      reject(res, category, "invalid_grant", "Refresh token reuse detected; session revoked.");
    } else if (outcome.category === "expired" || outcome.category === "revoked") {
      reject(res, category, "invalid_grant", "Refresh token expired or revoked.");
    } else {
      reject(res, category, "invalid_grant");
    }
    return;
  }
  if (outcome.mode === "grace") {
    audit({ action: "oauth.refresh_concurrent", correlationId: req.correlationId, outcome: "success" });
  }

  const { record } = outcome;
  const accessToken = await signAccessToken({
    subject: record.subject,
    clientId: record.clientId,
    scope: record.scope,
    resource: record.resource,
  });

  audit({ action: "oauth.token_issued", correlationId: req.correlationId, outcome: "success", detail: "refresh_rotation" });
  res.json({
    access_token: accessToken,
    token_type: "Bearer",
    expires_in: env.MCP_OAUTH_ACCESS_TOKEN_TTL_SECONDS,
    refresh_token: outcome.refreshToken,
    scope: record.scope,
  });
});
