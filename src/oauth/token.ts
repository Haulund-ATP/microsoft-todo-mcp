import { Router } from "express";
import { randomBytes } from "node:crypto";
import { z } from "zod";
import { loadEnv } from "../config/env.js";
import { OAuthClientsRepo, OAuthCodesRepo, OAuthTokensRepo } from "../storage/oauthRepo.js";
import { verifyPkceS256 } from "./pkce.js";
import { signAccessToken } from "./jwt.js";
import { audit } from "../logging/audit.js";
import { childLogger } from "../logging/logger.js";

/**
 * Token endpoint: authorization_code grant (with mandatory PKCE
 * verification) and refresh_token grant (with rotation — every refresh
 * both issues a new refresh token AND invalidates the one just used;
 * reuse of an already-rotated refresh token revokes the entire descendant
 * chain, treating it as a replay/compromise signal).
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
  const parsed = tokenRequestSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "invalid_request", error_description: parsed.error.message });
    return;
  }
  const body = parsed.data;
  const env = loadEnv();
  const clientsRepo = new OAuthClientsRepo();
  const client = await clientsRepo.get(body.client_id);
  if (!client) {
    res.status(400).json({ error: "invalid_client" });
    return;
  }

  if (body.grant_type === "authorization_code") {
    const codesRepo = new OAuthCodesRepo();
    const record = await codesRepo.consume(body.code);
    if (!record) {
      res.status(400).json({ error: "invalid_grant", error_description: "Unknown, expired, or already-used code." });
      return;
    }
    if (record.clientId !== body.client_id || record.redirectUri !== body.redirect_uri) {
      res.status(400).json({ error: "invalid_grant" });
      return;
    }
    if (!verifyPkceS256(body.code_verifier, record.codeChallenge)) {
      log.warn({ oauthErrorCategory: "pkce_mismatch" }, "PKCE verification failed");
      res.status(400).json({ error: "invalid_grant", error_description: "PKCE verification failed." });
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
  const tokensRepo = new OAuthTokensRepo();
  const existing = await tokensRepo.get(body.refresh_token);
  if (!existing || existing.clientId !== body.client_id) {
    res.status(400).json({ error: "invalid_grant" });
    return;
  }
  if (existing.revoked || new Date(existing.expiresAt).getTime() < Date.now()) {
    res.status(400).json({ error: "invalid_grant", error_description: "Refresh token expired or revoked." });
    return;
  }
  if (existing.used) {
    // Replay of an already-rotated refresh token: treat as compromise,
    // revoke the whole chain descending from it.
    audit({ action: "oauth.refresh_reused", correlationId: req.correlationId, outcome: "denied" });
    await tokensRepo.revoke(body.refresh_token);
    await tokensRepo.revokeChainFrom(existing.tokenId);
    res.status(400).json({ error: "invalid_grant", error_description: "Refresh token reuse detected; session revoked." });
    return;
  }

  await tokensRepo.markUsed(body.refresh_token);
  const accessToken = await signAccessToken({
    subject: existing.subject,
    clientId: existing.clientId,
    scope: existing.scope,
    resource: existing.resource,
  });
  const newRefreshToken = randomBytes(32).toString("base64url");
  const env2 = loadEnv();
  const now = Date.now();
  await tokensRepo.store(newRefreshToken, {
    clientId: existing.clientId,
    subject: existing.subject,
    resource: existing.resource,
    scope: existing.scope,
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + env2.MCP_OAUTH_REFRESH_TOKEN_TTL_SECONDS * 1000).toISOString(),
    revoked: false,
    used: false,
    rotatedFrom: existing.tokenId,
  });

  audit({ action: "oauth.token_issued", correlationId: req.correlationId, outcome: "success", detail: "refresh_rotation" });
  res.json({
    access_token: accessToken,
    token_type: "Bearer",
    expires_in: env2.MCP_OAUTH_ACCESS_TOKEN_TTL_SECONDS,
    refresh_token: newRefreshToken,
    scope: existing.scope,
  });
});
