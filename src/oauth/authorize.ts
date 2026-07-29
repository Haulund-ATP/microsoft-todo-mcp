import { Router } from "express";
import { randomBytes } from "node:crypto";
import { z } from "zod";
import { loadEnv } from "../config/env.js";
import { OAuthClientsRepo, OAuthCodesRepo } from "../storage/oauthRepo.js";
import { isPlausibleCodeChallenge } from "./pkce.js";
import { readOwnerSession, isOwner } from "../admin/session.js";
import { issueCsrfToken, requireCsrf } from "../admin/csrf.js";
import { audit } from "../logging/audit.js";

/**
 * The MCP-side OAuth 2.1 authorization endpoint. Validates every parameter
 * required for a secure PKCE authorization_code flow, then — since this
 * deployment has exactly one resource owner — shows a one-click consent
 * screen (rather than a full multi-user consent system) before issuing a
 * short-lived, one-time authorization code bound to the exact client,
 * redirect_uri, resource, and PKCE challenge presented.
 */

const authorizeQuerySchema = z.object({
  response_type: z.literal("code"),
  client_id: z.string().min(1),
  redirect_uri: z.string().url(),
  state: z.string().min(1).max(512),
  code_challenge: z.string().refine(isPlausibleCodeChallenge, "malformed code_challenge"),
  code_challenge_method: z.literal("S256"),
  resource: z.string().url(),
  scope: z.string().default("mcp"),
});

export const authorizeRouter = Router();

function fullOriginalUrl(req: import("express").Request): string {
  return `${req.baseUrl}${req.path}?${new URLSearchParams(req.query as Record<string, string>).toString()}`;
}

authorizeRouter.get("/oauth/authorize", async (req, res) => {
  const parsed = authorizeQuerySchema.safeParse(req.query);
  if (!parsed.success) {
    res.status(400).json({ error: "invalid_request", error_description: parsed.error.message });
    return;
  }
  const params = parsed.data;

  const env = loadEnv();
  if (params.resource.replace(/\/$/, "") !== env.PUBLIC_BASE_URL.replace(/\/$/, "")) {
    res.status(400).json({ error: "invalid_target", error_description: "Unknown resource." });
    return;
  }

  const clientsRepo = new OAuthClientsRepo();
  const client = await clientsRepo.get(params.client_id);
  if (!client) {
    res.status(400).json({ error: "invalid_client" });
    return;
  }
  const redirectUris = await clientsRepo.listRedirectUris(params.client_id);
  if (!redirectUris.includes(params.redirect_uri)) {
    // Do NOT redirect on a redirect_uri mismatch — that would itself be an
    // open-redirect-adjacent hazard. Fail closed with a JSON error instead.
    res.status(400).json({ error: "invalid_request", error_description: "redirect_uri does not match a registered value exactly." });
    return;
  }

  const session = readOwnerSession(req);
  if (!isOwner(session)) {
    res.redirect(`/oauth/owner/login?returnTo=${encodeURIComponent(fullOriginalUrl(req))}`);
    return;
  }

  const csrfToken = issueCsrfToken(res);
  // The consent form both posts to (same-origin) and — on approval — is
  // redirected by the server to the client's redirect_uri, which is
  // virtually never same-origin (ChatGPT/Claude/etc. all live elsewhere).
  // CSP3's form-action restricts that follow-up redirect too, so the
  // default helmet `form-action 'self'` would block every real client.
  // redirect_uri was already validated above against this exact client's
  // registered list, so it's safe to widen form-action to that one origin
  // for this specific response only.
  const redirectOrigin = new URL(params.redirect_uri).origin;
  res.set(
    "Content-Security-Policy",
    `default-src 'self'; base-uri 'self'; form-action 'self' ${redirectOrigin}; frame-ancestors 'none'; object-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'`
  );
  res.set("Content-Type", "text/html; charset=utf-8");
  res.send(renderConsentPage(client.clientName, params, csrfToken));
});

authorizeRouter.post("/oauth/authorize/approve", requireCsrf, async (req, res) => {
  const parsed = authorizeQuerySchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "invalid_request" });
    return;
  }
  const params = parsed.data;
  const session = readOwnerSession(req);
  if (!isOwner(session)) {
    res.status(401).json({ error: "login_required" });
    return;
  }

  const clientsRepo = new OAuthClientsRepo();
  const redirectUris = await clientsRepo.listRedirectUris(params.client_id);
  if (!redirectUris.includes(params.redirect_uri)) {
    res.status(400).json({ error: "invalid_request" });
    return;
  }

  const code = randomBytes(32).toString("base64url");
  const codesRepo = new OAuthCodesRepo();
  const now = Date.now();
  await codesRepo.store(code, {
    clientId: params.client_id,
    redirectUri: params.redirect_uri,
    codeChallenge: params.code_challenge,
    codeChallengeMethod: "S256",
    resource: params.resource,
    scope: params.scope,
    subject: session!.ownerClaim,
    state: params.state,
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + loadEnv().MCP_OAUTH_AUTH_CODE_TTL_SECONDS * 1000).toISOString(),
    consumed: false,
  });

  audit({ action: "oauth.token_issued", correlationId: req.correlationId, outcome: "success", detail: "auth_code" });

  const redirectUrl = new URL(params.redirect_uri);
  redirectUrl.searchParams.set("code", code);
  redirectUrl.searchParams.set("state", params.state);
  res.redirect(redirectUrl.toString());
});

function renderConsentPage(
  clientName: string,
  params: z.infer<typeof authorizeQuerySchema>,
  csrfToken: string
): string {
  const escape = (s: string) => s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Authorize ${escape(clientName)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
</head><body style="font-family: system-ui, sans-serif; max-width: 480px; margin: 4rem auto; padding: 0 1rem;">
<h1>Authorize access</h1>
<p><strong>${escape(clientName)}</strong> is requesting access to your Microsoft To Do MCP server.</p>
<form method="post" action="/oauth/authorize/approve">
  <input type="hidden" name="_csrf" value="${escape(csrfToken)}">
  <input type="hidden" name="response_type" value="${escape(params.response_type)}">
  <input type="hidden" name="client_id" value="${escape(params.client_id)}">
  <input type="hidden" name="redirect_uri" value="${escape(params.redirect_uri)}">
  <input type="hidden" name="state" value="${escape(params.state)}">
  <input type="hidden" name="code_challenge" value="${escape(params.code_challenge)}">
  <input type="hidden" name="code_challenge_method" value="${escape(params.code_challenge_method)}">
  <input type="hidden" name="resource" value="${escape(params.resource)}">
  <input type="hidden" name="scope" value="${escape(params.scope)}">
  <button type="submit" style="padding: 0.6rem 1.2rem; font-size: 1rem;">Allow</button>
</form>
</body></html>`;
}
