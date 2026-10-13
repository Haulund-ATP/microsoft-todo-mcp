import { Router, type Request } from "express";
import { randomBytes } from "node:crypto";
import { z } from "zod";
import { loadEnv } from "../config/env.js";
import { OAuthClientsRepo, OAuthCodesRepo } from "../storage/oauthRepo.js";
import { ProfilesRepo } from "../storage/profilesRepo.js";
import { isPlausibleCodeChallenge } from "./pkce.js";
import { readOwnerSession, isOwner } from "../admin/session.js";
import { issueCsrfToken, requireCsrf } from "../admin/csrf.js";
import { audit } from "../logging/audit.js";
import { childLogger } from "../logging/logger.js";

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

function fullOriginalUrl(req: Request): string {
  return `${req.baseUrl}${req.path}?${new URLSearchParams(req.query as Record<string, string>).toString()}`;
}

/**
 * The canonical resource (no trailing slash) if `resource` names this
 * server, else undefined. The access token's audience is taken from the
 * stored code and verified exactly against this canonical form, so a
 * client sending "https://host/" must still get a token for "https://host".
 */
function canonicalResource(resource: string): string | undefined {
  const base = loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
  return resource.replace(/\/$/, "") === base ? base : undefined;
}

authorizeRouter.get("/oauth/authorize", async (req, res) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/authorize" });
  const parsed = authorizeQuerySchema.safeParse(req.query);
  if (!parsed.success) {
    log.warn({ oauthErrorCategory: "authorize_invalid_request" }, "authorize request rejected");
    res.status(400).json({ error: "invalid_request", error_description: parsed.error.message });
    return;
  }
  const params = parsed.data;

  if (!canonicalResource(params.resource)) {
    log.warn({ oauthErrorCategory: "authorize_unknown_resource" }, "authorize request rejected");
    res.status(400).json({ error: "invalid_target", error_description: "Unknown resource." });
    return;
  }

  const clientsRepo = new OAuthClientsRepo();
  const client = await clientsRepo.get(params.client_id);
  if (!client) {
    log.warn({ oauthErrorCategory: "authorize_unknown_client" }, "authorize request rejected");
    res.status(400).json({ error: "invalid_client" });
    return;
  }
  const redirectUris = await clientsRepo.listRedirectUris(params.client_id);
  if (!redirectUris.includes(params.redirect_uri)) {
    // Do NOT redirect on a redirect_uri mismatch — that would itself be an
    // open-redirect-adjacent hazard. Fail closed with a JSON error instead.
    log.warn({ oauthErrorCategory: "authorize_redirect_mismatch" }, "authorize request rejected");
    res.status(400).json({ error: "invalid_request", error_description: "redirect_uri does not match a registered value exactly." });
    return;
  }

  const session = readOwnerSession(req);
  if (!isOwner(session)) {
    res.redirect(`/oauth/owner/login?returnTo=${encodeURIComponent(fullOriginalUrl(req))}`);
    return;
  }

  // Display only; a storage hiccup must not block the consent screen.
  const profileAliases = await new ProfilesRepo()
    .list()
    .then((profiles) => profiles.map((p) => p.profileAlias).sort())
    .catch(() => []);

  const csrfToken = issueCsrfToken(req, res);
  res.set("Content-Type", "text/html; charset=utf-8");
  res.send(renderConsentPage(client.clientName, profileAliases, params, csrfToken));
});

authorizeRouter.post("/oauth/authorize/approve", requireCsrf, async (req, res) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/authorize/approve" });
  const parsed = authorizeQuerySchema.safeParse(req.body);
  if (!parsed.success) {
    log.warn({ oauthErrorCategory: "approve_invalid_request" }, "consent approval rejected");
    res.status(400).json({ error: "invalid_request" });
    return;
  }
  const params = parsed.data;
  const session = readOwnerSession(req);
  if (!isOwner(session)) {
    log.warn({ oauthErrorCategory: "approve_login_required" }, "consent approval rejected");
    res.status(401).json({ error: "login_required" });
    return;
  }

  const resource = canonicalResource(params.resource);
  if (!resource) {
    log.warn({ oauthErrorCategory: "approve_unknown_resource" }, "consent approval rejected");
    res.status(400).json({ error: "invalid_target" });
    return;
  }

  const clientsRepo = new OAuthClientsRepo();
  const redirectUris = await clientsRepo.listRedirectUris(params.client_id);
  if (!redirectUris.includes(params.redirect_uri)) {
    log.warn({ oauthErrorCategory: "approve_redirect_mismatch" }, "consent approval rejected");
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
    resource,
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
  // Deliberately NOT a 302: a redirect answering a form POST is governed by
  // CSP form-action for every hop of the chain, including any further
  // redirects the client's own callback issues (e.g. chatgpt.com bouncing
  // elsewhere). Chrome then blocks the navigation silently and the "Allow"
  // button appears to do nothing. A fresh navigation from a rendered page
  // (meta refresh + a visible link as fallback) isn't subject to form-action.
  res.set("Content-Type", "text/html; charset=utf-8");
  res.set("Cache-Control", "no-store");
  res.send(renderRedirectPage(redirectUrl.toString()));
});

const escape = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

const PAGE_STYLE =
  "font-family: system-ui, sans-serif; max-width: 480px; margin: 3rem auto; padding: 0 1rem; line-height: 1.5;";
const BUTTON_STYLE =
  "display: block; width: 100%; padding: 1rem; font-size: 1.15rem; font-weight: 600; border: 0; border-radius: 0.5rem; background: #1a73e8; color: #fff; text-align: center; text-decoration: none; cursor: pointer;";

export function renderRedirectPage(url: string): string {
  return `<!doctype html>
<html lang="da"><head><meta charset="utf-8"><title>Adgang givet</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="0;url=${escape(url)}">
</head><body style="${PAGE_STYLE}">
<h1>Adgang givet</h1>
<p>Du sendes tilbage til appen. Sker der ikke noget, så tryk her:</p>
<p><a href="${escape(url)}" style="${BUTTON_STYLE}">Fortsæt</a></p>
<p style="color: #555;">Når appen viser forbindelsen, kan du lukke denne fane.</p>
</body></html>`;
}

export function renderConsentPage(
  clientName: string,
  profileAliases: string[],
  params: z.infer<typeof authorizeQuerySchema>,
  csrfToken: string
): string {
  const profiles = profileAliases.length
    ? `<p>Adgangen gælder alle profiler på serveren: ${profileAliases.map((a) => `<strong>${escape(a)}</strong>`).join(", ")}.</p>`
    : "";
  return `<!doctype html>
<html lang="da"><head><meta charset="utf-8"><title>Giv ${escape(clientName)} adgang</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
</head><body style="${PAGE_STYLE}">
<h1>Giv adgang</h1>
<p><strong>${escape(clientName)}</strong> beder om adgang til din Microsoft To Do-server.</p>
${profiles}
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
  <button type="submit" style="${BUTTON_STYLE}">Giv adgang</button>
</form>
<p style="color: #555;">Luk siden, hvis du ikke har bedt om dette.</p>
</body></html>`;
}
