import { Router } from "express";
import { randomUUID } from "node:crypto";
import {
  ADMIN_CONSENT_STATE_PREFIX,
  generatePkce,
  generateState,
  buildMicrosoftAuthorizeUrl,
  buildTenantAdminConsentUrl,
  completeMicrosoftSignIn,
  isAdminConsentError,
  buildAdminConsentError,
} from "../graph/upstreamOAuth.js";
import { isOwner, readOwnerSession } from "./session.js";
import { audit } from "../logging/audit.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { PendingUpstreamAuthRepo } from "../storage/oauthRepo.js";
import { requireOwner } from "./requireOwner.js";
import { requireCsrf } from "./csrf.js";
import { childLogger } from "../logging/logger.js";
import { ALL_ORGANIZATIONAL_MODULES, parseModuleList, serializeModuleList, supportedModules, type ModuleId } from "../graph/modules.js";

/**
 * Admin-only flow for connecting a new Microsoft account (or re-authorizing
 * an existing one) via the upstream Microsoft OAuth (Authorization Code +
 * PKCE). Distinct from src/admin/ownerLogin.ts (which authenticates the
 * admin) and from src/oauth/* (the MCP-side OAuth 2.1 authorization server
 * facing ChatGPT/Claude).
 */

const PENDING_TTL_MS = 10 * 60 * 1000;

export const connectRouter = Router();

/** Reads `modules` from a query string or a repeated checkbox field. */
function modulesFromQuery(value: unknown): ModuleId[] {
  const raw = Array.isArray(value) ? value.join(",") : typeof value === "string" ? value : "";
  return parseModuleList(raw);
}

connectRouter.get("/admin/connect/start", requireOwner, async (req, res) => {
  const pkce = generatePkce();
  const state = generateState() + "." + randomUUID();
  const reauthConnectionId = typeof req.query.reauth === "string" ? req.query.reauth : undefined;
  let modules = modulesFromQuery(req.query.modules);
  if (reauthConnectionId) {
    const existing = await new ConnectionsRepo().get(reauthConnectionId);
    if (!existing) {
      res.status(404).send("Unknown connection.");
      return;
    }
    // Re-request what the connection already has plus the new modules, and
    // never ask a personal account for scopes Microsoft rejects for MSA
    // (that would fail the whole sign-in, To Do included).
    modules = supportedModules([...parseModuleList(existing.grantedModules), ...modules], existing.accountType);
  }
  const now = Date.now();
  await new PendingUpstreamAuthRepo().store({
    state,
    kind: "connect",
    codeVerifier: pkce.verifier,
    reauthConnectionId,
    requestedModules: serializeModuleList(modules),
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + PENDING_TTL_MS).toISOString(),
  });
  const url = await buildMicrosoftAuthorizeUrl(pkce, state, modules);
  res.redirect(url);
});

const TENANT_PATTERN = /^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+)$/i;

/**
 * Tenant-wide admin consent for every module at once. Either for the tenant
 * of an existing organizational connection (`?reauth=<connectionId>`) or for
 * a tenant named by id/domain before any account from it is connected
 * (`?tenant=contoso.com`). After a successful consent the owner is sent
 * straight on to (re)connect the account with all modules — no further
 * consent prompts, since the tenant has approved them.
 */
connectRouter.get("/admin/connect/admin-consent", requireOwner, async (req, res) => {
  const reauthConnectionId = typeof req.query.reauth === "string" ? req.query.reauth : undefined;
  let tenant = typeof req.query.tenant === "string" ? req.query.tenant.trim() : "";
  if (reauthConnectionId) {
    const existing = await new ConnectionsRepo().get(reauthConnectionId);
    if (!existing || existing.accountType !== "organizational") {
      res.status(400).send("Admin consent only applies to a work/school connection.");
      return;
    }
    tenant = existing.tenantId;
  }
  if (!TENANT_PATTERN.test(tenant)) {
    res.status(400).send("Give the tenant as its id (GUID) or a verified domain, e.g. contoso.com.");
    return;
  }
  const state = `${ADMIN_CONSENT_STATE_PREFIX}${generateState()}.${randomUUID()}`;
  const now = Date.now();
  await new PendingUpstreamAuthRepo().store({
    state,
    kind: "admin_consent",
    codeVerifier: "",
    reauthConnectionId,
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + PENDING_TTL_MS).toISOString(),
  });
  res.redirect(buildTenantAdminConsentUrl(tenant, state));
});

function adminConsentPage(title: string, text: string): string {
  const esc = (s: string) => s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
  return `<!doctype html><html lang="da"><head><meta charset="utf-8"><title>${esc(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="font-family: system-ui, sans-serif; max-width: 640px; margin: 3rem auto; padding: 0 1rem;">
<h1>${esc(title)}</h1><p>${esc(text)}</p></body></html>`;
}

/**
 * Microsoft's redirect after a tenant admin consent. Handled before the
 * owner check: the administrator may be someone else (e.g. another
 * organisation's IT admin) who only needs a confirmation page.
 */
connectRouter.get("/oauth/microsoft/callback", async (req, res, next) => {
  const state = typeof req.query.state === "string" ? req.query.state : "";
  const isAdminConsent = req.query.admin_consent !== undefined || state.startsWith(ADMIN_CONSENT_STATE_PREFIX);
  if (!isAdminConsent) {
    next();
    return;
  }
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/microsoft/callback" });
  const entry = state ? await new PendingUpstreamAuthRepo().consume(state) : undefined;
  const pending = entry?.kind === "admin_consent" ? entry : undefined;
  res.set("Content-Type", "text/html; charset=utf-8");

  if (req.query.error || String(req.query.admin_consent).toLowerCase() !== "true") {
    log.warn({ rejectReason: String(req.query.error ?? "admin_consent_not_granted") }, "tenant admin consent not granted");
    res.status(400).send(adminConsentPage("Samtykke ikke givet", "Administratorsamtykket blev ikke givet. Ingen rettigheder er ændret."));
    return;
  }

  audit({ action: "connection.admin_consent", correlationId: req.correlationId, connectionId: pending?.reauthConnectionId, outcome: "success" });
  if (pending && isOwner(readOwnerSession(req))) {
    const modules = serializeModuleList(ALL_ORGANIZATIONAL_MODULES);
    const target = pending.reauthConnectionId
      ? `/admin/connect/start?reauth=${encodeURIComponent(pending.reauthConnectionId)}&modules=${modules}`
      : `/admin/connect/start?modules=${modules}`;
    res.redirect(target);
    return;
  }
  res.send(
    adminConsentPage(
      "Samtykke givet",
      "Administratorsamtykket for hele organisationen er givet. Du kan lukke siden. Ejeren af serveren tilslutter kontoen fra sin kontoside."
    )
  );
});

connectRouter.get("/oauth/microsoft/callback", requireOwner, async (req, res) => {
  const log = childLogger({ correlationId: req.correlationId, route: "/oauth/microsoft/callback" });
  const { code, state, error, error_description: errorDescription } = req.query;

  if (error) {
    log.warn({ rejectReason: String(error) }, "Microsoft connect returned error");
    res.status(400).send(`Connection failed: ${String(errorDescription ?? error)}`);
    return;
  }
  if (typeof code !== "string" || typeof state !== "string") {
    res.status(400).json({ error: "invalid_request" });
    return;
  }
  const entry = await new PendingUpstreamAuthRepo().consume(state);
  if (!entry || entry.kind !== "connect") {
    res.status(400).json({ error: "invalid_state" });
    return;
  }

  try {
    await completeMicrosoftSignIn(
      code,
      entry.codeVerifier,
      req.correlationId,
      entry.reauthConnectionId,
      parseModuleList(entry.requestedModules)
    );
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
