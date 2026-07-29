import { randomBytes, createHash } from "node:crypto";
import { loadEnv } from "../config/env.js";
import { createMsalAppStateless, createMsalAppForConnection, GRAPH_SCOPES } from "./msalClient.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { childLogger } from "../logging/logger.js";
import { audit } from "../logging/audit.js";

/**
 * Upstream (Microsoft Entra ID / Graph) Authorization Code + PKCE (S256)
 * flow. This is distinct from the MCP-side OAuth 2.1 authorization server
 * in src/oauth/ — this module is the *client* leg talking to Microsoft.
 */

export interface PkcePair {
  verifier: string;
  challenge: string;
}

export function generatePkce(): PkcePair {
  const verifier = base64UrlEncode(randomBytes(32));
  const challenge = base64UrlEncode(createHash("sha256").update(verifier).digest());
  return { verifier, challenge };
}

function base64UrlEncode(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

export function generateState(): string {
  return base64UrlEncode(randomBytes(24));
}

function redirectUri(): string {
  return `${loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "")}/oauth/microsoft/callback`;
}

/** Builds the Microsoft authorize URL the admin browser is redirected to. */
export async function buildMicrosoftAuthorizeUrl(pkce: PkcePair, state: string): Promise<string> {
  const app = await createMsalAppStateless();
  return app.getAuthCodeUrl({
    scopes: GRAPH_SCOPES,
    redirectUri: redirectUri(),
    state,
    codeChallenge: pkce.challenge,
    codeChallengeMethod: "S256",
    prompt: "select_account",
  });
}

export interface AdminConsentRequiredError {
  kind: "admin_consent_required";
  tenantId: string;
  adminConsentUrl: string;
  message: string;
}

export function isAdminConsentError(err: unknown): err is Error & { errorCode?: string } {
  const message = (err as { errorMessage?: string; message?: string })?.errorMessage
    ?? (err as { message?: string })?.message
    ?? "";
  return (
    message.includes("AADSTS65001") || // user/admin has not consented
    message.includes("AADSTS90094") || // admin consent required
    message.includes("consent_required")
  );
}

export function buildAdminConsentError(tenantId: string, _originalErrorMessage: string): AdminConsentRequiredError {
  const env = loadEnv();
  const clientId = env.AZURE_CLIENT_ID;
  const adminConsentUrl =
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/adminconsent` +
    `?client_id=${encodeURIComponent(clientId)}` +
    `&redirect_uri=${encodeURIComponent(redirectUri())}`;
  return {
    kind: "admin_consent_required",
    tenantId,
    adminConsentUrl,
    message: `This Microsoft 365 tenant (${tenantId}) requires an administrator to consent to this app before it can be used. Ask a tenant administrator to open: ${adminConsentUrl}`,
  };
}

export interface CompletedSignIn {
  homeAccountId: string;
  tenantId: string;
  objectId: string;
  displayName: string;
  maskedEmail: string;
  accountType: "personal" | "organizational";
}

function maskEmail(email: string | undefined | null): string {
  if (!email || !email.includes("@")) return "unknown";
  const [localPart, domain] = email.split("@");
  const local = localPart ?? "";
  const visible = local.slice(0, 1);
  return `${visible}${"*".repeat(Math.max(local.length - 1, 3))}@${domain ?? "unknown"}`;
}

/**
 * Exchanges the authorization code for tokens using a *stateless* MSAL app
 * first (no connection exists yet), then persists the resulting cache to a
 * brand-new connection record and re-homes the cache under that
 * connection's id-bound cache plugin.
 */
export async function completeMicrosoftSignIn(
  code: string,
  codeVerifier: string,
  correlationId: string,
  reauthConnectionId?: string
): Promise<CompletedSignIn> {
  const log = childLogger({ correlationId });
  const app = await createMsalAppStateless();
  let result;
  try {
    result = await app.acquireTokenByCode({
      code,
      scopes: GRAPH_SCOPES,
      redirectUri: redirectUri(),
      codeVerifier,
    });
  } catch (err) {
    if (isAdminConsentError(err)) {
      log.warn({ oauthErrorCategory: "admin_consent_required" }, "Microsoft sign-in requires admin consent");
    }
    throw err;
  }

  const account = result.account;
  if (!account) throw new Error("Microsoft sign-in did not return an account.");

  const tenantId = account.tenantId ?? "9188040d-6c67-4c5b-b112-36a304b66dad"; // MSA "consumers" pseudo-tenant
  const accountType: "personal" | "organizational" =
    tenantId === "9188040d-6c67-4c5b-b112-36a304b66dad" ? "personal" : "organizational";

  const repo = new ConnectionsRepo();
  let connection;
  if (reauthConnectionId) {
    const existing = await repo.get(reauthConnectionId);
    if (!existing) throw new Error(`Unknown connection for reauth: ${reauthConnectionId}`);
    await repo.update(reauthConnectionId, {
      tenantId,
      msalHomeAccountId: account.homeAccountId,
      objectId: account.localAccountId,
      displayName: account.name ?? account.username ?? existing.displayName,
      maskedEmail: maskEmail(account.username),
      status: "active",
      errorDetail: undefined,
    });
    connection = { ...existing, tenantId, msalHomeAccountId: account.homeAccountId, objectId: account.localAccountId };
  } else {
    const alias = await nextAvailableAlias(repo, accountType);
    connection = await repo.create({
      alias,
      accountType,
      tenantId,
      msalHomeAccountId: account.homeAccountId,
      objectId: account.localAccountId,
      displayName: account.name ?? account.username ?? "Microsoft account",
      maskedEmail: maskEmail(account.username),
      lastUsedAt: null,
      status: "active",
    });
  }

  // Re-run acquisition through the connection-bound app so the cache
  // plugin persists it under the permanent connectionId/blob mapping.
  const boundApp = await createMsalAppForConnection(connection.connectionId, repo);
  await boundApp.acquireTokenSilent({ account, scopes: GRAPH_SCOPES }).catch(async () => {
    // First silent call may miss cache timing; force a cache write by
    // re-deserializing what the stateless app already has.
    const serialized = app.getTokenCache().serialize();
    await repo.saveTokenCache(connection.connectionId, serialized);
  });

  audit({
    action: reauthConnectionId ? "connection.reauth" : "connection.created",
    correlationId,
    connectionId: connection.connectionId,
    outcome: "success",
  });

  return {
    homeAccountId: account.homeAccountId,
    tenantId,
    objectId: account.localAccountId,
    displayName: connection.displayName,
    maskedEmail: connection.maskedEmail,
    accountType,
  };
}

async function nextAvailableAlias(repo: ConnectionsRepo, accountType: "personal" | "organizational"): Promise<string> {
  const base = accountType === "personal" ? "personal" : "work";
  const existing = new Set((await repo.list()).map((c) => c.alias));
  if (!existing.has(base)) return base;
  let n = 2;
  while (existing.has(`${base}-${n}`)) n += 1;
  return `${base}-${n}`;
}

/** Acquires a Graph access token for an existing connection, silently refreshing as needed. */
export async function getAccessTokenForConnection(connectionId: string): Promise<string> {
  const repo = new ConnectionsRepo();
  const connection = await repo.get(connectionId);
  if (!connection) throw new Error(`Unknown connection: ${connectionId}`);
  const app = await createMsalAppForConnection(connectionId, repo);
  const accounts = await app.getTokenCache().getAllAccounts();
  const account = accounts.find((a) => a.homeAccountId === connection.msalHomeAccountId);
  if (!account) {
    await repo.update(connectionId, { status: "consent_required" });
    throw new Error(`No cached account for connection ${connectionId}; re-authentication required.`);
  }
  try {
    const result = await app.acquireTokenSilent({ account, scopes: GRAPH_SCOPES });
    await repo.update(connectionId, { lastUsedAt: new Date().toISOString(), status: "active" });
    return result.accessToken;
  } catch (err) {
    if (isAdminConsentError(err)) {
      await repo.update(connectionId, { status: "consent_required" });
    } else {
      await repo.update(connectionId, { status: "error", errorDetail: "silent_token_acquisition_failed" });
    }
    throw err;
  }
}
