/** Domain types shared across storage repositories. */

export type ConnectionStatus = "active" | "expired" | "revoked" | "consent_required" | "error";
export type AccountType = "personal" | "organizational";

export interface Connection {
  connectionId: string;
  alias: string; // URL-safe unique alias, e.g. "personal", "work"
  accountType: AccountType;
  tenantId: string;
  msalHomeAccountId: string;
  objectId: string;
  displayName: string;
  /** Masked, e.g. "t***@outlook.com" — never store/display the full address in logs. */
  maskedEmail: string;
  createdAt: string; // ISO 8601
  lastUsedAt: string | null; // ISO 8601
  status: ConnectionStatus;
  /** Name of the blob holding the AES-GCM-encrypted MSAL token cache. */
  tokenCacheBlobName: string;
  errorDetail?: string;
}

export interface ProfileBinding {
  profileAlias: string; // "personal" | "work" | custom alias; also serves as route segment
  connectionId: string;
  createdAt: string;
}

export interface OAuthClient {
  clientId: string;
  clientName: string;
  redirectUris: string[];
  registrationMethod: "dynamic" | "client_id_metadata_document" | "pre_registered";
  createdAt: string;
  /** Only set for confidential clients; MCP connectors are expected to be public + PKCE. */
  clientSecretHash?: string;
}

export interface OAuthAuthorizationCode {
  code: string; // the lookup key (hashed at rest, see oauthRepo)
  clientId: string;
  redirectUri: string;
  codeChallenge: string;
  codeChallengeMethod: "S256";
  resource: string;
  scope: string;
  subject: string; // stable claim identifying the end user (owner)
  state: string;
  createdAt: string;
  expiresAt: string;
  consumed: boolean;
}

/**
 * Server-side state for an in-flight upstream Microsoft sign-in redirect
 * (owner login or connecting/reauthorizing a Microsoft account). Must be
 * persisted rather than kept in-memory: Container Apps consumption-plan
 * replicas scale to zero after 5 minutes idle, and a slow interactive
 * Microsoft sign-in (MFA, etc.) can outlast that — an in-memory store would
 * silently lose the handshake on a cold-started replica.
 */
export interface PendingUpstreamAuth {
  state: string; // lookup key
  kind: "owner_login" | "connect";
  codeVerifier: string;
  /** owner_login only: where to redirect after a successful login. */
  returnTo?: string;
  /** connect only: set when this is a reauth of an existing connection rather than a brand new one. */
  reauthConnectionId?: string;
  createdAt: string;
  expiresAt: string;
}

export interface OAuthRefreshToken {
  tokenId: string; // lookup key (hashed at rest)
  clientId: string;
  subject: string;
  resource: string;
  scope: string;
  createdAt: string;
  expiresAt: string;
  revoked: boolean;
  /** Rotation chain: previous token id, for replay detection. */
  rotatedFrom?: string;
  used: boolean;
}
