import type { OAuthRefreshToken } from "../storage/types.js";

/**
 * Refresh-token rotation with a short reuse grace window.
 *
 * Every refresh issues a new refresh token and marks the presented one used.
 * Presenting a used token again normally means it was stolen and replayed,
 * so the whole chain is revoked. But ChatGPT refreshes once per tool call and
 * runs tool calls in parallel, so two requests routinely redeem the same
 * token within the same second. Treating that as theft revoked the user's
 * connection and forced a reconnect. A token redeemed again within
 * `graceSeconds` of its first use is therefore rotated once more instead
 * (the same "reuse interval" Auth0 and Okta use); outside the window the
 * replay response is unchanged.
 */

export interface RefreshTokenStore {
  get(plainToken: string): Promise<(OAuthRefreshToken & { etag?: string }) | undefined>;
  claim(plainToken: string, etag: string | undefined, usedAt: string): Promise<boolean>;
  store(plainToken: string, data: Omit<OAuthRefreshToken, "tokenId">): Promise<void>;
  revoke(plainToken: string): Promise<void>;
  revokeChainFrom(tokenId: string): Promise<void>;
}

export type RefreshFailure = "unknown" | "client_mismatch" | "revoked" | "expired" | "reuse";

export type RefreshOutcome =
  | { ok: true; mode: "rotated" | "grace"; refreshToken: string; record: OAuthRefreshToken }
  | { ok: false; category: RefreshFailure };

export interface RotateOptions {
  now: number;
  graceSeconds: number;
  ttlSeconds: number;
  newToken: () => string;
}

export async function rotateRefreshToken(
  store: RefreshTokenStore,
  plainToken: string,
  clientId: string,
  options: RotateOptions
): Promise<RefreshOutcome> {
  const { now } = options;
  const existing = await store.get(plainToken);
  if (!existing) return { ok: false, category: "unknown" };
  if (existing.clientId !== clientId) return { ok: false, category: "client_mismatch" };
  if (existing.revoked) return { ok: false, category: "revoked" };
  if (new Date(existing.expiresAt).getTime() < now) return { ok: false, category: "expired" };

  let mode: "rotated" | "grace" = "rotated";
  const claimed = !existing.used && (await store.claim(plainToken, existing.etag, new Date(now).toISOString()));
  if (!claimed) {
    let usedAt = existing.usedAt;
    if (!existing.used) {
      // Lost the race to a concurrent request: re-read to learn when it won.
      const fresh = await store.get(plainToken);
      if (fresh?.revoked) return { ok: false, category: "revoked" };
      usedAt = fresh?.usedAt;
    }
    const withinGrace = usedAt !== undefined && now - new Date(usedAt).getTime() <= options.graceSeconds * 1000;
    if (!withinGrace) {
      await store.revoke(plainToken);
      await store.revokeChainFrom(existing.tokenId);
      return { ok: false, category: "reuse" };
    }
    mode = "grace";
  }

  const refreshToken = options.newToken();
  const record: OAuthRefreshToken = {
    tokenId: "",
    clientId: existing.clientId,
    subject: existing.subject,
    resource: existing.resource,
    scope: existing.scope,
    createdAt: new Date(now).toISOString(),
    expiresAt: new Date(now + options.ttlSeconds * 1000).toISOString(),
    revoked: false,
    used: false,
    // Grace-issued siblings share the parent, so a later out-of-window
    // replay of that parent still revokes every descendant.
    rotatedFrom: existing.tokenId,
  };
  const { tokenId: _tokenId, ...data } = record;
  await store.store(refreshToken, data);
  return { ok: true, mode, refreshToken, record };
}
