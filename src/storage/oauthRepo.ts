import { createHash } from "node:crypto";
import { getTableClient, TABLE_NAMES } from "./clients.js";
import type { OAuthAuthorizationCode, OAuthClient, OAuthRefreshToken, PendingUpstreamAuth } from "./types.js";

/**
 * Authorization codes and refresh tokens are looked up by the SHA-256 hash
 * of the actual secret value, never by the raw value itself — mirrors how
 * password/token storage should work, and means a Table Storage export or
 * backup does not expose usable bearer secrets.
 */
export function hashSecret(value: string): string {
  return createHash("sha256").update(value, "utf8").digest("hex");
}

const CLIENT_PARTITION = "client";
const CODE_PARTITION = "code";
const TOKEN_PARTITION = "token";
const PENDING_AUTH_PARTITION = "pending";

function stripMeta<T extends object>(entity: T): T {
  const { partitionKey: _pk, rowKey: _rk, etag: _etag, timestamp: _ts, ...rest } = entity as T & {
    partitionKey?: string;
    rowKey?: string;
    etag?: string;
    timestamp?: string;
  };
  return rest as T;
}

export class OAuthClientsRepo {
  private table() {
    return getTableClient(TABLE_NAMES.oauthClients);
  }

  async get(clientId: string): Promise<OAuthClient | undefined> {
    try {
      const entity = await this.table().getEntity(CLIENT_PARTITION, clientId);
      return stripMeta(entity) as unknown as OAuthClient;
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
  }

  async register(client: OAuthClient): Promise<void> {
    await this.table().upsertEntity(
      { partitionKey: CLIENT_PARTITION, rowKey: client.clientId, ...client, redirectUris: JSON.stringify(client.redirectUris) },
      "Replace"
    );
  }

  async listRedirectUris(clientId: string): Promise<string[]> {
    const client = await this.get(clientId);
    if (!client) return [];
    const raw = (client as unknown as { redirectUris: unknown }).redirectUris;
    return typeof raw === "string" ? (JSON.parse(raw) as string[]) : (raw as string[]);
  }
}

export class OAuthCodesRepo {
  private table() {
    return getTableClient(TABLE_NAMES.oauthCodes);
  }

  /** Stores a code; `plainCode` is hashed before persisting. */
  async store(plainCode: string, data: Omit<OAuthAuthorizationCode, "code">): Promise<void> {
    await this.table().createEntity({
      partitionKey: CODE_PARTITION,
      rowKey: hashSecret(plainCode),
      ...data,
    });
  }

  async consume(
    plainCode: string
  ): Promise<{ record: OAuthAuthorizationCode } | { reason: "unknown" | "consumed" | "expired" }> {
    const rowKey = hashSecret(plainCode);
    let entity;
    try {
      entity = await this.table().getEntity(CODE_PARTITION, rowKey);
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return { reason: "unknown" };
      throw err;
    }
    const record = stripMeta(entity) as unknown as OAuthAuthorizationCode;
    if (record.consumed) return { reason: "consumed" }; // replay attempt
    if (new Date(record.expiresAt).getTime() < Date.now()) return { reason: "expired" };

    // Mark consumed (one-time use) rather than deleting immediately, so a
    // rapid replay is detected as "already consumed" rather than "unknown".
    // Conditional on the etag we read, so two concurrent exchanges of the
    // same code can't both pass the check above.
    if (!(await conditionalMerge(this.table(), { partitionKey: CODE_PARTITION, rowKey, consumed: true }, entity.etag))) {
      return { reason: "consumed" };
    }
    return { record: { ...record, code: plainCode } };
  }
}

/** Merge-updates an entity only if it still has `etag`; false if someone else changed it first. */
async function conditionalMerge(
  table: ReturnType<typeof getTableClient>,
  entity: { partitionKey: string; rowKey: string; [key: string]: unknown },
  etag: string | undefined
): Promise<boolean> {
  try {
    await table.updateEntity(entity, "Merge", { etag: etag ?? "*" });
    return true;
  } catch (err) {
    if ((err as { statusCode?: number }).statusCode === 412) return false;
    throw err;
  }
}

export class OAuthTokensRepo {
  private table() {
    return getTableClient(TABLE_NAMES.oauthTokens);
  }

  async store(plainToken: string, data: Omit<OAuthRefreshToken, "tokenId">): Promise<void> {
    await this.table().createEntity({
      partitionKey: TOKEN_PARTITION,
      rowKey: hashSecret(plainToken),
      ...data,
    });
  }

  async get(plainToken: string): Promise<(OAuthRefreshToken & { etag?: string }) | undefined> {
    try {
      const entity = await this.table().getEntity(TOKEN_PARTITION, hashSecret(plainToken));
      return { ...(stripMeta(entity) as unknown as OAuthRefreshToken), tokenId: hashSecret(plainToken), etag: entity.etag };
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
  }

  async revoke(plainToken: string): Promise<void> {
    const rowKey = hashSecret(plainToken);
    await this.table()
      .updateEntity({ partitionKey: TOKEN_PARTITION, rowKey, revoked: true }, "Merge")
      .catch((err) => {
        if ((err as { statusCode?: number }).statusCode !== 404) throw err;
      });
  }

  /**
   * Marks the token used, but only if it is unchanged since it was read
   * (`etag`). Returns false when a concurrent request redeemed it first.
   */
  async claim(plainToken: string, etag: string | undefined, usedAt: string): Promise<boolean> {
    return conditionalMerge(
      this.table(),
      { partitionKey: TOKEN_PARTITION, rowKey: hashSecret(plainToken), used: true, usedAt },
      etag
    );
  }

  /** Revokes every refresh token descending from a compromised token (replay response). */
  async revokeChainFrom(tokenId: string): Promise<void> {
    for await (const entity of this.table().listEntities({
      queryOptions: { filter: `PartitionKey eq '${TOKEN_PARTITION}' and rotatedFrom eq '${tokenId}'` },
    })) {
      const rowKey = String(entity.rowKey);
      await this.table().updateEntity(
        { partitionKey: TOKEN_PARTITION, rowKey, revoked: true },
        "Merge"
      );
      await this.revokeChainFrom(rowKey);
    }
  }
}

/**
 * Persists the PKCE verifier + redirect state for an in-flight upstream
 * Microsoft sign-in (see PendingUpstreamAuth for why this must not be an
 * in-memory Map). One-time use: consume() deletes the row.
 */
export class PendingUpstreamAuthRepo {
  private table() {
    return getTableClient(TABLE_NAMES.pendingAuth);
  }

  async store(entry: PendingUpstreamAuth): Promise<void> {
    await this.table().upsertEntity(
      { partitionKey: PENDING_AUTH_PARTITION, rowKey: entry.state, ...entry },
      "Replace"
    );
  }

  async consume(state: string): Promise<PendingUpstreamAuth | undefined> {
    let entity;
    try {
      entity = await this.table().getEntity(PENDING_AUTH_PARTITION, state);
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
    // Delete first (one-time use) regardless of expiry, so a replay of an
    // expired-but-not-yet-cleaned-up state can't be consumed twice either.
    await this.table().deleteEntity(PENDING_AUTH_PARTITION, state).catch(() => undefined);
    const record = stripMeta(entity) as unknown as PendingUpstreamAuth;
    if (new Date(record.expiresAt).getTime() < Date.now()) return undefined;
    return record;
  }
}
