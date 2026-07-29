import { createHash } from "node:crypto";
import { getTableClient, TABLE_NAMES } from "./clients.js";
import type { OAuthAuthorizationCode, OAuthClient, OAuthRefreshToken } from "./types.js";

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

  async consume(plainCode: string): Promise<OAuthAuthorizationCode | undefined> {
    const rowKey = hashSecret(plainCode);
    let entity;
    try {
      entity = await this.table().getEntity(CODE_PARTITION, rowKey);
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
    const record = stripMeta(entity) as unknown as OAuthAuthorizationCode;
    if (record.consumed) return undefined; // replay attempt
    if (new Date(record.expiresAt).getTime() < Date.now()) return undefined;

    // Mark consumed (one-time use) rather than deleting immediately, so a
    // rapid replay is detected as "already consumed" rather than "unknown".
    await this.table().updateEntity(
      { partitionKey: CODE_PARTITION, rowKey, consumed: true },
      "Merge"
    );
    return { ...record, code: plainCode };
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

  async get(plainToken: string): Promise<OAuthRefreshToken | undefined> {
    try {
      const entity = await this.table().getEntity(TOKEN_PARTITION, hashSecret(plainToken));
      return { ...(stripMeta(entity) as unknown as OAuthRefreshToken), tokenId: hashSecret(plainToken) };
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

  async markUsed(plainToken: string): Promise<void> {
    const rowKey = hashSecret(plainToken);
    await this.table().updateEntity(
      { partitionKey: TOKEN_PARTITION, rowKey, used: true },
      "Merge"
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
