import { randomUUID } from "node:crypto";
import { getBlobServiceClient, getTableClient, BLOB_CONTAINERS, TABLE_NAMES } from "./clients.js";
import type { Connection } from "./types.js";
import { encryptTokenCache, decryptTokenCache } from "../crypto/tokenCacheCipher.js";

/**
 * Connections partition key is fixed ("connection") since the total number
 * of connections per deployment is small (single-owner, a handful of
 * Microsoft accounts); alias uniqueness is enforced at the application
 * layer via `findByAlias`.
 */
const PARTITION_KEY = "connection";

interface ConnectionEntity extends Connection {
  partitionKey: string;
  rowKey: string;
}

function toEntity(c: Connection): ConnectionEntity {
  return { partitionKey: PARTITION_KEY, rowKey: c.connectionId, ...c };
}

function fromEntity(e: ConnectionEntity): Connection {
  const { partitionKey: _pk, rowKey: _rk, etag: _etag, timestamp: _ts, ...rest } = e as ConnectionEntity & {
    etag?: string;
    timestamp?: string;
  };
  return rest as Connection;
}

export class ConnectionsRepo {
  private table() {
    return getTableClient(TABLE_NAMES.connections);
  }

  async list(): Promise<Connection[]> {
    const out: Connection[] = [];
    for await (const entity of this.table().listEntities<ConnectionEntity>({
      queryOptions: { filter: `PartitionKey eq '${PARTITION_KEY}'` },
    })) {
      out.push(fromEntity(entity));
    }
    return out;
  }

  async get(connectionId: string): Promise<Connection | undefined> {
    try {
      const entity = await this.table().getEntity<ConnectionEntity>(PARTITION_KEY, connectionId);
      return fromEntity(entity);
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
  }

  async findByAlias(alias: string): Promise<Connection | undefined> {
    const all = await this.list();
    return all.find((c) => c.alias === alias);
  }

  async create(input: Omit<Connection, "connectionId" | "createdAt" | "tokenCacheBlobName">): Promise<Connection> {
    const connectionId = randomUUID();
    const connection: Connection = {
      ...input,
      connectionId,
      createdAt: new Date().toISOString(),
      tokenCacheBlobName: `${connectionId}.cache`,
    };
    await this.table().createEntity(toEntity(connection));
    return connection;
  }

  async update(connectionId: string, patch: Partial<Connection>): Promise<void> {
    await this.table().updateEntity(
      { partitionKey: PARTITION_KEY, rowKey: connectionId, ...patch },
      "Merge"
    );
  }

  async delete(connectionId: string): Promise<void> {
    const connection = await this.get(connectionId);
    await this.table().deleteEntity(PARTITION_KEY, connectionId);
    if (connection) {
      const container = getBlobServiceClient().getContainerClient(BLOB_CONTAINERS.msalCache);
      await container.getBlockBlobClient(connection.tokenCacheBlobName).deleteIfExists();
    }
  }

  async saveTokenCache(connectionId: string, serializedCache: string): Promise<void> {
    const connection = await this.get(connectionId);
    if (!connection) throw new Error(`Unknown connection: ${connectionId}`);
    const { envelope } = await encryptTokenCache(serializedCache, connectionId);
    const container = getBlobServiceClient().getContainerClient(BLOB_CONTAINERS.msalCache);
    // BlockBlobClient.upload always overwrites the blob content by default.
    await container
      .getBlockBlobClient(connection.tokenCacheBlobName)
      .upload(envelope, Buffer.byteLength(envelope, "utf8"));
  }

  async loadTokenCache(connectionId: string): Promise<string | undefined> {
    const connection = await this.get(connectionId);
    if (!connection) return undefined;
    const container = getBlobServiceClient().getContainerClient(BLOB_CONTAINERS.msalCache);
    const blob = container.getBlockBlobClient(connection.tokenCacheBlobName);
    if (!(await blob.exists())) return undefined;
    const download = await blob.downloadToBuffer();
    const envelope = download.toString("utf8");
    return decryptTokenCache({ envelope }, connectionId);
  }
}
