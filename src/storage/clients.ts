import { TableClient, TableServiceClient } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";
import { DefaultAzureCredential } from "@azure/identity";
import { loadEnv } from "../config/env.js";

/**
 * Storage account access. Prefers managed identity via
 * AZURE_STORAGE_ACCOUNT_NAME (production); falls back to a connection
 * string (Azurite / local dev) if that is what's configured.
 */

function accountUrl(kind: "table" | "blob"): string {
  const env = loadEnv();
  const name = env.AZURE_STORAGE_ACCOUNT_NAME;
  if (!name) throw new Error("AZURE_STORAGE_ACCOUNT_NAME is not set.");
  return `https://${name}.${kind}.core.windows.net`;
}

export function getBlobServiceClient(): BlobServiceClient {
  const env = loadEnv();
  if (env.AZURE_STORAGE_CONNECTION_STRING && env.NODE_ENV !== "production") {
    return BlobServiceClient.fromConnectionString(env.AZURE_STORAGE_CONNECTION_STRING);
  }
  return new BlobServiceClient(accountUrl("blob"), new DefaultAzureCredential());
}

export function getTableServiceClient(): TableServiceClient {
  const env = loadEnv();
  if (env.AZURE_STORAGE_CONNECTION_STRING && env.NODE_ENV !== "production") {
    return TableServiceClient.fromConnectionString(env.AZURE_STORAGE_CONNECTION_STRING);
  }
  return new TableServiceClient(accountUrl("table"), new DefaultAzureCredential());
}

export function getTableClient(tableName: string): TableClient {
  const env = loadEnv();
  if (env.AZURE_STORAGE_CONNECTION_STRING && env.NODE_ENV !== "production") {
    return TableClient.fromConnectionString(env.AZURE_STORAGE_CONNECTION_STRING, tableName);
  }
  return new TableClient(accountUrl("table"), tableName, new DefaultAzureCredential());
}

export const TABLE_NAMES = {
  connections: "connections",
  profiles: "profiles",
  oauthClients: "oauthclients",
  oauthCodes: "oauthcodes",
  oauthTokens: "oauthtokens",
} as const;

export const BLOB_CONTAINERS = {
  msalCache: "msal-token-caches",
} as const;

/** Ensures all required tables/containers exist. Call once at startup. */
export async function ensureStorageProvisioned(): Promise<void> {
  const tableService = getTableServiceClient();
  for (const table of Object.values(TABLE_NAMES)) {
    await tableService.createTable(table).catch(ignoreAlreadyExists);
  }
  const blobService = getBlobServiceClient();
  for (const container of Object.values(BLOB_CONTAINERS)) {
    await blobService
      .getContainerClient(container)
      .createIfNotExists({ access: undefined }); // private container, no public access
  }
}

function ignoreAlreadyExists(err: unknown): void {
  const code = (err as { statusCode?: number } | undefined)?.statusCode;
  if (code === 409) return;
  throw err;
}
