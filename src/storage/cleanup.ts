import { getTableClient, TABLE_NAMES } from "./clients.js";
import { logger } from "../logging/logger.js";
import { DeletedTasksRepo } from "./deletedTasksRepo.js";

/**
 * Deletes expired one-time OAuth rows. Nothing removed them before, so the
 * tables only grew (1,800+ refresh tokens, plus every abandoned sign-in and
 * unredeemed authorization code). Expired rows are dead: codes and pending
 * sign-ins are rejected after `expiresAt`, and an expired refresh token is
 * rejected before any replay check, so its row no longer protects anything.
 *
 * `expiresAt` is an ISO-8601 string, which Table Storage compares
 * lexicographically — correct for same-format UTC timestamps.
 */

const EXPIRING_TABLES = [TABLE_NAMES.oauthCodes, TABLE_NAMES.oauthTokens, TABLE_NAMES.pendingAuth] as const;

/** Keep rows a day past expiry so a just-expired code or token still gets a precise rejection reason. */
const RETENTION_MS = 24 * 60 * 60 * 1000;

export async function deleteExpiredOAuthRows(now = Date.now()): Promise<Record<string, number>> {
  const cutoff = new Date(now - RETENTION_MS).toISOString();
  const deleted: Record<string, number> = {};
  for (const tableName of EXPIRING_TABLES) {
    const table = getTableClient(tableName);
    let count = 0;
    for await (const entity of table.listEntities({
      queryOptions: { filter: `expiresAt lt '${cutoff}'`, select: ["PartitionKey", "RowKey"] },
    })) {
      await table.deleteEntity(String(entity.partitionKey), String(entity.rowKey)).catch((err) => {
        if ((err as { statusCode?: number }).statusCode !== 404) throw err;
      });
      count++;
    }
    deleted[tableName] = count;
  }
  return deleted;
}

const CLEANUP_INTERVAL_MS = 24 * 60 * 60 * 1000;

/**
 * Runs once at startup (the replica scales to zero, so most runs happen on
 * a cold start) and then daily while the process stays up. Never throws:
 * cleanup failing must not affect serving.
 */
export function scheduleOAuthCleanup(): void {
  const run = () =>
    Promise.all([
      deleteExpiredOAuthRows()
        .then((deleted) => logger.info({ deleted }, "expired oauth rows cleaned up"))
        .catch((err) => logger.warn({ err: err instanceof Error ? err.message : String(err) }, "oauth cleanup failed")),
      new DeletedTasksRepo()
        .deleteExpired()
        .then((deleted) => logger.info({ deleted }, "expired deleted-task snapshots cleaned up"))
        .catch((err) => logger.warn({ err: err instanceof Error ? err.message : String(err) }, "snapshot cleanup failed")),
    ]);
  void run();
  setInterval(run, CLEANUP_INTERVAL_MS).unref();
}
