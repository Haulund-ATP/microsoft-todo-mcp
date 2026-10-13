import { randomBytes } from "node:crypto";
import { getBlobServiceClient, BLOB_CONTAINERS } from "./clients.js";
import { decryptDeletedTask, encryptDeletedTask } from "../crypto/tokenCacheCipher.js";
import type { ChecklistItem } from "../graph/todoApi.js";

/**
 * Undo snapshots of deleted tasks: the task exactly as Graph returned it,
 * plus its checklist, taken right before the delete. Microsoft To Do keeps
 * no reachable history — Graph can't see deleted tasks — so after an
 * over-eager chat client deleted an open task with a purchase date
 * (2026-10-13) its date was simply gone.
 *
 * One encrypted blob per snapshot under "<connectionId>/", named
 * "<deletedAt>-<random>" so a prefix listing is chronological and expiry
 * needs no extra index. Kept RETENTION_DAYS, then deleted by the daily
 * cleanup.
 */

export const RETENTION_DAYS = 30;

export interface DeletedTaskSnapshot {
  snapshotId: string;
  connectionId: string;
  listId: string;
  listName?: string;
  taskId: string;
  deletedAt: string;
  deletedBy: string;
  task: Record<string, unknown>;
  checklist: ChecklistItem[];
  restoredAt?: string;
  restoredTaskId?: string;
}

const compactTime = (iso: string) => iso.replace(/[-:.]/g, "");

/** Parses the deletedAt encoded in a snapshot id ("20261013T102521123Z-ab12cd"). */
export function snapshotTime(snapshotId: string): number {
  const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(\d{3})Z-/.exec(snapshotId);
  return m ? Date.UTC(+m[1]!, +m[2]! - 1, +m[3]!, +m[4]!, +m[5]!, +m[6]!, +m[7]!) : NaN;
}

const isSafeSegment = (value: string) => /^[A-Za-z0-9-]+$/.test(value);

export class DeletedTasksRepo {
  private container() {
    return getBlobServiceClient().getContainerClient(BLOB_CONTAINERS.deletedTasks);
  }

  async save(snapshot: Omit<DeletedTaskSnapshot, "snapshotId">): Promise<string> {
    const snapshotId = `${compactTime(snapshot.deletedAt)}-${randomBytes(4).toString("hex")}`;
    const json = JSON.stringify({ ...snapshot, snapshotId });
    const { envelope } = await encryptDeletedTask(json, snapshot.connectionId);
    await this.container()
      .getBlockBlobClient(`${snapshot.connectionId}/${snapshotId}`)
      .upload(envelope, Buffer.byteLength(envelope, "utf8"));
    return snapshotId;
  }

  async get(connectionId: string, snapshotId: string): Promise<DeletedTaskSnapshot | undefined> {
    if (!isSafeSegment(snapshotId)) return undefined;
    const blob = this.container().getBlockBlobClient(`${connectionId}/${snapshotId}`);
    if (!(await blob.exists())) return undefined;
    const envelope = (await blob.downloadToBuffer()).toString("utf8");
    return JSON.parse(await decryptDeletedTask({ envelope }, connectionId)) as DeletedTaskSnapshot;
  }

  /** Newest first. */
  async list(connectionId: string, limit: number): Promise<DeletedTaskSnapshot[]> {
    const names: string[] = [];
    for await (const blob of this.container().listBlobsFlat({ prefix: `${connectionId}/` })) names.push(blob.name);
    const newest = names.sort().reverse().slice(0, limit);
    const snapshots = await Promise.all(newest.map((name) => this.get(connectionId, name.slice(connectionId.length + 1))));
    return snapshots.filter((s): s is DeletedTaskSnapshot => s !== undefined);
  }

  async markRestored(snapshot: DeletedTaskSnapshot, restoredTaskId: string, restoredAt: string): Promise<void> {
    const { snapshotId, ...rest } = snapshot;
    const json = JSON.stringify({ ...rest, snapshotId, restoredAt, restoredTaskId });
    const { envelope } = await encryptDeletedTask(json, snapshot.connectionId);
    await this.container()
      .getBlockBlobClient(`${snapshot.connectionId}/${snapshotId}`)
      .upload(envelope, Buffer.byteLength(envelope, "utf8"));
  }

  async deleteExpired(now = Date.now()): Promise<number> {
    const cutoff = now - RETENTION_DAYS * 24 * 60 * 60 * 1000;
    let deleted = 0;
    for await (const blob of this.container().listBlobsFlat()) {
      const id = blob.name.split("/")[1] ?? "";
      if (snapshotTime(id) < cutoff) {
        await this.container().getBlockBlobClient(blob.name).deleteIfExists();
        deleted++;
      }
    }
    return deleted;
  }
}
