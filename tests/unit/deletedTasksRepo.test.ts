import { describe, expect, it, vi } from "vitest";

const blobs = new Map<string, string>();

vi.mock("../../src/storage/clients.js", () => ({
  BLOB_CONTAINERS: { deletedTasks: "deleted-tasks" },
  getBlobServiceClient: () => ({
    getContainerClient: () => ({
      getBlockBlobClient: (name: string) => ({
        upload: async (data: string) => void blobs.set(name, data),
        exists: async () => blobs.has(name),
        downloadToBuffer: async () => Buffer.from(blobs.get(name)!, "utf8"),
        deleteIfExists: async () => void blobs.delete(name),
      }),
      async *listBlobsFlat(opts?: { prefix?: string }) {
        for (const name of [...blobs.keys()]) if (!opts?.prefix || name.startsWith(opts.prefix)) yield { name };
      },
    }),
  }),
}));

// Reversible stand-in for the Key Vault-backed cipher, bound to the connection like the real AAD.
vi.mock("../../src/crypto/tokenCacheCipher.js", () => ({
  encryptDeletedTask: async (json: string, conn: string) => ({ envelope: Buffer.from(`${conn}|${json}`).toString("base64") }),
  decryptDeletedTask: async ({ envelope }: { envelope: string }, conn: string) => {
    const text = Buffer.from(envelope, "base64").toString("utf8");
    if (!text.startsWith(`${conn}|`)) throw new Error("aad mismatch");
    return text.slice(conn.length + 1);
  },
}));

const { DeletedTasksRepo, snapshotTime } = await import("../../src/storage/deletedTasksRepo.js");

const base = (deletedAt: string, title: string) => ({
  connectionId: "conn-a", listId: "L", taskId: title, deletedAt, deletedBy: "delete_task",
  task: { title, dueDateTime: { dateTime: "2026-10-20T00:00:00" } }, checklist: [],
});

describe("DeletedTasksRepo", () => {
  it("stores encrypted snapshots per connection, lists newest first, and marks restores", async () => {
    blobs.clear();
    const repo = new DeletedTasksRepo();
    const older = await repo.save(base("2026-10-13T08:25:00.000Z", "Older"));
    const newer = await repo.save(base("2026-10-13T08:30:00.000Z", "Køb B vitamin"));
    await repo.save({ ...base("2026-10-13T08:31:00.000Z", "Other account"), connectionId: "conn-b" });

    expect([...blobs.values()].join()).not.toContain("vitamin");
    expect(snapshotTime(newer)).toBe(Date.parse("2026-10-13T08:30:00.000Z"));

    const listed = await repo.list("conn-a", 10);
    expect(listed.map((s) => s.task.title)).toEqual(["Køb B vitamin", "Older"]);
    expect(listed[0]).toMatchObject({ snapshotId: newer, task: { dueDateTime: { dateTime: "2026-10-20T00:00:00" } } });

    await repo.markRestored(listed[1]!, "new-id", "2026-10-13T09:00:00.000Z");
    expect(await repo.get("conn-a", older)).toMatchObject({ restoredTaskId: "new-id", restoredAt: "2026-10-13T09:00:00.000Z" });
  });

  it("rejects snapshot ids that try to leave the connection's folder", async () => {
    expect(await new DeletedTasksRepo().get("conn-a", "../conn-b/x")).toBeUndefined();
  });

  it("deletes snapshots older than 30 days only", async () => {
    blobs.clear();
    const repo = new DeletedTasksRepo();
    await repo.save(base("2026-09-01T00:00:00.000Z", "Old"));
    const keep = await repo.save(base("2026-10-01T00:00:00.000Z", "Recent"));
    expect(await repo.deleteExpired(Date.parse("2026-10-13T00:00:00.000Z"))).toBe(1);
    expect([...blobs.keys()]).toEqual([`conn-a/${keep}`]);
  });
});
