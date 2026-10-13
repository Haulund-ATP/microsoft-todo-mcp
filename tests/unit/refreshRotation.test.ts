import { describe, expect, it } from "vitest";
import { rotateRefreshToken, type RefreshTokenStore } from "../../src/oauth/refreshRotation.js";
import type { OAuthRefreshToken } from "../../src/storage/types.js";

/** In-memory store with Table Storage's etag semantics (conditional merge fails on a stale etag). */
function memoryStore() {
  const rows = new Map<string, Omit<OAuthRefreshToken, "tokenId"> & { etag: number }>();
  let etagCounter = 0;
  const store: RefreshTokenStore & { rows: typeof rows; seed(token: string, data?: Partial<OAuthRefreshToken>): void } = {
    rows,
    seed(token, data = {}) {
      rows.set(token, {
        clientId: "client", subject: "owner", resource: "https://todo.example", scope: "mcp",
        createdAt: "2026-10-13T05:00:00.000Z", expiresAt: "2026-11-13T05:00:00.000Z",
        revoked: false, used: false, ...data, etag: ++etagCounter,
      });
    },
    async get(token) {
      const row = rows.get(token);
      if (!row) return undefined;
      const { etag, ...rest } = row;
      return { ...rest, tokenId: `id:${token}`, etag: String(etag) };
    },
    async claim(token, etag, usedAt) {
      // Yield first so two concurrent callers both read before either writes.
      await new Promise((resolve) => setTimeout(resolve, 0));
      const row = rows.get(token)!;
      if (etag !== String(row.etag)) return false;
      rows.set(token, { ...row, used: true, usedAt, etag: ++etagCounter });
      return true;
    },
    async store(token, data) {
      rows.set(token, { ...data, etag: ++etagCounter });
    },
    async revoke(token) {
      const row = rows.get(token);
      if (row) rows.set(token, { ...row, revoked: true, etag: ++etagCounter });
    },
    async revokeChainFrom(tokenId) {
      for (const [token, row] of rows) {
        if (row.rotatedFrom === tokenId) {
          rows.set(token, { ...row, revoked: true, etag: ++etagCounter });
          await store.revokeChainFrom(`id:${token}`);
        }
      }
    },
  };
  return store;
}

const NOW = Date.parse("2026-10-13T06:00:00.000Z");
let counter = 0;
const options = (now = NOW) => ({ now, graceSeconds: 60, ttlSeconds: 3600, newToken: () => `new-${++counter}` });

describe("rotateRefreshToken", () => {
  it("rotates an unused token, marks it used and links the new token to it", async () => {
    const store = memoryStore();
    store.seed("rt");
    const outcome = await rotateRefreshToken(store, "rt", "client", options());
    expect(outcome).toMatchObject({ ok: true, mode: "rotated" });
    expect(store.rows.get("rt")).toMatchObject({ used: true, usedAt: new Date(NOW).toISOString() });
    if (outcome.ok) expect(store.rows.get(outcome.refreshToken)).toMatchObject({ used: false, rotatedFrom: "id:rt" });
  });

  it("survives two parallel refreshes of the same token (the 13 Oct reconnect)", async () => {
    const store = memoryStore();
    store.seed("rt");
    const [a, b] = await Promise.all([
      rotateRefreshToken(store, "rt", "client", options()),
      rotateRefreshToken(store, "rt", "client", options()),
    ]);
    expect([a.ok, b.ok]).toEqual([true, true]);
    const modes = [a, b].map((o) => (o.ok ? o.mode : o.category)).sort();
    expect(modes).toEqual(["grace", "rotated"]);
    expect(store.rows.get("rt")!.revoked).toBe(false);
  });

  it("rotates again when a used token is replayed inside the grace window", async () => {
    const store = memoryStore();
    store.seed("rt", { used: true, usedAt: new Date(NOW - 30_000).toISOString() });
    const outcome = await rotateRefreshToken(store, "rt", "client", options());
    expect(outcome).toMatchObject({ ok: true, mode: "grace" });
  });

  it("revokes the whole chain when a used token is replayed after the grace window", async () => {
    const store = memoryStore();
    store.seed("rt", { used: true, usedAt: new Date(NOW - 61_000).toISOString() });
    store.seed("child", { rotatedFrom: "id:rt" });
    store.seed("grandchild", { rotatedFrom: "id:child" });
    const outcome = await rotateRefreshToken(store, "rt", "client", options());
    expect(outcome).toEqual({ ok: false, category: "reuse" });
    expect(["rt", "child", "grandchild"].map((t) => store.rows.get(t)!.revoked)).toEqual([true, true, true]);
  });

  it("treats a legacy used token without usedAt as a replay", async () => {
    const store = memoryStore();
    store.seed("rt", { used: true });
    expect(await rotateRefreshToken(store, "rt", "client", options())).toEqual({ ok: false, category: "reuse" });
  });

  it.each([
    [{ revoked: true }, "client", "revoked"],
    [{ expiresAt: new Date(NOW - 1).toISOString() }, "client", "expired"],
    [{}, "other-client", "client_mismatch"],
  ] as const)("rejects %o for %s as %s without issuing a token", async (data, clientId, category) => {
    const store = memoryStore();
    store.seed("rt", data);
    const before = store.rows.size;
    expect(await rotateRefreshToken(store, "rt", clientId, options())).toEqual({ ok: false, category });
    expect(store.rows.size).toBe(before);
  });

  it("rejects unknown tokens", async () => {
    expect(await rotateRefreshToken(memoryStore(), "nope", "client", options())).toEqual({ ok: false, category: "unknown" });
  });
});
