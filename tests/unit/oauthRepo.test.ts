import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { hashSecret } from "../../src/storage/oauthRepo.js";

describe("hashSecret", () => {
  it("produces a stable sha256 hex digest", () => {
    const value = "some-authorization-code-or-refresh-token";
    const expected = createHash("sha256").update(value, "utf8").digest("hex");
    expect(hashSecret(value)).toBe(expected);
  });

  it("produces different hashes for different inputs", () => {
    expect(hashSecret("a")).not.toBe(hashSecret("b"));
  });

  it("never returns the original value", () => {
    const value = "super-secret-refresh-token";
    expect(hashSecret(value)).not.toContain(value);
  });
});
