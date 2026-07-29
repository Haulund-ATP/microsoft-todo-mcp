import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { isPlausibleCodeChallenge, verifyPkceS256 } from "../../src/oauth/pkce.js";

function challengeFor(verifier: string): string {
  return createHash("sha256")
    .update(verifier, "ascii")
    .digest("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

describe("pkce", () => {
  it("verifies a matching verifier/challenge pair", () => {
    const verifier = "a".repeat(64);
    const challenge = challengeFor(verifier);
    expect(verifyPkceS256(verifier, challenge)).toBe(true);
  });

  it("rejects a mismatched verifier", () => {
    const verifier = "a".repeat(64);
    const challenge = challengeFor("b".repeat(64));
    expect(verifyPkceS256(verifier, challenge)).toBe(false);
  });

  it("rejects a challenge of a different length without throwing", () => {
    expect(verifyPkceS256("verifier", "short")).toBe(false);
  });

  describe("isPlausibleCodeChallenge", () => {
    it("accepts a well-formed 43-character challenge", () => {
      expect(isPlausibleCodeChallenge("a".repeat(43))).toBe(true);
    });

    it("rejects challenges that are too short", () => {
      expect(isPlausibleCodeChallenge("short")).toBe(false);
    });

    it("rejects challenges with invalid characters", () => {
      expect(isPlausibleCodeChallenge("!".repeat(43))).toBe(false);
    });
  });
});
