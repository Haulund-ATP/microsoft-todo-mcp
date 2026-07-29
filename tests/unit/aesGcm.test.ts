import { describe, expect, it } from "vitest";
import { decryptString, encryptString, generateDataEncryptionKey } from "../../src/crypto/aesGcm.js";

describe("aesGcm", () => {
  it("round-trips plaintext through encrypt/decrypt", () => {
    const key = generateDataEncryptionKey();
    const plaintext = JSON.stringify({ hello: "world", n: 42 });
    const { envelope } = encryptString(plaintext, key);
    const decrypted = decryptString({ envelope }, key);
    expect(decrypted).toBe(plaintext);
  });

  it("produces a different ciphertext each time (random IV)", () => {
    const key = generateDataEncryptionKey();
    const a = encryptString("same input", key);
    const b = encryptString("same input", key);
    expect(a.envelope).not.toBe(b.envelope);
  });

  it("fails to decrypt with the wrong key", () => {
    const key = generateDataEncryptionKey();
    const otherKey = generateDataEncryptionKey();
    const { envelope } = encryptString("secret", key);
    expect(() => decryptString({ envelope }, otherKey)).toThrow();
  });

  it("fails to decrypt when AAD does not match (connection-id binding)", () => {
    const key = generateDataEncryptionKey();
    const { envelope } = encryptString("secret", key, Buffer.from("connection-a"));
    expect(() => decryptString({ envelope }, key, Buffer.from("connection-b"))).toThrow();
  });

  it("rejects a key that is not 32 bytes", () => {
    const shortKey = Buffer.alloc(16);
    expect(() => encryptString("x", shortKey)).toThrow(/32 bytes/);
  });

  it("rejects a tampered envelope (auth tag check)", () => {
    const key = generateDataEncryptionKey();
    const { envelope } = encryptString("secret", key);
    const raw = Buffer.from(envelope, "base64");
    raw[raw.length - 1] ^= 0xff; // flip last ciphertext byte
    const tampered = raw.toString("base64");
    expect(() => decryptString({ envelope: tampered }, key)).toThrow();
  });
});
