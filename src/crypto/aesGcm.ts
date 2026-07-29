import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";

/**
 * AES-256-GCM envelope encryption for the MSAL token cache blob before it is
 * persisted to Azure Storage. The data-encryption key (DEK) itself lives in
 * Key Vault (see src/config/secrets.ts) and is never written to disk in
 * plaintext form outside of process memory.
 *
 * Wire format (all binary, base64-encoded for storage as text):
 *   version(1 byte) || iv(12 bytes) || authTag(16 bytes) || ciphertext
 */

const ALGORITHM = "aes-256-gcm";
const IV_LENGTH = 12; // 96-bit IV recommended for GCM
const AUTH_TAG_LENGTH = 16;
const FORMAT_VERSION = 1;

export interface EncryptedPayload {
  /** base64-encoded envelope: version || iv || authTag || ciphertext */
  envelope: string;
}

function assertKeyLength(key: Buffer): void {
  if (key.length !== 32) {
    throw new Error(`AES-256-GCM key must be 32 bytes, got ${key.length}.`);
  }
}

export function encryptString(plaintext: string, key: Buffer, aad?: Buffer): EncryptedPayload {
  assertKeyLength(key);
  const iv = randomBytes(IV_LENGTH);
  const cipher = createCipheriv(ALGORITHM, key, iv);
  if (aad) cipher.setAAD(aad);
  const ciphertext = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const authTag = cipher.getAuthTag();

  const envelope = Buffer.concat([
    Buffer.from([FORMAT_VERSION]),
    iv,
    authTag,
    ciphertext,
  ]).toString("base64");

  return { envelope };
}

export function decryptString(payload: EncryptedPayload, key: Buffer, aad?: Buffer): string {
  assertKeyLength(key);
  const raw = Buffer.from(payload.envelope, "base64");
  const version = raw.readUInt8(0);
  if (version !== FORMAT_VERSION) {
    throw new Error(`Unsupported encryption envelope version: ${version}`);
  }
  const iv = raw.subarray(1, 1 + IV_LENGTH);
  const authTag = raw.subarray(1 + IV_LENGTH, 1 + IV_LENGTH + AUTH_TAG_LENGTH);
  const ciphertext = raw.subarray(1 + IV_LENGTH + AUTH_TAG_LENGTH);

  const decipher = createDecipheriv(ALGORITHM, key, iv);
  if (aad) decipher.setAAD(aad);
  decipher.setAuthTag(authTag);
  const plaintext = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
  return plaintext.toString("utf8");
}

/** Generates a fresh 256-bit key, e.g. for initial provisioning into Key Vault. */
export function generateDataEncryptionKey(): Buffer {
  return randomBytes(32);
}
