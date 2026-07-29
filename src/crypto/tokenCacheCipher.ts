import { getSecretStore } from "../config/secrets.js";
import { decryptString, encryptString, type EncryptedPayload } from "./aesGcm.js";

/**
 * Encrypts/decrypts the serialized MSAL token cache using the
 * data-encryption key stored in Key Vault under the secret named by
 * `MSAL_CACHE_DEK_SECRET_NAME`. The connection id is bound in as additional
 * authenticated data (AAD) so a ciphertext blob cannot be silently
 * re-associated with a different connection record.
 */

const MSAL_CACHE_DEK_SECRET_NAME = "msal-cache-encryption-key";

async function loadKey(): Promise<Buffer> {
  const store = getSecretStore();
  const b64 = await store.getSecret(MSAL_CACHE_DEK_SECRET_NAME);
  const key = Buffer.from(b64, "base64");
  if (key.length !== 32) {
    throw new Error(
      `Key Vault secret "${MSAL_CACHE_DEK_SECRET_NAME}" must decode to 32 bytes (base64 of a 256-bit key).`
    );
  }
  return key;
}

export async function encryptTokenCache(
  serializedCache: string,
  connectionId: string
): Promise<EncryptedPayload> {
  const key = await loadKey();
  return encryptString(serializedCache, key, Buffer.from(connectionId, "utf8"));
}

export async function decryptTokenCache(
  payload: EncryptedPayload,
  connectionId: string
): Promise<string> {
  const key = await loadKey();
  return decryptString(payload, key, Buffer.from(connectionId, "utf8"));
}
