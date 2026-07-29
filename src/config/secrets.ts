import { DefaultAzureCredential } from "@azure/identity";
import { SecretClient } from "@azure/keyvault-secrets";
import { loadEnv } from "./env.js";

/**
 * Thin wrapper around Key Vault secret retrieval using the runtime's
 * managed identity (falls back through DefaultAzureCredential's chain,
 * which also supports `az login` locally for development).
 *
 * Secrets are cached in-memory for the process lifetime; there is no
 * background refresh because Container Apps consumption-plan instances are
 * short-lived (scale-to-zero) and a redeploy/cold-start naturally picks up
 * rotated secrets.
 */
export class SecretStore {
  private readonly client: SecretClient;
  private readonly cache = new Map<string, string>();

  constructor(vaultUri?: string) {
    const env = loadEnv();
    const uri = vaultUri ?? env.AZURE_KEY_VAULT_URI;
    const credential = new DefaultAzureCredential(
      env.AZURE_CLIENT_ID_MANAGED_IDENTITY
        ? { managedIdentityClientId: env.AZURE_CLIENT_ID_MANAGED_IDENTITY }
        : undefined
    );
    this.client = new SecretClient(uri, credential);
  }

  async getSecret(name: string): Promise<string> {
    const cached = this.cache.get(name);
    if (cached !== undefined) return cached;
    const secret = await this.client.getSecret(name);
    if (!secret.value) {
      throw new Error(`Key Vault secret "${name}" has no value.`);
    }
    this.cache.set(name, secret.value);
    return secret.value;
  }

  /** Explicit invalidation, e.g. after a manual key-rotation admin action. */
  invalidate(name: string): void {
    this.cache.delete(name);
  }
}

let singleton: SecretStore | undefined;

export function getSecretStore(): SecretStore {
  if (!singleton) singleton = new SecretStore();
  return singleton;
}
