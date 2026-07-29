import {
  ConfidentialClientApplication,
  type AuthenticationResult,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { loadEnv } from "../config/env.js";
import { getSecretStore } from "../config/secrets.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";

/**
 * Builds a per-connection MSAL ConfidentialClientApplication whose token
 * cache is persisted (encrypted) to Azure Storage via a cache plugin. Using
 * a confidential client (client secret held in Key Vault) lets the same
 * multi-tenant + personal-account app registration serve both account
 * types without the app needing to be "public client" (no PKCE needed on
 * this leg; PKCE is still used for the *user-facing* browser redirect).
 *
 * Authority: /common — required for sign-in audience
 * "AzureADandPersonalMicrosoftAccount" (work/school + personal MSA).
 */

const AUTHORITY = "https://login.microsoftonline.com/common";
export const GRAPH_SCOPES = ["Tasks.ReadWrite", "User.Read", "openid", "profile", "offline_access"];

async function getClientSecret(): Promise<string> {
  const env = loadEnv();
  if (env.NODE_ENV !== "production" && env.AZURE_CLIENT_SECRET) {
    return env.AZURE_CLIENT_SECRET;
  }
  const store = getSecretStore();
  return store.getSecret("graph-client-secret");
}

function cachePluginFor(connectionId: string, repo: ConnectionsRepo): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext) {
      const serialized = await repo.loadTokenCache(connectionId);
      if (serialized) ctx.tokenCache.deserialize(serialized);
    },
    async afterCacheAccess(ctx: TokenCacheContext) {
      if (ctx.cacheHasChanged) {
        await repo.saveTokenCache(connectionId, ctx.tokenCache.serialize());
      }
    },
  };
}

/**
 * A short-lived MSAL app instance is created per connection per request
 * because the cache plugin closes over `connectionId`. MSAL Node app
 * construction is cheap (no network call), so this is not a performance
 * concern in a scale-to-zero deployment.
 */
export async function createMsalAppForConnection(
  connectionId: string,
  repo: ConnectionsRepo = new ConnectionsRepo()
): Promise<ConfidentialClientApplication> {
  const env = loadEnv();
  const clientSecret = await getClientSecret();
  const config: Configuration = {
    auth: {
      clientId: env.AZURE_CLIENT_ID,
      authority: AUTHORITY,
      clientSecret,
    },
    cache: { cachePlugin: cachePluginFor(connectionId, repo) },
  };
  return new ConfidentialClientApplication(config);
}

/** Same as above but without an existing connection — used during initial sign-in. */
export async function createMsalAppStateless(): Promise<ConfidentialClientApplication> {
  const env = loadEnv();
  const clientSecret = await getClientSecret();
  return new ConfidentialClientApplication({
    auth: { clientId: env.AZURE_CLIENT_ID, authority: AUTHORITY, clientSecret },
  });
}

export type { AuthenticationResult };
