import { getTableClient, TABLE_NAMES } from "./clients.js";
import type { ProfileBinding } from "./types.js";

const PARTITION_KEY = "profile";

/** Routes/aliases that a profile can never claim, since they are reserved endpoints. */
export const RESERVED_ALIASES = new Set([
  "mcp",
  "accounts",
  "health",
  "ready",
  "version",
  ".well-known",
]);

const ALIAS_PATTERN = /^[a-z0-9]([a-z0-9-]{0,48}[a-z0-9])?$/;

export function isValidAlias(alias: string): boolean {
  return ALIAS_PATTERN.test(alias) && !RESERVED_ALIASES.has(alias);
}

interface ProfileEntity extends ProfileBinding {
  partitionKey: string;
  rowKey: string;
}

export class ProfilesRepo {
  private table() {
    return getTableClient(TABLE_NAMES.profiles);
  }

  async list(): Promise<ProfileBinding[]> {
    const out: ProfileBinding[] = [];
    for await (const entity of this.table().listEntities<ProfileEntity>({
      queryOptions: { filter: `PartitionKey eq '${PARTITION_KEY}'` },
    })) {
      const { partitionKey: _pk, rowKey: _rk, ...rest } = entity;
      out.push(rest as ProfileBinding);
    }
    return out;
  }

  async get(profileAlias: string): Promise<ProfileBinding | undefined> {
    try {
      const entity = await this.table().getEntity<ProfileEntity>(PARTITION_KEY, profileAlias);
      const { partitionKey: _pk, rowKey: _rk, ...rest } = entity;
      return rest as ProfileBinding;
    } catch (err) {
      if ((err as { statusCode?: number }).statusCode === 404) return undefined;
      throw err;
    }
  }

  async bind(profileAlias: string, connectionId: string): Promise<ProfileBinding> {
    if (!isValidAlias(profileAlias)) {
      throw new Error(
        `Invalid or reserved profile alias "${profileAlias}". Must match ${ALIAS_PATTERN} and not collide with reserved routes.`
      );
    }
    const binding: ProfileBinding = {
      profileAlias,
      connectionId,
      createdAt: new Date().toISOString(),
    };
    await this.table().upsertEntity(
      { partitionKey: PARTITION_KEY, rowKey: profileAlias, ...binding },
      "Replace"
    );
    return binding;
  }

  async unbind(profileAlias: string): Promise<void> {
    await this.table().deleteEntity(PARTITION_KEY, profileAlias).catch((err) => {
      if ((err as { statusCode?: number }).statusCode !== 404) throw err;
    });
  }
}
