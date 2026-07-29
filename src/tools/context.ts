import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { ProfilesRepo } from "../storage/profilesRepo.js";

/**
 * Per-request/session context available to every tool handler. `boundConnectionId`
 * is set when the MCP transport was mounted at a profile-specific route
 * (/mcp/personal, /mcp/work, /mcp/{alias}) — in that case tools MUST use it
 * and MUST NOT accept a differing connection_id. On the universal /mcp
 * route, `boundConnectionId` is undefined and every tool requires an
 * explicit `connection_id` argument instead, so the server never guesses
 * between accounts.
 */
export interface ToolContext {
  profileAlias: string; // "mcp" for the universal route
  boundConnectionId?: string;
  correlationId: string;
}

export class UnresolvedConnectionError extends Error {
  constructor() {
    super(
      "This tool requires a connection_id because it was called on the universal /mcp endpoint, which is not bound to a single profile. Call list_connections or list_profiles first, then pass the desired connection_id explicitly."
    );
    this.name = "UnresolvedConnectionError";
  }
}

export async function resolveConnectionId(ctx: ToolContext, explicitConnectionId?: string): Promise<string> {
  if (ctx.boundConnectionId) {
    if (explicitConnectionId && explicitConnectionId !== ctx.boundConnectionId) {
      throw new Error(
        `This profile ("${ctx.profileAlias}") is bound to connection ${ctx.boundConnectionId}; it cannot be used with a different connection_id.`
      );
    }
    return ctx.boundConnectionId;
  }
  if (!explicitConnectionId) {
    throw new UnresolvedConnectionError();
  }
  const repo = new ConnectionsRepo();
  const connection = await repo.get(explicitConnectionId);
  if (!connection) throw new Error(`Unknown connection_id: ${explicitConnectionId}`);
  return explicitConnectionId;
}

export async function loadBoundConnectionIdForProfile(profileAlias: string): Promise<string | undefined> {
  if (profileAlias === "mcp") return undefined;
  const profilesRepo = new ProfilesRepo();
  const binding = await profilesRepo.get(profileAlias);
  return binding?.connectionId;
}
