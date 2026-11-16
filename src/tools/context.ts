import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { ProfilesRepo } from "../storage/profilesRepo.js";
import { MODULE_IDS, parseModuleList, type ModuleId } from "../graph/modules.js";
import type { ResourceKind } from "../oauth/resources.js";

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
  /** Modules whose tools are registered on this endpoint. Defaults to To Do only. */
  modules?: readonly ModuleId[];
  /** Whether send tools (mail, Teams messages) are registered. Defaults to false. */
  allowSend?: boolean;
}

export function hasModule(ctx: ToolContext, module: ModuleId): boolean {
  return (ctx.modules ?? ["todo"]).includes(module);
}

export interface ProfileContext {
  boundConnectionId?: string;
  modules: ModuleId[];
  allowSend: boolean;
}

/**
 * Resolves what an MCP route exposes. The endpoint family decides:
 *
 * - "todo" (/mcp, /mcp/{alias}) is always To Do only and never sends —
 *   exactly what this server was before Microsoft 365 modules existed.
 * - "m365" (/m365/{alias}) exposes every module the bound connection has
 *   consented to, and send tools only if the profile allows sending. The
 *   universal /m365 route exposes every module (each call names its
 *   connection, which must have consented) but never send tools.
 */
export async function loadProfileContext(profileAlias: string, kind: ResourceKind = "todo"): Promise<ProfileContext> {
  const universal = profileAlias === "mcp";
  if (universal) {
    return { modules: kind === "m365" ? [...MODULE_IDS] : ["todo"], allowSend: false };
  }
  const binding = await new ProfilesRepo().get(profileAlias);
  if (!binding) return { modules: [], allowSend: false };
  if (kind === "todo") {
    return { boundConnectionId: binding.connectionId, modules: ["todo"], allowSend: false };
  }
  const connection = await new ConnectionsRepo().get(binding.connectionId);
  return {
    boundConnectionId: binding.connectionId,
    modules: connection ? parseModuleList(connection.grantedModules) : [],
    allowSend: binding.allowSend === true,
  };
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
