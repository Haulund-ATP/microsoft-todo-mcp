import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { ProfilesRepo } from "../storage/profilesRepo.js";
import type { ToolContext } from "./context.js";

/** Account/profile discovery tools — always safe, read-only, no Graph calls. */
export function registerAccountTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "list_connections",
    {
      title: "List Microsoft connections",
      description:
        "Lists every Microsoft account connection known to this MCP server (personal and organizational), with alias, status, and last-used time. Does not return tokens or secrets.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const connections = await new ConnectionsRepo().list();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              connections.map((c) => ({
                connectionId: c.connectionId,
                alias: c.alias,
                accountType: c.accountType,
                displayName: c.displayName,
                maskedEmail: c.maskedEmail,
                status: c.status,
                createdAt: c.createdAt,
                lastUsedAt: c.lastUsedAt,
              })),
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.registerTool(
    "get_active_connection",
    {
      title: "Get active connection for this profile",
      description:
        "Returns the single Microsoft connection this MCP endpoint is bound to (for /mcp/personal, /mcp/work, /mcp/{alias}), or null if this is the universal /mcp endpoint which requires an explicit connection_id per call.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      if (!ctx.boundConnectionId) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ bound: false, profileAlias: ctx.profileAlias }, null, 2),
            },
          ],
        };
      }
      const connection = await new ConnectionsRepo().get(ctx.boundConnectionId);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ bound: true, profileAlias: ctx.profileAlias, connection }, null, 2),
          },
        ],
      };
    }
  );

  server.registerTool(
    "list_profiles",
    {
      title: "List profiles",
      description: "Lists every profile-to-connection binding (e.g. personal -> connection X, work -> connection Y).",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const profiles = await new ProfilesRepo().list();
      return { content: [{ type: "text", text: JSON.stringify(profiles, null, 2) }] };
    }
  );
}

export const connectionIdInputShape = {
  connection_id: z
    .string()
    .optional()
    .describe(
      "Required only when calling the universal /mcp endpoint (not needed on /mcp/personal, /mcp/work, or other bound profile routes)."
    ),
};
