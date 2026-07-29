import { Router, type Request, type Response } from "express";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { requireBearer } from "../oauth/requireBearer.js";
import { registerAllTools } from "../tools/registerAll.js";
import { loadBoundConnectionIdForProfile } from "../tools/context.js";
import { isValidAlias, RESERVED_ALIASES } from "../storage/profilesRepo.js";
import { childLogger } from "../logging/logger.js";

/**
 * Mounts the MCP Streamable HTTP transport at:
 *   /mcp              — universal, requires explicit connection_id per tool call
 *   /mcp/personal      — bound to the "personal" profile's connection
 *   /mcp/work          — bound to the "work" profile's connection
 *   /mcp/{profileAlias} — any other validated, non-reserved alias
 *
 * Each request builds a fresh McpServer + transport (stateless mode) since
 * Container Apps consumption-plan replicas can scale to zero between
 * calls and we do not want in-memory session affinity requirements.
 */

const SERVER_NAME = "microsoft-todo-mcp";
const SERVER_VERSION = process.env.npm_package_version ?? "0.1.0";

export const profilesRouter = Router();

profilesRouter.use("/mcp", requireBearer());
profilesRouter.use("/mcp/:alias", requireBearer());

async function handleMcpRequest(profileAlias: string, req: Request, res: Response): Promise<void> {
  const correlationId = req.correlationId ?? randomUUID();
  const log = childLogger({ correlationId, profileAlias, route: req.path });

  if (profileAlias !== "mcp" && !isValidAlias(profileAlias)) {
    res.status(404).json({ error: "unknown_profile" });
    return;
  }

  const boundConnectionId = await loadBoundConnectionIdForProfile(profileAlias);
  if (profileAlias !== "mcp" && !boundConnectionId) {
    res.status(404).json({
      error: "profile_not_bound",
      error_description: `Profile "${profileAlias}" exists as a valid alias but has no connection bound yet. Bind it from /accounts first.`,
    });
    return;
  }

  const server = new McpServer({ name: `${SERVER_NAME}:${profileAlias}`, version: SERVER_VERSION });
  registerAllTools(server, { profileAlias, boundConnectionId, correlationId });

  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  res.on("close", () => {
    transport.close().catch(() => undefined);
    server.close().catch(() => undefined);
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (_err) {
    log.error({ resultStatus: "error" }, "MCP transport error");
    if (!res.headersSent) {
      res.status(500).json({ error: "internal_error" });
    }
  }
}

profilesRouter.all("/mcp", (req, res) => {
  void handleMcpRequest("mcp", req, res);
});

profilesRouter.all("/mcp/:alias", (req, res) => {
  const alias = req.params.alias;
  if (RESERVED_ALIASES.has(alias)) {
    res.status(404).json({ error: "unknown_profile" });
    return;
  }
  void handleMcpRequest(alias, req, res);
});
