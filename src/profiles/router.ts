import { Router, type Request, type Response } from "express";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { requireBearer } from "../oauth/requireBearer.js";
import type { ResourceKind } from "../oauth/resources.js";
import { registerAllTools } from "../tools/registerAll.js";
import { loadProfileContext } from "../tools/context.js";
import { isValidAlias, RESERVED_ALIASES } from "../storage/profilesRepo.js";
import { childLogger } from "../logging/logger.js";
import { captureErrorBody, clientRef, summarizeJsonRpc } from "../logging/mcpDiagnostics.js";

/**
 * Mounts the MCP Streamable HTTP transport on two endpoint families, each its
 * own OAuth protected resource (see src/oauth/resources.ts):
 *
 *   To Do — exactly what this server always was:
 *     /mcp                 universal, explicit connection_id per tool call
 *     /mcp/{profileAlias}  bound to that profile's connection
 *
 *   Microsoft 365 — every module the connection has consented to (To Do,
 *   mail, calendar, contacts, files, SharePoint, Teams):
 *     /m365                universal, explicit connection_id, never send tools
 *     /m365/{profileAlias} bound; send tools only if the profile allows it
 *
 * Each request builds a fresh McpServer + transport (stateless mode) since
 * Container Apps consumption-plan replicas can scale to zero between
 * calls and we do not want in-memory session affinity requirements.
 */

const SERVER_NAME = "microsoft-todo-mcp";
const SERVER_VERSION = process.env.npm_package_version ?? "0.1.0";

const MOUNTS: Record<ResourceKind, string> = { todo: "/mcp", m365: "/m365" };

export const profilesRouter = Router();

async function handleMcpRequest(kind: ResourceKind, profileAlias: string, req: Request, res: Response): Promise<void> {
  const correlationId = req.correlationId ?? randomUUID();
  const log = childLogger({ correlationId, profileAlias, route: req.path });
  const universal = profileAlias === "mcp";

  if (!universal && !isValidAlias(profileAlias)) {
    log.warn({ rejectReason: "unknown_profile", endpoint: kind }, "mcp request rejected");
    res.status(404).json({ error: "unknown_profile" });
    return;
  }

  const { boundConnectionId, modules, allowSend } = await loadProfileContext(profileAlias, kind);
  if (!universal && !boundConnectionId) {
    log.warn({ rejectReason: "profile_not_bound", endpoint: kind }, "mcp request rejected");
    res.status(404).json({
      error: "profile_not_bound",
      error_description: `Profile "${profileAlias}" exists as a valid alias but has no connection bound yet. Bind it from /accounts first.`,
    });
    return;
  }

  const name = kind === "m365" ? `${SERVER_NAME}:m365:${profileAlias}` : `${SERVER_NAME}:${profileAlias}`;
  const server = new McpServer({ name, version: SERVER_VERSION });
  registerAllTools(server, { profileAlias, boundConnectionId, correlationId, modules, allowSend });

  // Content-free trace of what the client asked for and how we answered, so
  // connector problems (e.g. a client that never finishes connecting) can be
  // read from the logs. See src/logging/mcpDiagnostics.ts.
  const rpc = summarizeJsonRpc(req.body);
  const errorBody = captureErrorBody(res);
  const started = Date.now();
  const mcpLog = {
    endpoint: kind,
    httpMethod: req.method,
    rpc,
    protocolVersionHeader: req.header("mcp-protocol-version"),
    sessionIdHeader: Boolean(req.header("mcp-session-id")),
    accept: req.header("accept")?.slice(0, 80),
    client: req.mcpAuth ? clientRef(req.mcpAuth.clientId) : undefined,
    toolsRegisteredFor: modules,
    allowSend,
  };
  log.info({ mcp: mcpLog }, "mcp request");
  res.on("finish", () => {
    const status = res.statusCode;
    const entry = { mcp: { endpoint: kind, httpMethod: req.method, methods: rpc?.methods, status, ms: Date.now() - started, error: status >= 400 ? errorBody() : undefined } };
    if (status >= 400) log.warn(entry, "mcp response error");
    else log.info(entry, "mcp response");
  });

  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  res.on("close", () => {
    transport.close().catch(() => undefined);
    server.close().catch(() => undefined);
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch {
    log.error({ resultStatus: "error" }, "MCP transport error");
    if (!res.headersSent) {
      res.status(500).json({ error: "internal_error" });
    }
  }
}

for (const kind of Object.keys(MOUNTS) as ResourceKind[]) {
  const mount = MOUNTS[kind];

  // The bearer check is attached per route (not as a prefix middleware) so it
  // sees req.params.alias: every /m365/{alias} is its own OAuth resource.
  profilesRouter.all(mount, requireBearer(kind), (req, res) => {
    void handleMcpRequest(kind, "mcp", req, res);
  });

  profilesRouter.all(`${mount}/:alias`, (req, res, next) => {
    if (RESERVED_ALIASES.has(req.params.alias ?? "")) {
      res.status(404).json({ error: "unknown_profile" });
      return;
    }
    next();
  }, requireBearer(kind), (req, res) => {
    void handleMcpRequest(kind, req.params.alias!, req, res);
  });
}
