import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { registerAccountTools } from "./accounts.js";
import { registerReadTools } from "./readTools.js";
import { registerWriteTools } from "./writeTools.js";

export function registerAllTools(server: McpServer, ctx: ToolContext): void {
  registerAccountTools(server, ctx);
  registerReadTools(server, ctx);
  registerWriteTools(server, ctx);
}
