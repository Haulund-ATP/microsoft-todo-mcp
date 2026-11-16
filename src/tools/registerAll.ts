import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { hasModule, type ToolContext } from "./context.js";
import { registerAccountTools } from "./accounts.js";
import { registerReadTools } from "./readTools.js";
import { registerWriteTools } from "./writeTools.js";
import { registerBulkRenameTools } from "./bulkRenameTools.js";
import { registerBulkTools } from "./bulkTools.js";
import { registerShoppingTools } from "./shoppingTools.js";

/**
 * Account/profile discovery tools are always registered; every other tool
 * family only when its module is enabled for this endpoint (see
 * loadProfileContext), so a To Do-only profile never sees mail or file tools.
 */
export function registerAllTools(server: McpServer, ctx: ToolContext): void {
  registerAccountTools(server, ctx);
  if (hasModule(ctx, "todo")) {
    registerReadTools(server, ctx);
    registerWriteTools(server, ctx);
    registerBulkRenameTools(server, ctx);
    registerBulkTools(server, ctx);
    registerShoppingTools(server, ctx);
  }
}
