import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { hasModule, type ToolContext } from "./context.js";
import { registerAccountTools } from "./accounts.js";
import { registerReadTools } from "./readTools.js";
import { registerWriteTools } from "./writeTools.js";
import { registerBulkRenameTools } from "./bulkRenameTools.js";
import { registerBulkTools } from "./bulkTools.js";
import { registerShoppingTools } from "./shoppingTools.js";
import { registerMailTools } from "./m365/mailTools.js";
import { registerMailRuleTools } from "./m365/mailRuleTools.js";
import { registerCalendarTools } from "./m365/calendarTools.js";
import { registerContactsTools } from "./m365/contactsTools.js";
import { registerFilesTools } from "./m365/filesTools.js";
import { registerSitesTools } from "./m365/sitesTools.js";
import { registerTeamsTools } from "./m365/teamsTools.js";

/**
 * Account/profile discovery tools are always registered; every other tool
 * family only when its module is enabled for this endpoint (see
 * loadProfileContext): /mcp/{alias} is To Do only, /m365/{alias} gets every
 * module its connection has consented to. Send tools inside the families
 * additionally require ctx.allowSend.
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
  if (hasModule(ctx, "mail")) registerMailTools(server, ctx);
  if (hasModule(ctx, "mail_rules")) registerMailRuleTools(server, ctx);
  if (hasModule(ctx, "calendar")) registerCalendarTools(server, ctx);
  if (hasModule(ctx, "contacts")) registerContactsTools(server, ctx);
  if (hasModule(ctx, "files")) registerFilesTools(server, ctx);
  if (hasModule(ctx, "sites")) registerSitesTools(server, ctx);
  // Teams tools check "teams" and "teams_channels" themselves.
  if (hasModule(ctx, "teams") || hasModule(ctx, "teams_channels")) registerTeamsTools(server, ctx);
}
