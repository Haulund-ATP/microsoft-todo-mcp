import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { ToolContext } from "./context.js";
import { connectionIdInputShape } from "./accounts.js";
import { runGraphTool } from "./helpers.js";
import { buildShoppingCandidates } from "../graph/shoppingCandidates.js";

const isoDate = z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "Use YYYY-MM-DD.");

export const buildShoppingCandidatesInputShape = {
  recipe_list_id: z.string().min(1).describe("The meal-plan list whose tasks carry recipe JSON notes."),
  from: isoDate.describe("First planned date to include (YYYY-MM-DD)."),
  to: isoDate.describe("Last planned date to include (YYYY-MM-DD)."),
  staging_list_id: z.string().min(1).describe("List the items to check are written to (e.g. 'Indkøb – afklaring')."),
  shopping_list_id: z.string().min(1).describe("The household shopping list, completed items included."),
  pantry_list_id: z.string().min(1).optional().describe("Optional list of staples always at home; matching ingredients are skipped."),
  freezer_list_id: z.string().min(1).optional().describe("Optional freezer list; meat found there is not bought."),
  dry_run: z.boolean().default(true),
  ...connectionIdInputShape,
};

export function registerShoppingTools(server: McpServer, ctx: ToolContext): void {
  server.registerTool(
    "build_shopping_candidates",
    {
      title: "Build shopping candidates from a meal plan",
      description:
        "Collects what the recipes planned between from and to need and writes it, one task per item, to a staging " +
        "list where the household ticks off what they already have; Home Assistant then moves the rest to the " +
        "shopping list. Each recipe task's note must be JSON with dates_planned (YYYY-MM-DD list) and kød / " +
        "ingredienser, each a list of {vare: [alternatives, preferred first], mængde}. Meat found in the freezer " +
        "list and staples on the pantry list are left out, items already open on the shopping list or the staging " +
        "list are skipped, and amounts for the same item are combined. When a completed shopping item matches, the " +
        "staging task's note carries its id so that item is reopened instead of a duplicate being added. Defaults to " +
        "a read-only preview (dry_run=true); use dry_run=false to write. Safe to rerun: items already staged are " +
        "skipped. recipes_missing_ingredients lists planned recipes whose note has no ingredienser — add those by " +
        "hand or fill in the note.",
      inputSchema: buildShoppingCandidatesInputShape,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) =>
      runGraphTool(ctx, "build_shopping_candidates", args.connection_id, (client) =>
        buildShoppingCandidates(client, {
          recipeListId: args.recipe_list_id,
          from: args.from,
          to: args.to,
          stagingListId: args.staging_list_id,
          shoppingListId: args.shopping_list_id,
          pantryListId: args.pantry_list_id,
          freezerListId: args.freezer_list_id,
          dryRun: args.dry_run,
        })
      )
  );
}
