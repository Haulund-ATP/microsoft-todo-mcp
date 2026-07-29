import { describe, it, expect } from "vitest";
import { registerWriteTools } from "../../src/tools/writeTools.js";
import type { ToolContext } from "../../src/tools/context.js";

/** Captures registered tool callbacks without needing a real McpServer/transport. */
function fakeServer() {
  const handlers = new Map<string, (input: Record<string, unknown>) => Promise<unknown>>();
  return {
    registerTool: (name: string, _config: unknown, handler: (input: Record<string, unknown>) => Promise<unknown>) => {
      handlers.set(name, handler);
    },
    handlers,
  };
}

describe("update_checklist_item — at-least-one-field validation", () => {
  it("rejects a call with neither display_name nor is_checked, without touching Graph", async () => {
    const server = fakeServer();
    const ctx: ToolContext = { profileAlias: "mcp", correlationId: "corr-1" };
    registerWriteTools(server as never, ctx);

    const handler = server.handlers.get("update_checklist_item")!;
    const result = (await handler({ list_id: "L", task_id: "T", item_id: "item-1" })) as {
      isError?: boolean;
      content: Array<{ type: string; text: string }>;
    };

    expect(result.isError).toBe(true);
    const payload = JSON.parse(result.content[0].text);
    expect(payload.errorType).toBe("validation_error");
  });
});
