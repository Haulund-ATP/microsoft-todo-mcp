import { createGraphClientForConnection } from "../graph/client.js";
import { resolveConnectionId, type ToolContext } from "./context.js";
import { toToolError } from "./errors.js";
import { childLogger } from "../logging/logger.js";

export interface ToolTextResult {
  [key: string]: unknown;
  content: Array<{ type: "text"; text: string }>;
  isError?: boolean;
}

/**
 * Resolves the connection for this call, builds a Graph client for it, and
 * runs `fn`, converting any thrown error into a structured, LLM-readable
 * tool error result rather than letting the transport surface a raw stack
 * trace (which could leak internals).
 */
export async function runGraphTool(
  ctx: ToolContext,
  toolName: string,
  explicitConnectionId: string | undefined,
  fn: (client: ReturnType<typeof createGraphClientForConnection>, connectionId: string) => Promise<unknown>
): Promise<ToolTextResult> {
  const log = childLogger({ correlationId: ctx.correlationId, toolName, profileAlias: ctx.profileAlias });
  const started = Date.now();
  let connectionId: string | undefined;
  try {
    connectionId = await resolveConnectionId(ctx, explicitConnectionId);
    const client = createGraphClientForConnection(connectionId);
    const result = await fn(client, connectionId);
    log.info({ resultStatus: "success", responseTimeMs: Date.now() - started }, "tool completed");
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  } catch (err) {
    const payload = await toToolError(err, connectionId);
    log.warn(
      {
        resultStatus: payload.errorType === "throttled" ? "throttled" : payload.errorType === "consent_required" ? "consent_required" : "error",
        responseTimeMs: Date.now() - started,
      },
      "tool failed"
    );
    return { content: [{ type: "text", text: JSON.stringify(payload, null, 2) }], isError: true };
  }
}
