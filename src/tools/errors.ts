import { isConsentRequired, isThrottled, toGraphErrorInfo } from "../graph/client.js";
import { isAdminConsentError, buildAdminConsentError } from "../graph/upstreamOAuth.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";

/** Structured content returned inside an MCP tool's isError result. */
export interface ToolErrorPayload {
  errorType: "consent_required" | "throttled" | "not_found" | "graph_error" | "validation_error" | "unknown";
  message: string;
  retryAfterSeconds?: number;
  adminConsentUrl?: string;
}

export async function toToolError(err: unknown, connectionId?: string): Promise<ToolErrorPayload> {
  if (isAdminConsentError(err)) {
    let tenantId = "unknown";
    if (connectionId) {
      const connection = await new ConnectionsRepo().get(connectionId);
      if (connection) tenantId = connection.tenantId;
    }
    const info = buildAdminConsentError(tenantId, (err as Error).message);
    return { errorType: "consent_required", message: info.message, adminConsentUrl: info.adminConsentUrl };
  }
  if (isConsentRequired(err)) {
    return { errorType: "consent_required", message: "Microsoft Graph denied access; the connection may need to be re-authorized." };
  }
  if (isThrottled(err)) {
    const info = toGraphErrorInfo(err);
    return {
      errorType: "throttled",
      message: "Microsoft Graph rate-limited this request. Please retry shortly.",
      retryAfterSeconds: info.retryAfterSeconds,
    };
  }
  const info = toGraphErrorInfo(err);
  if (info.statusCode === 404) {
    return { errorType: "not_found", message: "The requested list, task, or checklist item was not found." };
  }
  if (err instanceof Error) {
    return { errorType: "unknown", message: err.message };
  }
  return { errorType: "unknown", message: "An unexpected error occurred." };
}
