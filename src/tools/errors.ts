import { isConsentRequired, isThrottled, toGraphErrorInfo } from "../graph/client.js";
import { isAdminConsentError, buildAdminConsentError } from "../graph/upstreamOAuth.js";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";

/**
 * Thrown by checklist write tools when the target list is shared and the
 * current connection does not own it. See docs/security.md#shared-list-
 * checklist-risk for why this gate exists: Microsoft's To Do shared-list
 * sync has widely-reported (if not officially documented) data-loss issues
 * unrelated to how this server calls the Graph API.
 */
export class SharedListChecklistWriteBlockedError extends Error {
  constructor(public readonly listId: string) {
    super(
      `Task list "${listId}" is shared and not owned by this connection. Checklist-item writes on shared lists ` +
        `can be silently lost or reset by Microsoft To Do's shared-list sync (a known issue outside this server's ` +
        `control — see docs/security.md). Pass acknowledge_shared_list_risk: true to proceed anyway.`
    );
    this.name = "SharedListChecklistWriteBlockedError";
  }
}

/** Structured content returned inside an MCP tool's isError result. */
export interface ToolErrorPayload {
  errorType:
    | "consent_required"
    | "throttled"
    | "not_found"
    | "graph_error"
    | "validation_error"
    | "shared_list_write_blocked"
    | "unknown";
  message: string;
  retryAfterSeconds?: number;
  adminConsentUrl?: string;
}

export async function toToolError(err: unknown, connectionId?: string): Promise<ToolErrorPayload> {
  if (err instanceof SharedListChecklistWriteBlockedError) {
    return { errorType: "shared_list_write_blocked", message: err.message };
  }
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
