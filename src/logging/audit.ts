import { logger } from "./logger.js";

/**
 * Audit log for admin/connection lifecycle events. Distinct from the
 * general request logger so it can be routed/retained differently if
 * needed. Same redaction rule applies: no secrets, no PII, no task content.
 */

export type AuditAction =
  | "connection.created"
  | "connection.reauth"
  | "connection.revoked"
  | "connection.deleted"
  | "connection.alias_changed"
  | "profile.bound"
  | "profile.unbound"
  | "admin.login"
  | "admin.login_denied"
  | "oauth.client_registered"
  | "oauth.token_issued"
  | "oauth.token_revoked"
  | "oauth.refresh_reused" // replay-detection signal
  | "oauth.refresh_concurrent" // reuse inside the grace window, rotated instead of revoked
  | "connection.admin_consent"
  | "m365.send" // mail or Teams message sent (counts only)
  | "m365.delete" // mail, event, file or message deleted (ids/counts only)
  | "m365.rule"; // inbox rule created, changed or deleted (counts only, never names or conditions)

export interface AuditEvent {
  action: AuditAction;
  correlationId: string;
  connectionId?: string;
  profileAlias?: string;
  actorClaim?: string; // stable identity claim, never email
  outcome: "success" | "denied" | "error";
  detail?: string; // short, non-sensitive free text
}

export function audit(event: AuditEvent): void {
  logger.info({ audit: true, ...event }, `audit:${event.action}`);
}
