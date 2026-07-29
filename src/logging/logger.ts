import pino from "pino";
import { loadEnv } from "../config/env.js";

/**
 * Structured logger. Hard rule: NEVER log task titles/descriptions, full
 * email addresses, tokens, authorization codes, cookies, bearer headers, or
 * raw Graph responses. Call sites should only pass the fields defined in
 * `LogFields` below (or a subset) — treat any ad-hoc extra field as
 * suspect during code review.
 */

export interface LogFields {
  correlationId?: string;
  toolName?: string;
  profileAlias?: string;
  resultStatus?: "success" | "error" | "consent_required" | "throttled";
  responseTimeMs?: number;
  graphStatusCode?: number;
  throttled?: boolean;
  oauthErrorCategory?: string;
  deploymentVersion?: string;
  route?: string;
  httpMethod?: string;
  httpStatusCode?: number;
}

const REDACT_PATHS = [
  "req.headers.authorization",
  "req.headers.cookie",
  "res.headers['set-cookie']",
  "*.token",
  "*.accessToken",
  "*.refreshToken",
  "*.idToken",
  "*.code",
  "*.client_secret",
  "*.title",
  "*.description",
  "*.email",
];

const env = (() => {
  try {
    return loadEnv();
  } catch {
    // Logger may be constructed before env validation in some test contexts.
    return undefined;
  }
})();

export const logger = pino({
  level: env?.LOG_LEVEL ?? process.env.LOG_LEVEL ?? "info",
  redact: { paths: REDACT_PATHS, censor: "[redacted]" },
  formatters: {
    level(label) {
      return { level: label };
    },
  },
  timestamp: pino.stdTimeFunctions.isoTime,
});

export function childLogger(fields: LogFields) {
  return logger.child(fields);
}
