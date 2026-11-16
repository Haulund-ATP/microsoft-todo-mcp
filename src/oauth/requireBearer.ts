import type { NextFunction, Request, Response } from "express";
import { verifyAccessToken } from "./jwt.js";
import { protectedResourceMetadataUrl, resourceUrl, type ResourceKind } from "./resources.js";
import { childLogger } from "../logging/logger.js";

declare module "express-serve-static-core" {
  interface Request {
    mcpAuth?: { subject: string; clientId: string; scope: string };
  }
}

/**
 * Protects /mcp* endpoints. On failure, responds per RFC 6750 with a
 * WWW-Authenticate header pointing at the Protected Resource Metadata
 * document so a compliant MCP client can (re)discover the authorization
 * server and re-initiate the OAuth flow, rather than failing silently.
 */
export function requireBearer(kind: ResourceKind = "todo") {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    const log = childLogger({ correlationId: req.correlationId, route: req.path });
    const header = req.header("authorization");
    // Microsoft 365 profiles are separate resources (see resources.ts); the
    // route's :alias picks which. Mount this on routes, not as a prefix
    // middleware, so req.params.alias is set.
    const alias = kind === "m365" ? req.params.alias : undefined;
    const audience = resourceUrl(kind, alias);
    const challenge = `Bearer resource_metadata="${protectedResourceMetadataUrl(kind, alias)}"`;

    if (!header?.startsWith("Bearer ")) {
      // Logged at info, not warn: clients routinely probe without a token
      // first to discover the authorization server from this challenge.
      log.info({ rejectReason: "bearer_missing" }, "bearer token missing");
      res.setHeader("WWW-Authenticate", challenge);
      res.status(401).json({ error: "invalid_token", error_description: "Missing bearer token." });
      return;
    }
    const token = header.slice("Bearer ".length).trim();
    try {
      // Audience must be exactly this endpoint's resource: a To Do token is
      // rejected on /m365 and vice versa.
      const claims = await verifyAccessToken(token, audience);
      req.mcpAuth = { subject: String(claims.sub), clientId: String(claims.client_id), scope: String(claims.scope) };
      next();
    } catch (err) {
      log.warn({ rejectReason: categorizeBearerError(err) }, "bearer token rejected");
      res.setHeader("WWW-Authenticate", `${challenge}, error="invalid_token"`);
      res.status(401).json({ error: "invalid_token" });
    }
  };
}

/**
 * Maps a jose verification error to a short log category. Returns only
 * string literals — nothing from the error object itself is logged.
 */
export function categorizeBearerError(err: unknown): string {
  const { code, claim } = (err ?? {}) as { code?: string; claim?: string };
  switch (code) {
    case "ERR_JWT_EXPIRED":
      return "bearer_expired";
    case "ERR_JWT_CLAIM_VALIDATION_FAILED":
      return claim === "aud" ? "bearer_audience_mismatch" : claim === "iss" ? "bearer_issuer_mismatch" : "bearer_claim_invalid";
    case "ERR_JWS_SIGNATURE_VERIFICATION_FAILED":
      return "bearer_bad_signature";
    case "ERR_JWS_INVALID":
    case "ERR_JWT_INVALID":
      return "bearer_malformed";
    default:
      return "bearer_invalid";
  }
}
