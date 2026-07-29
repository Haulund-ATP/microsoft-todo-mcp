import type { NextFunction, Request, Response } from "express";
import { loadEnv } from "../config/env.js";
import { verifyAccessToken } from "./jwt.js";
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
export function requireBearer() {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    const log = childLogger({ correlationId: req.correlationId, route: req.path });
    const header = req.header("authorization");
    const base = loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
    const challenge = `Bearer resource_metadata="${base}/.well-known/oauth-protected-resource"`;

    if (!header?.startsWith("Bearer ")) {
      res.setHeader("WWW-Authenticate", challenge);
      res.status(401).json({ error: "invalid_token", error_description: "Missing bearer token." });
      return;
    }
    const token = header.slice("Bearer ".length).trim();
    try {
      const claims = await verifyAccessToken(token, base);
      req.mcpAuth = { subject: String(claims.sub), clientId: String(claims.client_id), scope: String(claims.scope) };
      next();
    } catch (_err) {
      log.warn({ oauthErrorCategory: "invalid_access_token" }, "bearer token rejected");
      res.setHeader("WWW-Authenticate", `${challenge}, error="invalid_token"`);
      res.status(401).json({ error: "invalid_token" });
    }
  };
}
