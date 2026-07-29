import { sign, unsign } from "cookie-signature";
import type { Request, Response } from "express";
import { loadEnv } from "../config/env.js";

/**
 * Owner session cookie: proves the browser belongs to the deployment
 * owner, used both for the /accounts admin page and to authenticate the
 * resource-owner leg of the MCP-side OAuth 2.1 `/oauth/authorize` flow.
 * Signed, HttpOnly, Secure, SameSite=Lax (Lax because the OAuth authorize
 * redirect is a top-level GET navigation from ChatGPT/Claude).
 */

const COOKIE_NAME = "mth_owner_session";
const SESSION_TTL_SECONDS = 60 * 60 * 12; // 12h

export interface OwnerSession {
  ownerClaim: string;
  issuedAt: number;
}

export function setOwnerSessionCookie(res: Response, session: OwnerSession): void {
  const secret = loadEnv().SESSION_COOKIE_SECRET;
  const value = sign(JSON.stringify(session), secret);
  res.cookie(COOKIE_NAME, value, {
    httpOnly: true,
    secure: true,
    sameSite: "lax",
    maxAge: SESSION_TTL_SECONDS * 1000,
    path: "/",
  });
}

export function clearOwnerSessionCookie(res: Response): void {
  res.clearCookie(COOKIE_NAME, { path: "/" });
}

export function readOwnerSession(req: Request): OwnerSession | undefined {
  const raw = req.cookies?.[COOKIE_NAME];
  if (!raw) return undefined;
  const secret = loadEnv().SESSION_COOKIE_SECRET;
  const unsigned = unsign(raw, secret);
  if (!unsigned) return undefined;
  try {
    const session = JSON.parse(unsigned) as OwnerSession;
    if (Date.now() - session.issuedAt > SESSION_TTL_SECONDS * 1000) return undefined;
    return session;
  } catch {
    return undefined;
  }
}

export function isOwner(session: OwnerSession | undefined): boolean {
  if (!session) return false;
  return session.ownerClaim === loadEnv().ADMIN_OWNER_CLAIM_VALUE;
}
