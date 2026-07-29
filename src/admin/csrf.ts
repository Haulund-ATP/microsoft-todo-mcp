import { randomBytes, createHmac, timingSafeEqual } from "node:crypto";
import type { NextFunction, Request, Response } from "express";
import { loadEnv } from "../config/env.js";

/**
 * Double-submit-cookie CSRF protection for the /accounts admin UI's
 * state-changing POST endpoints. A signed token is set as a readable
 * (non-HttpOnly) cookie; forms must echo it back as a hidden field, and we
 * verify the signature + equality server-side. This avoids pulling in an
 * unmaintained csurf dependency while keeping the primitive simple and
 * auditable.
 */

const CSRF_COOKIE_NAME = "mth_csrf";

function sign(value: string, secret: string): string {
  return createHmac("sha256", secret).update(value).digest("hex");
}

export function issueCsrfToken(res: Response): string {
  const nonce = randomBytes(18).toString("hex");
  const secret = loadEnv().CSRF_COOKIE_SECRET;
  const token = `${nonce}.${sign(nonce, secret)}`;
  res.cookie(CSRF_COOKIE_NAME, token, {
    httpOnly: false,
    secure: true,
    sameSite: "strict",
    path: "/accounts",
  });
  return token;
}

function isValidToken(token: string, secret: string): boolean {
  const [nonce, sig] = token.split(".");
  if (!nonce || !sig) return false;
  const expected = sign(nonce, secret);
  const a = Buffer.from(sig, "hex");
  const b = Buffer.from(expected, "hex");
  return a.length === b.length && timingSafeEqual(a, b);
}

/** Enforces that the submitted `_csrf` field matches the signed cookie. */
export function requireCsrf(req: Request, res: Response, next: NextFunction): void {
  const secret = loadEnv().CSRF_COOKIE_SECRET;
  const cookieToken = req.cookies?.[CSRF_COOKIE_NAME];
  const bodyToken = req.body?._csrf as string | undefined;
  if (
    typeof cookieToken !== "string" ||
    typeof bodyToken !== "string" ||
    cookieToken !== bodyToken ||
    !isValidToken(cookieToken, secret)
  ) {
    res.status(403).json({ error: "invalid_csrf_token" });
    return;
  }
  next();
}
