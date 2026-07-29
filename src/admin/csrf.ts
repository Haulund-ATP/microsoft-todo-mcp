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

function isValidToken(token: string, secret: string): boolean {
  const [nonce, sig] = token.split(".");
  if (!nonce || !sig) return false;
  const expected = sign(nonce, secret);
  const a = Buffer.from(sig, "hex");
  const b = Buffer.from(expected, "hex");
  return a.length === b.length && timingSafeEqual(a, b);
}

/**
 * Reuses the caller's existing, still-valid CSRF cookie instead of always
 * minting a fresh one. Minting unconditionally meant any second page load —
 * a reload, a second tab, returning from an OAuth redirect while the
 * original tab was still open — silently invalidated every token embedded
 * in a page rendered before it, since the double-submit cookie is one
 * global value per browser, not per page render. Reusing the token when
 * it's already valid keeps concurrently open/rendered pages' embedded
 * tokens in sync with the cookie the browser actually holds.
 */
export function issueCsrfToken(req: Request, res: Response): string {
  const secret = loadEnv().CSRF_COOKIE_SECRET;
  const existing = req.cookies?.[CSRF_COOKIE_NAME];
  const token =
    typeof existing === "string" && isValidToken(existing, secret) ? existing : mintToken(secret);
  res.cookie(CSRF_COOKIE_NAME, token, {
    httpOnly: false,
    secure: true,
    sameSite: "strict",
    // Used across /accounts, /admin/connect, and /oauth/authorize —
    // must not be scoped to a single prefix or the cookie won't be sent
    // back on the others' form submissions.
    path: "/",
  });
  return token;
}

function mintToken(secret: string): string {
  const nonce = randomBytes(18).toString("hex");
  return `${nonce}.${sign(nonce, secret)}`;
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
