import type { NextFunction, Request, Response } from "express";
import { isOwner, readOwnerSession } from "./session.js";

/** Gates any admin route behind an authenticated owner session, redirecting to login otherwise. */
export function requireOwner(req: Request, res: Response, next: NextFunction): void {
  const session = readOwnerSession(req);
  if (!isOwner(session)) {
    const returnTo = encodeURIComponent(req.originalUrl);
    res.redirect(`/oauth/owner/login?returnTo=${returnTo}`);
    return;
  }
  next();
}
