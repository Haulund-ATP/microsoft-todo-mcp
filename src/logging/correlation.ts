import { randomUUID } from "node:crypto";
import type { NextFunction, Request, Response } from "express";

declare module "express-serve-static-core" {
  interface Request {
    correlationId: string;
  }
}

/** Attaches a per-request correlation id, honoring an inbound header if present. */
export function correlationIdMiddleware(req: Request, res: Response, next: NextFunction): void {
  const inbound = req.header("x-correlation-id");
  req.correlationId = inbound && /^[a-zA-Z0-9-]{1,64}$/.test(inbound) ? inbound : randomUUID();
  res.setHeader("x-correlation-id", req.correlationId);
  next();
}
