import express, { type Request, type Response, type NextFunction } from "express";
import helmet from "helmet";
import cookieParser from "cookie-parser";
import { pinoHttp } from "pino-http";
import { logger } from "../logging/logger.js";
import { correlationIdMiddleware } from "../logging/correlation.js";
import { healthRouter } from "./health.js";
import { wellKnownRouter } from "../oauth/wellKnown.js";
import { registerRouter } from "../oauth/register.js";
import { authorizeRouter } from "../oauth/authorize.js";
import { tokenRouter } from "../oauth/token.js";
import { revokeRouter } from "../oauth/revoke.js";
import { ownerLoginRouter } from "../admin/ownerLogin.js";
import { connectRouter } from "../admin/connectRouter.js";
import { accountsRouter } from "../admin/accountsRouter.js";
import { profilesRouter } from "../profiles/router.js";
import { adminRateLimiter, authRateLimiter } from "../admin/rateLimit.js";

export function createApp() {
  const app = express();
  app.disable("x-powered-by");
  app.set("trust proxy", 1); // Container Apps ingress terminates TLS in front of us

  app.use(correlationIdMiddleware);
  app.use(
    pinoHttp({
      logger,
      genReqId: (req: Request) => req.correlationId,
      customProps: (req: Request) => ({ route: req.path }),
      redact: ["req.headers.authorization", "req.headers.cookie"],
    })
  );

  app.use(
    helmet({
      contentSecurityPolicy: {
        directives: {
          defaultSrc: ["'self'"],
          scriptSrc: ["'self'"],
          styleSrc: ["'self'", "'unsafe-inline'"], // admin page uses a small inline <style>
          imgSrc: ["'self'"],
          connectSrc: ["'self'"],
          frameAncestors: ["'none'"],
        },
      },
      referrerPolicy: { policy: "no-referrer" },
      crossOriginResourcePolicy: { policy: "same-origin" },
    })
  );

  app.use(cookieParser());
  app.use(express.json({ limit: "256kb" }));
  app.use(express.urlencoded({ extended: false, limit: "64kb" }));

  // Health/discovery endpoints: unauthenticated, no rate limit needed
  // beyond what Container Apps ingress already provides.
  app.use(healthRouter);
  app.use(wellKnownRouter);

  // OAuth 2.1 authorization server + owner login + Microsoft connect flow.
  app.use(authRateLimiter);
  app.use(registerRouter);
  app.use(authorizeRouter);
  app.use(tokenRouter);
  app.use(revokeRouter);
  app.use(ownerLoginRouter);
  app.use(connectRouter);

  // Admin UI.
  app.use(adminRateLimiter);
  app.use(accountsRouter);

  app.get("/", (_req, res) => {
    res.redirect("/accounts");
  });

  // MCP transport (bearer-protected inside the router).
  app.use(profilesRouter);

  app.use((req: Request, res: Response) => {
    res.status(404).json({ error: "not_found" });
  });

  // Centralized error handler: never leak stack traces or internals.
  app.use((err: unknown, req: Request, res: Response, _next: NextFunction) => {
    logger.error({ correlationId: req.correlationId, route: req.path }, "unhandled request error");
    if (!res.headersSent) {
      res.status(500).json({ error: "internal_error" });
    }
  });

  return app;
}
