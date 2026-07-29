import rateLimit from "express-rate-limit";

/** Generic rate limiter for the admin surface and OAuth endpoints. */
export const adminRateLimiter = rateLimit({
  windowMs: 60_000,
  limit: 60,
  standardHeaders: "draft-7",
  legacyHeaders: false,
});

/** Tighter limiter for login/token endpoints, which are higher-value targets. */
export const authRateLimiter = rateLimit({
  windowMs: 60_000,
  limit: 20,
  standardHeaders: "draft-7",
  legacyHeaders: false,
});
