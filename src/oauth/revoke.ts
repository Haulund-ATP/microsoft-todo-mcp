import { Router } from "express";
import { z } from "zod";
import { OAuthTokensRepo } from "../storage/oauthRepo.js";
import { audit } from "../logging/audit.js";

/** RFC 7009 token revocation for refresh tokens (access tokens are short-lived JWTs and expire on their own). */

const revokeSchema = z.object({
  token: z.string().min(1),
  token_type_hint: z.enum(["refresh_token", "access_token"]).optional(),
});

export const revokeRouter = Router();

revokeRouter.post("/oauth/revoke", async (req, res) => {
  const parsed = revokeSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "invalid_request" });
    return;
  }
  // Per RFC 7009, unknown/invalid tokens are not an error — always 200.
  const tokensRepo = new OAuthTokensRepo();
  await tokensRepo.revoke(parsed.data.token).catch(() => undefined);
  audit({ action: "oauth.token_revoked", correlationId: req.correlationId, outcome: "success" });
  res.status(200).end();
});
