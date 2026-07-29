import { Router } from "express";
import { randomUUID } from "node:crypto";
import { z } from "zod";
import { OAuthClientsRepo } from "../storage/oauthRepo.js";
import { audit } from "../logging/audit.js";

/**
 * OAuth 2.0 Dynamic Client Registration (RFC 7591), scoped down to what MCP
 * connectors need: a public client (PKCE, no client secret) with one or
 * more redirect URIs. See docs/oauth.md for why dynamic registration (over
 * pre-registration or Client ID Metadata Documents) was chosen as the
 * primary path, with pre-registration documented as a fallback.
 */

const registerRequestSchema = z.object({
  client_name: z.string().min(1).max(200),
  redirect_uris: z.array(z.string().url()).min(1).max(10),
  token_endpoint_auth_method: z.literal("none").optional(),
  grant_types: z.array(z.enum(["authorization_code", "refresh_token"])).optional(),
  response_types: z.array(z.literal("code")).optional(),
});

export const registerRouter = Router();

registerRouter.post("/oauth/register", async (req, res) => {
  const parsed = registerRequestSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: "invalid_client_metadata", error_description: parsed.error.message });
    return;
  }

  for (const uri of parsed.data.redirect_uris) {
    let url: URL;
    try {
      url = new URL(uri);
    } catch {
      res.status(400).json({ error: "invalid_redirect_uri" });
      return;
    }
    if (url.protocol !== "https:" && !(url.protocol === "http:" && url.hostname === "localhost")) {
      res.status(400).json({ error: "invalid_redirect_uri", error_description: "redirect_uris must use https (or http on localhost)." });
      return;
    }
  }

  const clientId = randomUUID();
  const repo = new OAuthClientsRepo();
  await repo.register({
    clientId,
    clientName: parsed.data.client_name,
    redirectUris: parsed.data.redirect_uris,
    registrationMethod: "dynamic",
    createdAt: new Date().toISOString(),
  });

  audit({ action: "oauth.client_registered", correlationId: req.correlationId, outcome: "success", detail: clientId });

  res.status(201).json({
    client_id: clientId,
    client_name: parsed.data.client_name,
    redirect_uris: parsed.data.redirect_uris,
    token_endpoint_auth_method: "none",
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
  });
});
