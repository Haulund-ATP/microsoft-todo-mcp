import { Router } from "express";
import { randomBytes, randomUUID } from "node:crypto";
import { z } from "zod";
import { OAuthClientsRepo, hashSecret } from "../storage/oauthRepo.js";
import { TOKEN_ENDPOINT_AUTH_METHODS } from "./clientAuth.js";
import { childLogger } from "../logging/logger.js";
import { urlHost } from "../logging/mcpDiagnostics.js";
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
  // "none" = public client (PKCE only). Claude registers with the MCP Python
  // SDK default "client_secret_post"; such clients get a secret that the
  // token endpoint then requires (see clientAuth.ts).
  token_endpoint_auth_method: z.enum(TOKEN_ENDPOINT_AUTH_METHODS).optional(),
  grant_types: z.array(z.enum(["authorization_code", "refresh_token"])).optional(),
  response_types: z.array(z.literal("code")).optional(),
});

export const registerRouter = Router();

registerRouter.post("/oauth/register", async (req, res) => {
  const parsed = registerRequestSchema.safeParse(req.body);
  if (!parsed.success) {
    // Field paths only (no values), e.g. ["token_endpoint_auth_method"].
    childLogger({ correlationId: req.correlationId, route: "/oauth/register" }).warn(
      { rejectReason: "invalid_client_metadata", fields: parsed.error.issues.map((i) => i.path.join(".")).slice(0, 10) },
      "client registration rejected"
    );
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
  const authMethod = parsed.data.token_endpoint_auth_method ?? "none";
  // The secret is returned once and only its SHA-256 hash is stored.
  const clientSecret = authMethod === "none" ? undefined : randomBytes(32).toString("base64url");
  const now = new Date();
  const repo = new OAuthClientsRepo();
  await repo.register({
    clientId,
    clientName: parsed.data.client_name,
    redirectUris: parsed.data.redirect_uris,
    registrationMethod: "dynamic",
    createdAt: now.toISOString(),
    tokenEndpointAuthMethod: authMethod,
    ...(clientSecret ? { clientSecretHash: hashSecret(clientSecret) } : {}),
  });

  audit({
    action: "oauth.client_registered",
    correlationId: req.correlationId,
    outcome: "success",
    detail: `${clientId} auth=${authMethod}`,
  });
  childLogger({ correlationId: req.correlationId, route: "/oauth/register" }).info(
    {
      oauthClientId: clientId,
      clientName: parsed.data.client_name.slice(0, 60),
      authMethod,
      redirectHosts: parsed.data.redirect_uris.map((u) => urlHost(u)),
      grantTypes: parsed.data.grant_types,
    },
    "client registered"
  );

  res.status(201).json({
    client_id: clientId,
    client_id_issued_at: Math.floor(now.getTime() / 1000),
    ...(clientSecret ? { client_secret: clientSecret, client_secret_expires_at: 0 } : {}),
    client_name: parsed.data.client_name,
    redirect_uris: parsed.data.redirect_uris,
    token_endpoint_auth_method: authMethod,
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
  });
});
