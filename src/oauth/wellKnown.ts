import { Router } from "express";
import { buildAuthorizationServerMetadata, buildProtectedResourceMetadata } from "./metadata.js";
import { getPublicJwks } from "./signingKeys.js";

export const wellKnownRouter = Router();

wellKnownRouter.get("/.well-known/oauth-protected-resource", (_req, res) => {
  res.json(buildProtectedResourceMetadata("todo"));
});

// The Microsoft 365 endpoint is its own protected resource (see resources.ts).
// Clients that build the RFC 9728 path-inserted URL from the full endpoint
// (/m365/{alias}) instead of following WWW-Authenticate land here too.
wellKnownRouter.get(["/.well-known/oauth-protected-resource/m365", "/.well-known/oauth-protected-resource/m365/:alias"], (_req, res) => {
  res.json(buildProtectedResourceMetadata("m365"));
});

wellKnownRouter.get("/.well-known/oauth-authorization-server", (_req, res) => {
  res.json(buildAuthorizationServerMetadata());
});

// Alias for clients that probe OIDC-discovery-shaped paths.
wellKnownRouter.get("/.well-known/openid-configuration", (_req, res) => {
  res.json(buildAuthorizationServerMetadata());
});

wellKnownRouter.get("/.well-known/jwks.json", async (_req, res) => {
  res.json(await getPublicJwks());
});
