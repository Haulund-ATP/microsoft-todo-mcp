import { Router } from "express";
import { buildAuthorizationServerMetadata, buildProtectedResourceMetadata } from "./metadata.js";
import { getPublicJwks } from "./signingKeys.js";

export const wellKnownRouter = Router();

wellKnownRouter.get("/.well-known/oauth-protected-resource", (_req, res) => {
  res.json(buildProtectedResourceMetadata());
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
