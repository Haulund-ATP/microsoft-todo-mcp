import { Router } from "express";
import { buildAuthorizationServerMetadata, buildProtectedResourceMetadata } from "./metadata.js";
import { getPublicJwks } from "./signingKeys.js";
import { isValidAlias } from "../storage/profilesRepo.js";

export const wellKnownRouter = Router();

wellKnownRouter.get("/.well-known/oauth-protected-resource", (_req, res) => {
  res.json(buildProtectedResourceMetadata("todo"));
});

// Microsoft 365: one protected resource per profile plus the universal one
// (see resources.ts).
wellKnownRouter.get("/.well-known/oauth-protected-resource/m365", (_req, res) => {
  res.json(buildProtectedResourceMetadata("m365"));
});

wellKnownRouter.get("/.well-known/oauth-protected-resource/m365/:alias", (req, res) => {
  const alias = req.params.alias;
  if (!isValidAlias(alias)) {
    res.status(404).json({ error: "not_found" });
    return;
  }
  res.json(buildProtectedResourceMetadata("m365", alias));
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
