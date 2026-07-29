import { Router } from "express";
import { z } from "zod";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { ProfilesRepo, isValidAlias } from "../storage/profilesRepo.js";
import { requireOwner } from "./requireOwner.js";
import { requireCsrf, issueCsrfToken } from "./csrf.js";
import { clearOwnerSessionCookie } from "./session.js";
import { audit } from "../logging/audit.js";
import { renderAccountsPage } from "./view.js";

export const accountsRouter = Router();

accountsRouter.get("/accounts", requireOwner, async (req, res) => {
  const [connections, profiles] = await Promise.all([new ConnectionsRepo().list(), new ProfilesRepo().list()]);
  const csrfToken = issueCsrfToken(res);
  res.set("Content-Type", "text/html; charset=utf-8");
  res.send(renderAccountsPage({ connections, profiles, csrfToken }));
});

accountsRouter.post("/accounts/logout", requireOwner, requireCsrf, (_req, res) => {
  clearOwnerSessionCookie(res);
  res.redirect("/accounts");
});

const bindSchema = z.object({
  profile_alias: z.string().min(1).max(50),
  connection_id: z.string().min(1),
});

accountsRouter.post("/accounts/profiles/bind", requireOwner, requireCsrf, async (req, res) => {
  const parsed = bindSchema.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).send("Invalid profile alias or connection.");
    return;
  }
  if (!isValidAlias(parsed.data.profile_alias)) {
    res.status(400).send("That alias is reserved or invalid. Use lowercase letters, digits, and hyphens.");
    return;
  }
  const repo = new ProfilesRepo();
  await repo.bind(parsed.data.profile_alias, parsed.data.connection_id);
  audit({
    action: "profile.bound",
    correlationId: req.correlationId,
    profileAlias: parsed.data.profile_alias,
    connectionId: parsed.data.connection_id,
    outcome: "success",
  });
  res.redirect("/accounts");
});

accountsRouter.post("/accounts/profiles/unbind/:alias", requireOwner, requireCsrf, async (req, res) => {
  const alias = req.params.alias;
  if (!alias) {
    res.status(400).send("Missing alias.");
    return;
  }
  const repo = new ProfilesRepo();
  await repo.unbind(alias);
  audit({ action: "profile.unbound", correlationId: req.correlationId, profileAlias: alias, outcome: "success" });
  res.redirect("/accounts");
});

const renameSchema = z.object({ alias: z.string().min(1).max(50) });

accountsRouter.post("/accounts/connections/:connectionId/alias", requireOwner, requireCsrf, async (req, res) => {
  const connectionId = req.params.connectionId;
  const parsed = renameSchema.safeParse(req.body);
  if (!connectionId || !parsed.success || !/^[a-z0-9-]{1,50}$/.test(parsed.data.alias)) {
    res.status(400).send("Alias must be lowercase letters, digits, and hyphens.");
    return;
  }
  const repo = new ConnectionsRepo();
  await repo.update(connectionId, { alias: parsed.data.alias });
  audit({
    action: "connection.alias_changed",
    correlationId: req.correlationId,
    connectionId,
    outcome: "success",
  });
  res.redirect("/accounts");
});
