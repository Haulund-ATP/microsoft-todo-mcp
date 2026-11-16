import { Router } from "express";
import { z } from "zod";
import { ConnectionsRepo } from "../storage/connectionsRepo.js";
import { ProfilesRepo, isValidAlias } from "../storage/profilesRepo.js";
import { requireOwner } from "./requireOwner.js";
import { requireCsrf, issueCsrfToken } from "./csrf.js";
import { clearOwnerSessionCookie } from "./session.js";
import { audit } from "../logging/audit.js";
import { renderAccountsPage } from "./view.js";
import { loadEnv } from "../config/env.js";

export const accountsRouter = Router();

accountsRouter.get("/accounts", requireOwner, async (req, res) => {
  const [connections, profiles] = await Promise.all([new ConnectionsRepo().list(), new ProfilesRepo().list()]);
  const csrfToken = issueCsrfToken(req, res);
  res.set("Content-Type", "text/html; charset=utf-8");
  res.send(renderAccountsPage({ connections, profiles, csrfToken, publicBaseUrl: loadEnv().PUBLIC_BASE_URL }));
});

/**
 * Copy buttons for the connector URLs on /accounts. Served as a same-origin
 * file because the CSP (script-src 'self') blocks inline scripts. Static and
 * contains no data, so it needs no owner session.
 */
export const COPY_SCRIPT = `document.addEventListener("click", async (event) => {
  const button = event.target instanceof Element ? event.target.closest("button[data-copy]") : null;
  if (!button) return;
  const text = button.getAttribute("data-copy") || "";
  const field = button.parentElement && button.parentElement.querySelector("input.copy-url");
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    if (field) { field.focus(); field.select(); document.execCommand("copy"); }
  }
  const label = button.textContent;
  button.textContent = "Copied";
  setTimeout(() => { button.textContent = label; }, 1500);
});
document.addEventListener("focusin", (event) => {
  if (event.target instanceof HTMLInputElement && event.target.classList.contains("copy-url")) event.target.select();
});
`;

accountsRouter.get("/accounts/copy.js", (_req, res) => {
  res.set("Content-Type", "application/javascript; charset=utf-8");
  res.set("Cache-Control", "public, max-age=3600");
  res.send(COPY_SCRIPT);
});

accountsRouter.post("/accounts/logout", requireOwner, requireCsrf, (_req, res) => {
  clearOwnerSessionCookie(res);
  res.redirect("/accounts");
});

const bindSchema = z.object({
  profile_alias: z.string().min(1).max(50),
  connection_id: z.string().min(1),
  /** Applies to /m365/{alias} only; /mcp/{alias} is always To Do without send. */
  allow_send: z.literal("on").optional(),
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
  const connection = await new ConnectionsRepo().get(parsed.data.connection_id);
  if (!connection) {
    res.status(400).send("Unknown connection.");
    return;
  }
  const repo = new ProfilesRepo();
  await repo.bind(parsed.data.profile_alias, parsed.data.connection_id, {
    allowSend: parsed.data.allow_send === "on",
  });
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
