import { Router } from "express";
import { getTableServiceClient } from "../storage/clients.js";

/**
 * /health, /ready, /version. None of these leak secrets, connection
 * strings, tenant ids, or account data — /health and /version are
 * unauthenticated by design (used by Container Apps probes and uptime
 * checks) so they are held to a strict "no sensitive data" bar.
 */
export const healthRouter = Router();

healthRouter.get("/health", (_req, res) => {
  res.status(200).json({ status: "ok" });
});

healthRouter.get("/ready", async (_req, res) => {
  try {
    // A cheap call that proves storage connectivity without touching any
    // tenant-specific or secret data.
    const client = getTableServiceClient();
    await client.listTables().next();
    res.status(200).json({ status: "ready" });
  } catch {
    res.status(503).json({ status: "not_ready" });
  }
});

healthRouter.get("/version", (_req, res) => {
  res.status(200).json({
    name: "microsoft-todo-mcp",
    version: process.env.npm_package_version ?? "0.1.0",
    commit: process.env.GIT_COMMIT_SHA ?? "unknown",
    nodeVersion: process.version,
  });
});
