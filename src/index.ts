import { loadEnv } from "./config/env.js";
import { logger } from "./logging/logger.js";
import { ensureStorageProvisioned } from "./storage/clients.js";
import { createApp } from "./http/app.js";

async function main(): Promise<void> {
  const env = loadEnv();

  await ensureStorageProvisioned();

  const app = createApp();
  const server = app.listen(env.PORT, () => {
    logger.info({ deploymentVersion: process.env.npm_package_version ?? "0.1.0" }, `microsoft-todo-mcp listening on :${env.PORT}`);
  });

  const shutdown = (signal: string) => {
    logger.info({ detail: signal }, "shutting down");
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(1), 10_000).unref();
  };
  process.on("SIGTERM", () => shutdown("SIGTERM"));
  process.on("SIGINT", () => shutdown("SIGINT"));
}

main().catch((err) => {
  logger.error({ err: err instanceof Error ? err.message : String(err) }, "fatal startup error");
  process.exit(1);
});
