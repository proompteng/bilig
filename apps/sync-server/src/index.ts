import { createSyncServer } from "./server.js";
import { DocumentSessionManager } from "./document-session-manager.js";
import { createHttpWorksheetExecutor } from "./worksheet-executor.js";
import { createZeroSyncService } from "./zero/service.js";

const host = process.env["HOST"] ?? "0.0.0.0";
const port = Number.parseInt(process.env["PORT"] ?? "4321", 10);
const localServerUrl =
  process.env["LOCAL_SERVER_URL"] ?? process.env["BILIG_LOCAL_SERVER_URL"] ?? "";
const publicServerUrl = process.env["BILIG_PUBLIC_SERVER_URL"] ?? "";
const browserAppBaseUrl = process.env["BILIG_WEB_APP_BASE_URL"] ?? "";
const maxImportBytes = Number.parseInt(process.env["BILIG_AGENT_IMPORT_MAX_BYTES"] ?? "", 10);

const sessionManager = new DocumentSessionManager(
  undefined,
  undefined,
  localServerUrl ? createHttpWorksheetExecutor({ baseUrl: localServerUrl }) : null,
  {
    ...(publicServerUrl ? { publicServerUrl } : {}),
    ...(browserAppBaseUrl ? { browserAppBaseUrl } : {}),
    ...(Number.isFinite(maxImportBytes) ? { maxImportBytes } : {}),
  },
);

const zeroSyncService = createZeroSyncService();

void zeroSyncService
  .initialize()
  .then(async () => {
    const { app } = createSyncServer({
      sessionManager,
      zeroSyncService,
    });

    try {
      await app.listen({ host, port });
      app.log.info(
        { host, port, zeroSync: zeroSyncService.enabled },
        "bilig sync server listening",
      );
    } catch (error) {
      app.log.error(error, "bilig sync server failed to start");
      process.exitCode = 1;
    }
    return undefined;
  })
  .catch((error) => {
    console.error("Failed to initialize Zero sync", error);
    process.exitCode = 1;
  });
