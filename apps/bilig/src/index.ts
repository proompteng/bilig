import { createSyncServer } from "./http/sync-server.js";
import { createLocalServer } from "./http/local-server.js";
import { DocumentSessionManager } from "./workbook-runtime/document-session-manager.js";
import { SyncDocumentSupervisor } from "./workbook-runtime/sync-document-supervisor.js";
import { createHttpWorksheetExecutor } from "./workbook-runtime/worksheet-executor.js";
import { createZeroSyncService } from "./zero/service.js";

async function main() {
  const host = process.env["HOST"] ?? "0.0.0.0";
  const appPort = Number.parseInt(process.env["PORT"] ?? "4321", 10);
  const localPort = Number.parseInt(process.env["LOCAL_PORT"] ?? "4381", 10);
  const localServerUrl =
    process.env["LOCAL_SERVER_URL"] ??
    process.env["BILIG_LOCAL_SERVER_URL"] ??
    `http://127.0.0.1:${localPort}`;
  const publicServerUrl = process.env["BILIG_PUBLIC_SERVER_URL"] ?? `http://127.0.0.1:${appPort}`;
  const browserAppBaseUrl = process.env["BILIG_WEB_APP_BASE_URL"] ?? publicServerUrl;
  const maxImportBytes = Number.parseInt(process.env["BILIG_AGENT_IMPORT_MAX_BYTES"] ?? "", 10);

  const sessionManager = new DocumentSessionManager(
    undefined,
    undefined,
    createHttpWorksheetExecutor({ baseUrl: localServerUrl }),
    {
      publicServerUrl,
      browserAppBaseUrl,
      ...(Number.isFinite(maxImportBytes) ? { maxImportBytes } : {}),
    },
  );
  const documentService = new SyncDocumentSupervisor(sessionManager);
  const zeroSyncService = createZeroSyncService();

  await zeroSyncService.initialize();

  const { app: syncApp } = createSyncServer({
    sessionManager,
    documentService,
    zeroSyncService,
  });

  const { app: localApp } = createLocalServer();

  try {
    await Promise.all([
      syncApp.listen({ host, port: appPort }),
      localApp.listen({ host, port: localPort }),
    ]);
    syncApp.log.info(
      { host, appPort, localPort, zeroSync: zeroSyncService.enabled },
      "bilig app listening",
    );
  } catch (error) {
    await zeroSyncService.close().catch(() => undefined);
    console.error(error);
    process.exit(1);
  }
}

main().catch(console.error);
