import { createSyncServer } from "./http/sync-server.js";
import { DocumentSessionManager } from "./workbook-runtime/document-session-manager.js";
import { SyncDocumentSupervisor } from "./workbook-runtime/sync-document-supervisor.js";
import { LocalDocumentSupervisor } from "./workbook-runtime/local-document-supervisor.js";
import { LocalWorkbookSessionManager } from "./workbook-runtime/local-workbook-session-manager.js";
import { createInProcessWorksheetExecutor } from "./workbook-runtime/worksheet-executor.js";
import { createZeroSyncService } from "./zero/service.js";
import { createWorkbookAgentService } from "./codex-app/workbook-agent-service.js";

async function main() {
  const host = process.env["HOST"] ?? "0.0.0.0";
  const appPort = Number.parseInt(process.env["PORT"] ?? "4321", 10);
  const publicServerUrl = process.env["BILIG_PUBLIC_SERVER_URL"] ?? `http://127.0.0.1:${appPort}`;
  const browserAppBaseUrl = process.env["BILIG_WEB_APP_BASE_URL"] ?? publicServerUrl;
  const maxImportBytes = Number.parseInt(process.env["BILIG_AGENT_IMPORT_MAX_BYTES"] ?? "", 10);
  const worksheetHostSessionManager = new LocalWorkbookSessionManager({
    publicServerUrl,
    browserAppBaseUrl,
    ...(Number.isFinite(maxImportBytes) ? { maxImportBytes } : {}),
  });
  const worksheetHostDocumentService = new LocalDocumentSupervisor(worksheetHostSessionManager);

  const sessionManager = new DocumentSessionManager(
    undefined,
    undefined,
    createInProcessWorksheetExecutor({
      documentService: worksheetHostDocumentService,
      serverUrl: publicServerUrl,
      browserAppBaseUrl,
    }),
    {
      publicServerUrl,
      browserAppBaseUrl,
      ...(Number.isFinite(maxImportBytes) ? { maxImportBytes } : {}),
    },
  );
  const documentService = new SyncDocumentSupervisor(sessionManager);
  const zeroSyncService = createZeroSyncService();
  const workbookAgentService = createWorkbookAgentService(zeroSyncService);

  await zeroSyncService.initialize();

  const { app: syncApp } = createSyncServer({
    sessionManager,
    documentService,
    zeroSyncService,
    workbookAgentService,
  });

  try {
    await syncApp.listen({ host, port: appPort });
    syncApp.log.info({ host, appPort, zeroSync: zeroSyncService.enabled }, "bilig app listening");
  } catch (error) {
    await workbookAgentService.close().catch(() => undefined);
    await zeroSyncService.close().catch(() => undefined);
    console.error(error);
    process.exit(1);
  }
}

main().catch(console.error);
