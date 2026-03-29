import { Effect } from "effect";

import { attachStdioAgentLoop } from "./stdio-handler.js";
import { LocalDocumentSupervisor } from "./document-supervisor.js";
import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";
import { createLocalServer } from "./server.js";
import { createHttpSyncRelay } from "./sync-relay.js";

const port = Number.parseInt(process.env["PORT"] ?? "4381", 10);
const host = process.env["HOST"] ?? "127.0.0.1";
const syncServerUrl = process.env["SYNC_SERVER_URL"] ?? process.env["BILIG_SYNC_SERVER_URL"] ?? "";
const stdioMode = process.env["BILIG_AGENT_STDIO"] === "1";
const publicServerUrl = process.env["BILIG_PUBLIC_SERVER_URL"] ?? "";
const browserAppBaseUrl = process.env["BILIG_WEB_APP_BASE_URL"] ?? "";
const maxImportBytes = Number.parseInt(process.env["BILIG_AGENT_IMPORT_MAX_BYTES"] ?? "", 10);

const sharedManagerOptions = {
  ...(publicServerUrl ? { publicServerUrl } : {}),
  ...(browserAppBaseUrl ? { browserAppBaseUrl } : {}),
  ...(Number.isFinite(maxImportBytes) ? { maxImportBytes } : {}),
};

const sessionManager = new LocalWorkbookSessionManager(
  syncServerUrl
    ? {
        createSyncRelay: (documentId) =>
          createHttpSyncRelay({
            documentId,
            baseUrl: syncServerUrl,
          }),
        ...sharedManagerOptions,
      }
    : sharedManagerOptions,
);
const documentService = new LocalDocumentSupervisor(sessionManager);
const { app } = createLocalServer({
  sessionManager,
  documentService,
  logger: !stdioMode,
});
const stdioLoop = stdioMode
  ? attachStdioAgentLoop({
      handler: {
        handleAgentFrame(frame) {
          return documentService.handleAgentFrame(frame).pipe(Effect.runPromise);
        },
      },
    })
  : null;

app.listen({ port, host }).catch((error: unknown) => {
  if (stdioMode) {
    console.error(error);
  } else {
    app.log.error(error);
  }
  process.exitCode = 1;
});

process.on("exit", () => {
  stdioLoop?.dispose();
});
