import { attachStdioAgentLoop } from "./stdio-handler.js";
import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";
import { createLocalServer } from "./server.js";
import { createHttpSyncRelay } from "./sync-relay.js";

const port = Number.parseInt(process.env["PORT"] ?? "4381", 10);
const host = process.env["HOST"] ?? "127.0.0.1";
const syncServerUrl = process.env["SYNC_SERVER_URL"] ?? process.env["BILIG_SYNC_SERVER_URL"] ?? "";
const stdioMode = process.env["BILIG_AGENT_STDIO"] === "1";

const sessionManager = new LocalWorkbookSessionManager(
  syncServerUrl
    ? {
        createSyncRelay: (documentId) => createHttpSyncRelay({
          documentId,
          baseUrl: syncServerUrl
        })
      }
    : {}
);
const { app } = createLocalServer({
  sessionManager,
  logger: !stdioMode
});
const stdioLoop = stdioMode
  ? attachStdioAgentLoop({
      handler: sessionManager
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
