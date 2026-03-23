import { createSyncServer } from "./server.js";
import { createHttpWorksheetExecutor } from "./worksheet-executor.js";

const host = process.env["HOST"] ?? "0.0.0.0";
const port = Number.parseInt(process.env["PORT"] ?? "4321", 10);
const localServerUrl =
  process.env["LOCAL_SERVER_URL"] ?? process.env["BILIG_LOCAL_SERVER_URL"] ?? "";

const { app } = createSyncServer({
  worksheetExecutor: localServerUrl
    ? createHttpWorksheetExecutor({ baseUrl: localServerUrl })
    : null,
});

app
  .listen({ host, port })
  .then(() => {
    app.log.info({ host, port }, "bilig sync server listening");
    return undefined;
  })
  .catch((error) => {
    app.log.error(error, "bilig sync server failed to start");
    process.exitCode = 1;
    return undefined;
  });
