import { createSyncServer } from "./server.js";

const host = process.env["HOST"] ?? "0.0.0.0";
const port = Number.parseInt(process.env["PORT"] ?? "4321", 10);

const { app } = createSyncServer();

app.listen({ host, port })
  .then(() => {
    app.log.info({ host, port }, "bilig sync server listening");
    return undefined;
  })
  .catch((error) => {
    app.log.error(error, "bilig sync server failed to start");
    process.exitCode = 1;
    return undefined;
  });
