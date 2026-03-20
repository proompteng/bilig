import { createLocalServer } from "./server.js";

const port = Number.parseInt(process.env["PORT"] ?? "4381", 10);
const host = process.env["HOST"] ?? "127.0.0.1";

const { app } = createLocalServer();

app.listen({ port, host }).catch((error: unknown) => {
  app.log.error(error);
  process.exitCode = 1;
});
