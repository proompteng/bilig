#!/usr/bin/env bun

import { existsSync, readFileSync } from "node:fs";
import { createServer } from "node:net";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const textDecoder = new TextDecoder();
const playwrightArgs = process.argv.slice(2);
const workspaceRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const requestedBrowserStack = process.env["BILIG_BROWSER_STACK"] ?? "auto";
const normalizedBrowserStack =
  requestedBrowserStack === "compose" ||
  requestedBrowserStack === "local" ||
  requestedBrowserStack === "auto"
    ? requestedBrowserStack
    : "local";
const isCi = process.env["CI"] === "1" || process.env["CI"] === "true";
if (requestedBrowserStack !== normalizedBrowserStack) {
  console.warn(
    `Unknown BILIG_BROWSER_STACK "${requestedBrowserStack}", defaulting to "local" stack.`,
  );
}
type ComposeInvocation = {
  label: string;
  command: string[];
  version: string;
};

let composeInvocation: ComposeInvocation | null = null;
let composeInvocationProbed = false;
let composeInvocationLogged = false;

function commandExists(command: string): boolean {
  return Bun.which(command) !== null;
}

function probeComposeInvocation(): ComposeInvocation | null {
  const dockerComposeCandidates = [
    { label: "docker compose", command: ["docker", "compose"] },
    { label: "docker-compose", command: ["docker-compose"] },
  ];

  for (const candidate of dockerComposeCandidates) {
    if (!commandExists(candidate.command[0])) {
      continue;
    }

    const result = Bun.spawnSync([...candidate.command, "version"], {
      stdin: "ignore",
      stdout: "pipe",
      stderr: "pipe",
    });
    if (result.exitCode !== 0) {
      continue;
    }

    const version = [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)]
      .join("")
      .trim();
    return {
      label: candidate.label,
      command: candidate.command,
      version,
    };
  }

  return null;
}

function resolveComposeInvocation(): ComposeInvocation | null {
  if (!composeInvocationProbed) {
    composeInvocation = probeComposeInvocation();
    composeInvocationProbed = true;
  }

  return composeInvocation;
}

function requireComposeInvocation(required: boolean): ComposeInvocation | null {
  const invocation = resolveComposeInvocation();

  if (!invocation && required) {
    throw new Error(
      "docker compose is required for BILIG_BROWSER_STACK=compose, but neither `docker compose` nor `docker-compose` is available.",
    );
  }

  if (invocation && !composeInvocationLogged) {
    const version = invocation.version ? ` (${invocation.version})` : "";
    console.info(`compose is available via "${invocation.label}"${version}.`);
    composeInvocationLogged = true;
  }

  return invocation;
}

const compose = resolveComposeInvocation();
const composeLabel = compose ? compose.label : "unavailable";
const browserStack =
  (normalizedBrowserStack === "compose" || normalizedBrowserStack === "auto") && compose
    ? "compose"
    : "local";

if ((normalizedBrowserStack === "compose" || normalizedBrowserStack === "auto") && compose) {
  console.info(`BILIG_BROWSER_STACK=compose requested; using compose command "${composeLabel}"`);
}

if (normalizedBrowserStack === "compose" && !compose && isCi) {
  throw new Error(
    "BILIG_BROWSER_STACK=compose is required in CI, but neither `docker compose` nor `docker-compose` is available.",
  );
}

if (normalizedBrowserStack === "auto" && !compose) {
  if (isCi) {
    throw new Error(
      "CI requires docker compose for browser tests, but `docker compose` and `docker-compose` are both unavailable. Set BILIG_BROWSER_STACK=local only when a local runtime provides /v2/session.",
    );
  }

  const fallbackCommand = "`docker compose` or `docker-compose`";
  console.warn(
    `BILIG_BROWSER_STACK is auto and compose is unavailable; falling back to local Playwright server for browser tests (requested compose command: ${fallbackCommand})`,
  );
}

if (normalizedBrowserStack === "compose" && !compose && !isCi) {
  const fallbackCommand = "`docker compose` or `docker-compose`";
  console.warn(
    `compose unavailable in this environment, falling back to local Playwright server for browser tests (requested compose command: ${fallbackCommand})`,
  );
}

if (browserStack === "local") {
  console.info(
    "BILIG_BROWSER_STACK=local selected; Playwright will run against local web stack and expects /v2/session to be available.",
  );
}

function isContainerizedRuntime(): boolean {
  return existsSync("/.dockerenv") || existsSync("/run/.containerenv");
}

function resolveCurrentContainerId(): string | null {
  const hostname = process.env["HOSTNAME"]?.trim();
  if (hostname) {
    return hostname;
  }

  try {
    const value = readFileSync("/etc/hostname", "utf8").trim();
    return value.length > 0 ? value : null;
  } catch {}

  return null;
}

function shouldUseComposeInternalNetwork(): boolean {
  return browserStack === "compose" && isContainerizedRuntime();
}

const composeFile = resolve(workspaceRoot, process.env["BILIG_E2E_COMPOSE_FILE"] ?? "compose.yaml");
const composeProjectDirectory = dirname(composeFile);
const composeProject = process.env["BILIG_E2E_COMPOSE_PROJECT"] ?? `bilig-e2e-${Date.now()}`;
const e2eWebPort = process.env["BILIG_E2E_WEB_PORT"] ?? "4180";
const e2eSyncServerPort = process.env["BILIG_E2E_SYNC_SERVER_PORT"] ?? "54422";
const e2eZeroPort = process.env["BILIG_E2E_ZERO_PORT"] ?? "54849";
const e2ePostgresPort = process.env["BILIG_E2E_POSTGRES_PORT"] ?? "55433";
const e2eBaseUrl =
  process.env["BILIG_E2E_BASE_URL"] ??
  `http://127.0.0.1:${e2eWebPort}`;
const e2eSyncServerUrl =
  process.env["BILIG_E2E_SYNC_SERVER_URL"] ??
  (shouldUseComposeInternalNetwork()
    ? "http://sync-server:4321"
    : `http://127.0.0.1:${e2eSyncServerPort}`);

const PREVIEW_PORTS = [4179, 4180];
const SESSION_BOOTSTRAP_TIMEOUT_MS = 30_000;

function parsePidList(output: string): number[] {
  if (!output) {
    return [];
  }
  return output
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value));
}

function parseSsPids(output: string): number[] {
  const matches = output.matchAll(/pid=(\d+)/g);
  const pids: number[] = [];
  for (const match of matches) {
    const pid = Number.parseInt(match[1] ?? "", 10);
    if (Number.isInteger(pid)) {
      pids.push(pid);
    }
  }
  return pids;
}

function getListeningPids(port: number): number[] {
  if (commandExists("lsof")) {
    const result = Bun.spawnSync(["lsof", "-tiTCP:" + String(port), "-sTCP:LISTEN"], {
      stdin: "ignore",
      stdout: "pipe",
      stderr: "ignore",
    });
    if (result.exitCode !== 0) {
      return [];
    }
    return parsePidList(textDecoder.decode(result.stdout).trim());
  }

  if (commandExists("ss")) {
    const result = Bun.spawnSync(["ss", "-ltnp", "sport = :" + String(port)], {
      stdin: "ignore",
      stdout: "pipe",
      stderr: "ignore",
    });
    if (result.exitCode !== 0) {
      return [];
    }
    return parseSsPids(textDecoder.decode(result.stdout));
  }

  if (commandExists("netstat")) {
    const result = Bun.spawnSync(["netstat", "-ltnp"], {
      stdin: "ignore",
      stdout: "pipe",
      stderr: "ignore",
    });
    if (result.exitCode !== 0) {
      return [];
    }
    const lines = textDecoder
      .decode(result.stdout)
      .split("\n")
      .filter((line) => line.includes(":" + String(port)) && line.includes("LISTEN"));
    const pids: number[] = [];
    for (const line of lines) {
      const fields = line.trim().split(/\s+/);
      const program = fields.at(-1) ?? "";
      const pid = Number.parseInt(program.split("/", 1)[0] ?? "", 10);
      if (Number.isInteger(pid)) {
        pids.push(pid);
      }
    }
    return pids;
  }

  if (commandExists("fuser")) {
    const result = Bun.spawnSync(["fuser", "-n", "tcp", String(port)], {
      stdin: "ignore",
      stdout: "pipe",
      stderr: "ignore",
    });
    if (result.exitCode !== 0) {
      return [];
    }
    return parsePidList(textDecoder.decode(result.stdout).trim());
  }

  console.warn(
    `No port-inspection command available; skipping preview server cleanup for port ${port}.`,
  );
  return [];
}

function sleep(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function terminatePreviewServers(): void {
  const pids = Array.from(new Set(PREVIEW_PORTS.flatMap((port) => getListeningPids(port))));
  if (pids.length === 0) {
    return;
  }

  for (const pid of pids) {
    try {
      process.kill(pid, "SIGTERM");
    } catch {}
  }

  sleep(300);

  for (const pid of pids) {
    try {
      process.kill(pid, 0);
      process.kill(pid, "SIGKILL");
    } catch {}
  }
}

function runPlaywright(args: string[]): void {
  const result = Bun.spawnSync(["pnpm", "exec", "playwright", "test", ...args], {
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    cwd: workspaceRoot,
    env: {
      ...process.env,
      BILIG_BROWSER_STACK: browserStack,
      BILIG_E2E_BASE_URL: e2eBaseUrl,
    },
  });
  if (result.exitCode !== 0) {
    process.exit(result.exitCode ?? 1);
  }
}

async function pollHttp(url: string, deadline: number, lastError = "unknown error"): Promise<void> {
  if (Date.now() >= deadline) {
    throw new Error(`Timed out waiting for ${url}: ${lastError}`);
  }
  try {
    const response = await fetch(url);
    if (response.ok) {
      return;
    }
    lastError = `HTTP ${response.status}`;
  } catch (error) {
    lastError = error instanceof Error ? error.message : String(error);
  }
  await Bun.sleep(250);
  await pollHttp(url, deadline, lastError);
}

async function waitForHttp(
  url: string,
  timeoutMs = 120_000,
  lastError = "unknown error",
): Promise<void> {
  await pollHttp(url, Date.now() + timeoutMs, lastError);
}

async function ensureRuntimeSessionEndpoint(baseUrl: string): Promise<void> {
  const sessionUrl = `${baseUrl.replace(/\/$/, "")}/v2/session`;
  await waitForHttp(
    sessionUrl,
    SESSION_BOOTSTRAP_TIMEOUT_MS,
    `runtime session endpoint ${sessionUrl} did not become available`,
  );
}

function buildComposeCommand(invocation: ComposeInvocation, args: string[]): string[] {
  return [
    ...invocation.command,
    "--project-directory",
    composeProjectDirectory,
    "-f",
    composeFile,
    "-p",
    composeProject,
    ...args,
  ];
}

async function startLocalTcpProxy(
  listenPort: number,
  targetHost: string,
  targetPort: number,
): Promise<() => Promise<void>> {
  const server = createServer((incomingSocket) => {
    const outgoingSocket = Bun.connect({
      hostname: targetHost,
      port: targetPort,
      socket: {
        data(socket, data) {
          incomingSocket.write(Buffer.from(data));
          return socket;
        },
        close() {
          incomingSocket.end();
        },
        error() {
          incomingSocket.destroy();
        },
        open(socket) {
          incomingSocket.on("data", (chunk) => {
            socket.write(chunk);
          });
          incomingSocket.on("end", () => {
            socket.end();
          });
          incomingSocket.on("error", () => {
            socket.terminate();
          });
        },
      },
    });

    incomingSocket.on("close", () => {
      outgoingSocket.terminate();
    });
  });

  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(listenPort, "127.0.0.1", () => {
      server.off("error", reject);
      resolve();
    });
  });

  console.info(
    `compose browser stack is running inside a container; proxying 127.0.0.1:${listenPort} to ${targetHost}:${targetPort}.`,
  );

  return async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((error) => {
        if (error) {
          reject(error);
          return;
        }
        resolve();
      });
    });
  };
}

function connectCurrentContainerToComposeNetwork(): void {
  if (!shouldUseComposeInternalNetwork()) {
    return;
  }
  if (!commandExists("docker")) {
    throw new Error(
      "compose browser stack is running inside a container, but the docker CLI is unavailable to attach the job container to the compose network.",
    );
  }

  const containerId = resolveCurrentContainerId();
  if (!containerId) {
    throw new Error(
      "compose browser stack is running inside a container, but the current container id could not be determined.",
    );
  }

  const networkName = `${composeProject}_default`;
  const result = Bun.spawnSync(["docker", "network", "connect", networkName, containerId], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
    cwd: workspaceRoot,
  });
  if (result.exitCode === 0) {
    console.info(
      `compose browser stack is running inside a container; attached ${containerId} to ${networkName}.`,
    );
    return;
  }

  const output = [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)]
    .join("")
    .trim();
  if (output.includes("already exists") || output.includes("already connected")) {
    return;
  }

  throw new Error(
    `failed to attach the current container to compose network ${networkName}: ${output || "unknown docker network connect failure"}`,
  );
}

function disconnectCurrentContainerFromComposeNetwork(): void {
  if (!shouldUseComposeInternalNetwork() || !commandExists("docker")) {
    return;
  }

  const containerId = resolveCurrentContainerId();
  if (!containerId) {
    return;
  }

  const networkName = `${composeProject}_default`;
  const result = Bun.spawnSync(["docker", "network", "disconnect", networkName, containerId], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
    cwd: workspaceRoot,
  });
  if (result.exitCode === 0) {
    return;
  }

  const output = [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)]
    .join("")
    .trim();
  if (output.includes("is not connected") || output.includes("No such container")) {
    return;
  }

  console.warn(
    `failed to detach the current container from compose network ${networkName}: ${output || "unknown docker network disconnect failure"}`,
  );
}

function runDockerCompose(args: string[], env = process.env): void {
  const invocation = requireComposeInvocation(true);
  if (!invocation) {
    throw new Error("compose command is unavailable; cannot run compose stack.");
  }

  const result = Bun.spawnSync(buildComposeCommand(invocation, args), {
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    cwd: workspaceRoot,
    env: {
      ...env,
      BILIG_E2E_WEB_PORT: e2eWebPort,
      BILIG_E2E_SYNC_SERVER_PORT: e2eSyncServerPort,
      BILIG_E2E_ZERO_PORT: e2eZeroPort,
      BILIG_E2E_POSTGRES_PORT: e2ePostgresPort,
    },
  });
  if (result.exitCode !== 0) {
    throw new Error(
      `${invocation.label} ${args.join(" ")} failed with exit code ${result.exitCode ?? 1}`,
    );
  }
}

function collectComposeLogs(): string {
  const invocation = requireComposeInvocation(false);
  if (!invocation) {
    return "compose command is unavailable; compose logs were not collected.";
  }

  const result = Bun.spawnSync(buildComposeCommand(invocation, ["logs", "--no-color"]), {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
    cwd: workspaceRoot,
  });
  return [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)].join("").trim();
}

function collectComposeStatus(): string {
  const invocation = requireComposeInvocation(false);
  if (!invocation) {
    return "compose command is unavailable; compose service status was not collected.";
  }

  const result = Bun.spawnSync(buildComposeCommand(invocation, ["ps", "--all"]), {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe",
    cwd: workspaceRoot,
  });
  return [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)].join("").trim();
}

async function runComposePlaywright(): Promise<void> {
  requireComposeInvocation(true);

  terminatePreviewServers();
  let closeWebProxy: (() => Promise<void>) | null = null;
  try {
    runDockerCompose(["up", "-d", "--build", "postgres", "sync-server", "zero-cache", "web"]);
    connectCurrentContainerToComposeNetwork();
    if (shouldUseComposeInternalNetwork()) {
      closeWebProxy = await startLocalTcpProxy(Number.parseInt(e2eWebPort, 10), "web", 3000);
    }
    await waitForHttp(`${e2eBaseUrl}/healthz`);
    await waitForHttp(`${e2eSyncServerUrl}/healthz`);
    runPlaywright(playwrightArgs);
  } catch (error) {
    const logs = collectComposeLogs();
    const status = collectComposeStatus();
    const details = [logs, status]
      .map((value) => value.trim())
      .filter(Boolean)
      .join("\n\n");
    const context = details ? `\n\nCompose diagnostics:\n${details}` : "";
    throw new Error(`${error instanceof Error ? error.message : String(error)}${context}`, {
      cause: error,
    });
  } finally {
    if (closeWebProxy) {
      await closeWebProxy().catch((error) => {
        console.warn(
          `failed to stop the local compose web proxy: ${error instanceof Error ? error.message : String(error)}`,
        );
      });
    }
    disconnectCurrentContainerFromComposeNetwork();
    try {
      runDockerCompose(["down", "-v", "--remove-orphans"]);
    } catch (error) {
      const logs = collectComposeLogs();
      const status = collectComposeStatus();
      const details = [logs, status]
        .map((value) => value.trim())
        .filter(Boolean)
        .join("\n\n");
      const context = details ? `\n\nCompose diagnostics:\n${details}` : "";
      console.error(
        `compose shutdown failed: ${error instanceof Error ? error.message : String(error)}${context}`,
      );
    }
  }
}

if (browserStack === "compose") {
  await runComposePlaywright();
} else {
  terminatePreviewServers();
  await ensureRuntimeSessionEndpoint(e2eBaseUrl);
  runPlaywright(playwrightArgs);
  terminatePreviewServers();
}
