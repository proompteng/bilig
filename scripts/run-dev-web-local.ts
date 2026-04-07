#!/usr/bin/env bun

import net from "node:net";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const composeFiles = ["compose.yaml", "compose.dev-local.yaml"] as const;
const composeProject = process.env["BILIG_DEV_COMPOSE_PROJECT"] ?? "bilig-dev-local";
const postgresService = "postgres";
const zeroCacheService = "zero-cache-local";
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const webAppDir = resolve(repoRoot, "apps/web");
const preferredAppPort = Number.parseInt(
  process.env["PORT"] ?? process.env["BILIG_SYNC_SERVER_PORT"] ?? "4321",
  10,
);
const postgresPort = process.env["BILIG_DEV_POSTGRES_PORT"] ?? "55432";
const preferredWebPort = Number.parseInt(process.env["BILIG_WEB_DEV_PORT"] ?? "5173", 10);
const configuredZeroProxyUpstream = process.env["BILIG_ZERO_PROXY_UPSTREAM"];
const preferredZeroPort = Number.parseInt(
  process.env["BILIG_DEV_ZERO_PORT"] ??
    (configuredZeroProxyUpstream ? new URL(configuredZeroProxyUpstream).port : "4848") ??
    "4848",
  10,
);
const zeroProxyUpstream =
  configuredZeroProxyUpstream ?? `http://127.0.0.1:${String(preferredZeroPort)}`;
const zeroHealthUrl = `${zeroProxyUpstream}/keepalive`;
const cleanupCompose = process.env["BILIG_DEV_CLEANUP_COMPOSE"] === "true";
let resolvedAppPort = String(preferredAppPort);

interface DevChildProcess {
  readonly exited: Promise<number | null>;
  kill(signal?: "SIGINT" | "SIGTERM" | number): void;
}

function composeEnv(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    BILIG_E2E_POSTGRES_PORT: postgresPort,
    BILIG_DEV_APP_PORT: resolvedAppPort,
    BILIG_DEV_ZERO_PORT: String(preferredZeroPort),
  };
}

function resolveComposeInvocation(): {
  readonly command: readonly string[];
  readonly label: string;
} | null {
  const dockerCompose = Bun.spawnSync(["docker", "compose", "version"], {
    stdin: "ignore",
    stdout: "ignore",
    stderr: "ignore",
    env: composeEnv(),
  });
  if (dockerCompose.exitCode === 0) {
    return { command: ["docker", "compose"], label: "docker compose" };
  }

  const dockerComposeStandalone = Bun.spawnSync(["docker-compose", "version"], {
    stdin: "ignore",
    stdout: "ignore",
    stderr: "ignore",
    env: composeEnv(),
  });
  if (dockerComposeStandalone.exitCode === 0) {
    return { command: ["docker-compose"], label: "docker-compose" };
  }

  return null;
}

function composeArgs(args: readonly string[]): readonly string[] {
  const invocation = resolveComposeInvocation();
  if (!invocation) {
    throw new Error("Docker Compose is required for pnpm dev:web-local.");
  }
  return [
    ...invocation.command,
    ...composeFiles.flatMap((file) => ["-f", file]),
    "-p",
    composeProject,
    ...args,
  ];
}

function runComposeSync(
  args: readonly string[],
  options?: { readonly allowFailure?: boolean },
): void {
  const invocation = resolveComposeInvocation();
  if (!invocation) {
    throw new Error("Docker Compose is required for pnpm dev:web-local.");
  }

  const result = Bun.spawnSync(composeArgs(args), {
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    env: composeEnv(),
  });

  if (!options?.allowFailure && result.exitCode !== 0) {
    throw new Error(
      `${invocation.label} ${args.join(" ")} failed with exit code ${result.exitCode ?? 1}`,
    );
  }
}

async function waitForHttp(url: string, timeoutMs = 120_000): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  const poll = async (lastError: string): Promise<void> => {
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for ${url} (${lastError})`);
    }
    try {
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
      await Bun.sleep(250);
      return poll(`HTTP ${response.status}`);
    } catch (error) {
      await Bun.sleep(250);
      return poll(error instanceof Error ? error.message : String(error));
    }
  };
  return poll("not started");
}

async function ensureDockerCompose(): Promise<void> {
  if (resolveComposeInvocation()) {
    return;
  }
  throw new Error("Docker Compose is required for pnpm dev:web-local.");
}

function listListeningPids(port: number): string[] {
  const result = Bun.spawnSync(["lsof", "-tiTCP:" + String(port), "-sTCP:LISTEN"], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "ignore",
  });
  if (result.exitCode !== 0) {
    return [];
  }
  return new TextDecoder()
    .decode(result.stdout)
    .split("\n")
    .map((value) => value.trim())
    .filter((value) => value.length > 0);
}

function commandForPid(pid: string): string {
  const result = Bun.spawnSync(["ps", "-p", pid, "-o", "command="], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "ignore",
  });
  if (result.exitCode !== 0) {
    return "";
  }
  return new TextDecoder().decode(result.stdout).trim();
}

function cwdForPid(pid: string): string {
  const result = Bun.spawnSync(["lsof", "-a", "-p", pid, "-d", "cwd", "-Fn"], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "ignore",
  });
  if (result.exitCode !== 0) {
    return "";
  }
  return (
    new TextDecoder()
      .decode(result.stdout)
      .split("\n")
      .find((line) => line.startsWith("n"))
      ?.slice(1)
      .trim() ?? ""
  );
}

function isRepoOwnedListener(pid: string): boolean {
  const command = commandForPid(pid);
  if (command.includes(repoRoot)) {
    return true;
  }
  const cwd = cwdForPid(pid);
  return cwd === repoRoot || cwd.startsWith(`${repoRoot}/`);
}

async function reapStaleRepoListeners(ports: readonly number[]): Promise<void> {
  const repoOwnedPids = [
    ...new Set(ports.flatMap((port) => listListeningPids(port)).filter(isRepoOwnedListener)),
  ];
  for (const pid of repoOwnedPids) {
    try {
      process.kill(Number.parseInt(pid, 10), "SIGTERM");
    } catch {}
  }

  const deadline = Date.now() + 5_000;
  const waitUntilCleared = async (): Promise<void> => {
    if (Date.now() > deadline) {
      return;
    }
    const occupied = ports.some((port) => listListeningPids(port).some(isRepoOwnedListener));
    if (!occupied) {
      return;
    }
    await Bun.sleep(100);
    return waitUntilCleared();
  };
  await waitUntilCleared();
}

async function resolveWebPort(preferredPort: number): Promise<number> {
  const explicitPort = process.env["BILIG_WEB_DEV_PORT"];
  if (explicitPort) {
    if (!(await isPortAvailable(preferredPort))) {
      throw new Error(`Web port ${preferredPort} is already in use.`);
    }
    return preferredPort;
  }
  return findAvailablePort(preferredPort, 10, `web port starting at ${preferredPort}`);
}

async function resolveAppPort(preferredPort: number): Promise<number> {
  const explicitPort = process.env["PORT"] ?? process.env["BILIG_SYNC_SERVER_PORT"];
  if (explicitPort) {
    if (!(await isPortAvailable(preferredPort))) {
      throw new Error(`App port ${preferredPort} is already in use.`);
    }
    return preferredPort;
  }
  return findAvailablePort(preferredPort, 10, `app port starting at ${preferredPort}`);
}

async function findAvailablePort(
  startPort: number,
  remainingOffsets: number,
  label: string,
  offset = 0,
): Promise<number> {
  if (offset >= remainingOffsets) {
    throw new Error(`Unable to find an available ${label}.`);
  }
  const candidate = startPort + offset;
  if (await isPortAvailable(candidate)) {
    return candidate;
  }
  return findAvailablePort(startPort, remainingOffsets, label, offset + 1);
}

async function isPortAvailable(port: number): Promise<boolean> {
  return await new Promise((resolvePortAvailability) => {
    const server = net.createServer();

    server.once("error", () => resolvePortAvailability(false));
    server.once("listening", () => {
      server.close(() => resolvePortAvailability(true));
    });

    server.listen(port, "127.0.0.1");
  });
}

function spawnAppDev(
  appPort: string,
  postgresUrl: string,
  publicServerUrl: string,
  webAppBaseUrl: string,
): DevChildProcess {
  return Bun.spawn(["pnpm", "--filter", "@bilig/app", "run", "dev"], {
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    env: {
      ...process.env,
      HOST: process.env["HOST"] ?? "0.0.0.0",
      PORT: appPort,
      DATABASE_URL: postgresUrl,
      BILIG_ZERO_PROXY_UPSTREAM: zeroProxyUpstream,
      BILIG_ZERO_CACHE_URL: "/zero",
      BILIG_PUBLIC_SERVER_URL: publicServerUrl,
      BILIG_WEB_APP_BASE_URL: webAppBaseUrl,
      BILIG_CORS_ORIGIN: webAppBaseUrl,
    },
  });
}

function spawnWebDev(webPort: number, publicServerUrl: string): DevChildProcess {
  return Bun.spawn(
    [
      "node",
      "../../node_modules/vite/bin/vite.js",
      "--host",
      "0.0.0.0",
      "--port",
      String(webPort),
      "--strictPort",
    ],
    {
      cwd: webAppDir,
      stdin: "inherit",
      stdout: "inherit",
      stderr: "inherit",
      env: {
        ...process.env,
        BILIG_SYNC_SERVER_PORT: new URL(publicServerUrl).port,
        BILIG_SYNC_SERVER_TARGET: publicServerUrl,
      },
    },
  );
}

function killIfRunning(process: DevChildProcess | null | undefined, signal: NodeJS.Signals): void {
  if (!process) {
    return;
  }
  try {
    process.kill(signal);
  } catch {}
}

function cleanupComposeStack(): void {
  if (!cleanupCompose) {
    return;
  }
  runComposeSync(["down", "-v", "--remove-orphans"], { allowFailure: true });
}

console.log("Starting local postgres dependency...");
await ensureDockerCompose();
cleanupComposeStack();
await reapStaleRepoListeners(
  Array.from({ length: 10 }, (_, index) => preferredWebPort + index).concat(preferredAppPort),
);
const appPort = await resolveAppPort(preferredAppPort);
resolvedAppPort = String(appPort);
const postgresUrl =
  process.env["DATABASE_URL"] ?? `postgresql://bilig:bilig@127.0.0.1:${postgresPort}/bilig`;
const publicServerUrl = process.env["BILIG_PUBLIC_SERVER_URL"] ?? `http://127.0.0.1:${appPort}`;
const appHealthUrl = `${publicServerUrl}/healthz`;
runComposeSync(["up", "-d", "--wait", postgresService]);
const webPort = await resolveWebPort(preferredWebPort);
const webAppBaseUrl = process.env["BILIG_WEB_APP_BASE_URL"] ?? `http://localhost:${webPort}`;

console.log(`Starting local app dev server (app=${publicServerUrl})...`);
const appChild = spawnAppDev(String(appPort), postgresUrl, publicServerUrl, webAppBaseUrl);
let webChild: DevChildProcess | null = null;

let shuttingDown = false;

function forwardSignal(signal: NodeJS.Signals): void {
  if (shuttingDown) {
    return;
  }
  shuttingDown = true;
  killIfRunning(appChild, signal);
  killIfRunning(webChild, signal);
  cleanupComposeStack();
}

process.on("SIGINT", () => forwardSignal("SIGINT"));
process.on("SIGTERM", () => forwardSignal("SIGTERM"));

try {
  await waitForHttp(appHealthUrl);
  console.log(`Starting local web dev server (web=${webAppBaseUrl})...`);
  webChild = spawnWebDev(webPort, publicServerUrl);
  console.log("App is healthy, starting local zero-cache...");
  runComposeSync(["up", "-d", zeroCacheService]);
  await waitForHttp(zeroHealthUrl);
  await waitForHttp(webAppBaseUrl);
  console.log(
    `Local dev stack ready: web=${webAppBaseUrl} app=${publicServerUrl} zero=${zeroProxyUpstream}`,
  );
} catch (error) {
  forwardSignal("SIGTERM");
  throw error;
}

const exitCode = await Promise.race([
  appChild.exited.then((code) => ({ code, source: "@bilig/app" })),
  (webChild?.exited ?? new Promise<never>(() => undefined)).then((code) => ({
    code,
    source: "@bilig/web",
  })),
]);

forwardSignal("SIGTERM");

if (exitCode.code !== 0) {
  console.error(`${exitCode.source} exited with code ${exitCode.code ?? 1}`);
}

process.exit(exitCode.code ?? 0);
