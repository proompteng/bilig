#!/usr/bin/env bun

import net from "node:net";
const composeFile = process.env["BILIG_E2E_COMPOSE_FILE"] ?? "compose.yaml";
const composeProject = process.env["BILIG_E2E_COMPOSE_PROJECT"] ?? "bilig-playwright";
const e2eWebPort = process.env["BILIG_E2E_WEB_PORT"] ?? "4180";
const e2eSyncServerPort = process.env["BILIG_E2E_SYNC_SERVER_PORT"] ?? "54422";
const e2eZeroPort = process.env["BILIG_E2E_ZERO_PORT"] ?? "54849";
const e2ePostgresPort = process.env["BILIG_E2E_POSTGRES_PORT"] ?? "55433";
const e2eBaseUrl = process.env["BILIG_E2E_BASE_URL"] ?? `http://127.0.0.1:${e2eWebPort}`;
const e2eSyncServerUrl =
  process.env["BILIG_E2E_SYNC_SERVER_URL"] ?? `http://127.0.0.1:${e2eSyncServerPort}`;

function dockerDaemonReady(): boolean {
  const result = Bun.spawnSync(["docker", "ps"], {
    stdin: "ignore",
    stdout: "ignore",
    stderr: "ignore",
    env: composeEnv(),
  });
  return result.exitCode === 0;
}

function resolveComposeInvocation(): {
  readonly command: readonly string[];
  readonly label: string;
} | null {
  const dockerCompose = Bun.spawnSync(["docker", "compose", "version"], {
    stdin: "ignore",
    stdout: "ignore",
    stderr: "ignore",
  });
  if (dockerCompose.exitCode === 0) {
    return { command: ["docker", "compose"], label: "docker compose" };
  }

  const dockerComposeStandalone = Bun.spawnSync(["docker-compose", "version"], {
    stdin: "ignore",
    stdout: "ignore",
    stderr: "ignore",
  });
  if (dockerComposeStandalone.exitCode === 0) {
    return { command: ["docker-compose"], label: "docker-compose" };
  }

  return null;
}

async function ensureDockerDaemon(): Promise<void> {
  if (dockerDaemonReady()) {
    return;
  }

  if (process.platform === "darwin" && Bun.which("open")) {
    Bun.spawnSync(["open", "-a", "Docker"], {
      stdin: "ignore",
      stdout: "ignore",
      stderr: "ignore",
    });
  }

  await waitForDockerDaemon(Date.now() + 120_000);
}

async function waitForDockerDaemon(deadline: number): Promise<void> {
  if (dockerDaemonReady()) {
    return;
  }
  if (Date.now() > deadline) {
    throw new Error("Docker daemon is unavailable for Playwright browser tests.");
  }
  await Bun.sleep(2_000);
  await waitForDockerDaemon(deadline);
}

function composeEnv(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    BILIG_E2E_WEB_PORT: e2eWebPort,
    BILIG_E2E_SYNC_SERVER_PORT: e2eSyncServerPort,
    BILIG_E2E_ZERO_PORT: e2eZeroPort,
    BILIG_E2E_POSTGRES_PORT: e2ePostgresPort,
  };
}

function runCompose(args: readonly string[], options?: { readonly allowFailure?: boolean }): void {
  const invocation = resolveComposeInvocation();
  if (!invocation) {
    throw new Error("Docker Compose is required for Playwright browser tests.");
  }

  const result = Bun.spawnSync(
    [...invocation.command, "-f", composeFile, "-p", composeProject, ...args],
    {
      stdin: "inherit",
      stdout: "inherit",
      stderr: "inherit",
      env: composeEnv(),
    },
  );

  if (!options?.allowFailure && result.exitCode !== 0) {
    throw new Error(
      `${invocation.label} ${args.join(" ")} failed with exit code ${result.exitCode ?? 1}`,
    );
  }
}

async function pollHttp(url: string, deadline: number, lastError?: string): Promise<void> {
  if (Date.now() > deadline) {
    throw new Error(`Timed out waiting for ${url}${lastError ? ` (${lastError})` : ""}`);
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

async function waitForHttp(url: string, timeoutMs = 240_000): Promise<void> {
  await pollHttp(url, Date.now() + timeoutMs);
}

async function pollTcp(
  host: string,
  port: number,
  deadline: number,
  lastError?: string,
): Promise<void> {
  if (Date.now() > deadline) {
    throw new Error(
      `Timed out waiting for tcp://${host}:${String(port)}${lastError ? ` (${lastError})` : ""}`,
    );
  }

  try {
    await new Promise<void>((resolve, reject) => {
      const socket = net.createConnection({ host, port });
      const cleanup = () => {
        socket.removeAllListeners();
        socket.destroy();
      };

      socket.once("connect", () => {
        cleanup();
        resolve();
      });
      socket.once("error", (error) => {
        cleanup();
        reject(error);
      });
    });
    return;
  } catch (error) {
    lastError = error instanceof Error ? error.message : String(error);
  }

  await Bun.sleep(250);
  await pollTcp(host, port, deadline, lastError);
}

async function waitForTcp(host: string, port: number, timeoutMs = 240_000): Promise<void> {
  await pollTcp(host, port, Date.now() + timeoutMs);
}

let stopping = false;
const keepAlive = setInterval(() => {}, 1 << 30);

async function shutdown(exitCode: number): Promise<never> {
  if (!stopping) {
    stopping = true;
    clearInterval(keepAlive);
    runCompose(["down", "-v", "--remove-orphans"], { allowFailure: true });
  }
  process.exit(exitCode);
}

process.on("SIGINT", () => {
  void shutdown(0);
});
process.on("SIGTERM", () => {
  void shutdown(0);
});

try {
  await ensureDockerDaemon();
  runCompose(["down", "-v", "--remove-orphans"], { allowFailure: true });
  runCompose(["up", "-d", "--build", "postgres", "bilig-app", "zero-cache"]);
  await waitForHttp(`${e2eBaseUrl}/healthz`);
  await waitForHttp(`${e2eSyncServerUrl}/healthz`);
  await waitForTcp("127.0.0.1", Number.parseInt(e2eZeroPort, 10));
  console.log(`Playwright E2E stack ready on ${e2eBaseUrl}`);
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  await shutdown(1);
}
