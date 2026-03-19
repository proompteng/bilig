import { execFileSync, spawnSync } from "node:child_process";

const PREVIEW_PORTS = [4179, 4180];

function getListeningPids(port) {
  try {
    const output = execFileSync("lsof", ["-tiTCP:" + String(port), "-sTCP:LISTEN"], {
      encoding: "utf8",
      stdio: ["ignore", "pipe", "ignore"]
    }).trim();
    if (!output) {
      return [];
    }
    return output
      .split(/\s+/)
      .map((value) => Number.parseInt(value, 10))
      .filter((value) => Number.isInteger(value));
  } catch {
    return [];
  }
}

function sleep(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function terminatePreviewServers() {
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

function runPlaywright(args) {
  const result = spawnSync("pnpm", ["exec", "playwright", "test", ...args], {
    stdio: "inherit",
    env: process.env
  });
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

terminatePreviewServers();
runPlaywright([]);
terminatePreviewServers();
runPlaywright(["-c", "playwright.web.config.ts"]);
terminatePreviewServers();
