#!/usr/bin/env bun

const textDecoder = new TextDecoder();

const PREVIEW_PORTS = [4179, 4180];

function getListeningPids(port) {
  const result = Bun.spawnSync(["lsof", "-tiTCP:" + String(port), "-sTCP:LISTEN"], {
    stdin: "ignore",
    stdout: "pipe",
    stderr: "ignore"
  });
  if (result.exitCode !== 0) {
    return [];
  }
  const output = textDecoder.decode(result.stdout).trim();
  if (!output) {
    return [];
  }
  return output
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value));
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
  const result = Bun.spawnSync(["pnpm", "exec", "playwright", "test", ...args], {
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    env: process.env
  });
  if (result.exitCode !== 0) {
    process.exit(result.exitCode ?? 1);
  }
}

terminatePreviewServers();
runPlaywright([]);
terminatePreviewServers();
runPlaywright(["-c", "playwright.web.config.ts"]);
terminatePreviewServers();
