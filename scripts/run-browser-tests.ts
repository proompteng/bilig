#!/usr/bin/env bun

const textDecoder = new TextDecoder();

const PREVIEW_PORTS = [4179, 4180];

function commandExists(command: string): boolean {
  return Bun.which(command) !== null;
}

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
    env: process.env,
  });
  if (result.exitCode !== 0) {
    process.exit(result.exitCode ?? 1);
  }
}

terminatePreviewServers();
runPlaywright([]);
terminatePreviewServers();
