#!/usr/bin/env bun

function parsePort(value: string | undefined): number {
  const parsed = Number.parseInt(value ?? "", 10);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error("Expected a numeric preview port.");
  }
  return parsed;
}

const port = parsePort(process.argv[2]);
const host = process.argv[3] ?? "127.0.0.1";

const child = Bun.spawn(
  ["pnpm", "exec", "vite", "preview", "--host", host, "--port", String(port), "--strictPort"],
  {
    cwd: process.cwd(),
    stdin: "ignore",
    stdout: "inherit",
    stderr: "inherit",
  },
);

function forwardAndExit(signal: NodeJS.Signals): void {
  try {
    child.kill(signal);
  } catch {}
}

process.on("SIGINT", () => forwardAndExit("SIGINT"));
process.on("SIGTERM", () => forwardAndExit("SIGTERM"));

process.exit(await child.exited);
