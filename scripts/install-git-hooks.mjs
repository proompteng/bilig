import { execFileSync } from "node:child_process";
import { existsSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import process from "node:process";

const require = createRequire(import.meta.url);
const { setHooksFromConfig, skipInstall } = require("simple-git-hooks");

function runGit(args) {
  return execFileSync("git", args, {
    cwd: process.cwd(),
    encoding: "utf8",
  }).trim();
}

function getLocalHooksPath() {
  try {
    const hooksPath = runGit(["config", "--local", "--get", "core.hooksPath"]);
    return hooksPath || null;
  } catch {
    return null;
  }
}

function isWorktreeCheckout() {
  if (!existsSync(".git")) {
    return false;
  }

  return !statSync(".git").isDirectory();
}

try {
  runGit(["rev-parse", "--is-inside-work-tree"]);
} catch {
  console.info("[INFO] No git repository found, skipping hook install.");
  process.exit(0);
}

if (skipInstall()) {
  process.exit(0);
}

if (isWorktreeCheckout() && !getLocalHooksPath()) {
  const hooksPath = runGit(["rev-parse", "--git-path", "hooks"]);
  execFileSync("git", ["config", "--local", "core.hooksPath", hooksPath], {
    cwd: process.cwd(),
    stdio: "inherit",
  });
}

await setHooksFromConfig(process.cwd());
