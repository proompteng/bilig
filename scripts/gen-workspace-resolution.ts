import { readFileSync, writeFileSync } from "node:fs";
import {
  createTsconfigPaths,
  scanWorkspaceResolution,
  workspaceResolutionJsonPath,
  workspaceResolutionTsconfigPath,
} from "./workspace-resolution.ts";

const checkOnly = process.argv.includes("--check");

function formatJson(value: unknown): string {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function readFileIfExists(path: string): string | null {
  try {
    return readFileSync(path, "utf8");
  } catch {
    return null;
  }
}

const resolution = scanWorkspaceResolution();
const resolutionJson = formatJson(resolution);
const tsconfigJson = formatJson({
  extends: "./tsconfig.base.json",
  compilerOptions: {
    paths: createTsconfigPaths(resolution),
  },
});

if (checkOnly) {
  const failures: string[] = [];
  if (readFileIfExists(workspaceResolutionJsonPath) !== resolutionJson) {
    failures.push(workspaceResolutionJsonPath);
  }
  if (readFileIfExists(workspaceResolutionTsconfigPath) !== tsconfigJson) {
    failures.push(workspaceResolutionTsconfigPath);
  }
  if (failures.length > 0) {
    throw new Error(
      `Workspace resolution artifacts are out of date:\n- ${failures.join("\n- ")}\nRun bun scripts/gen-workspace-resolution.ts.`,
    );
  }
} else {
  writeFileSync(workspaceResolutionJsonPath, resolutionJson);
  writeFileSync(workspaceResolutionTsconfigPath, tsconfigJson);
}
