#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import {
  extractKnownLimitations,
  readClassSurface,
  readInterfaceKeys,
  type HyperFormulaSurfaceSnapshot,
} from "./workpaper-surface-contract.js";

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const hyperFormulaRoot = resolve(
  process.env.HYPERFORMULA_REPO_DIR ?? "/Users/gregkonush/github.com/hyperformula",
);
const outputPath = join(
  rootDir,
  "packages",
  "headless",
  "src",
  "__tests__",
  "fixtures",
  "hyperformula-surface.json",
);
const isCheckMode = process.argv.includes("--check");

if (!existsSync(hyperFormulaRoot)) {
  throw new Error(
    `HyperFormula checkout not found at ${hyperFormulaRoot}. Set HYPERFORMULA_REPO_DIR or clone the repo locally.`,
  );
}

const packageJson = JSON.parse(
  readFileSync(join(hyperFormulaRoot, "package.json"), "utf8"),
) as unknown;
if (typeof packageJson !== "object" || packageJson === null) {
  throw new Error(
    `Unable to parse HyperFormula package.json at ${join(hyperFormulaRoot, "package.json")}`,
  );
}
const hyperFormulaVersion = Reflect.get(packageJson, "version");
if (typeof hyperFormulaVersion !== "string" || hyperFormulaVersion.length === 0) {
  throw new Error(
    `Unable to read HyperFormula version from ${join(hyperFormulaRoot, "package.json")}`,
  );
}

const commitResult = Bun.spawnSync(["git", "-C", hyperFormulaRoot, "rev-parse", "HEAD"], {
  stdin: "ignore",
  stdout: "pipe",
  stderr: "pipe",
});
if (commitResult.exitCode !== 0) {
  throw new Error(
    `Unable to read HyperFormula commit: ${new TextDecoder().decode(commitResult.stderr).trim()}`,
  );
}

const knownLimitationsMarkdown = readFileSync(
  join(hyperFormulaRoot, "docs", "guide", "known-limitations.md"),
  "utf8",
);

const snapshot: HyperFormulaSurfaceSnapshot = {
  hyperFormulaRoot,
  hyperFormulaVersion,
  hyperFormulaCommit: new TextDecoder().decode(commitResult.stdout).trim(),
  knownLimitations: extractKnownLimitations(knownLimitationsMarkdown),
  classSurface: readClassSurface(join(hyperFormulaRoot, "src", "HyperFormula.ts"), "HyperFormula"),
  configKeys: readInterfaceKeys(join(hyperFormulaRoot, "src", "ConfigParams.ts"), "ConfigParams"),
};

const serializedSnapshot = `${JSON.stringify(snapshot, null, 2)}\n`;

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(`Missing generated snapshot at ${outputPath}`);
  }
  const currentSnapshot = readFileSync(outputPath, "utf8");
  if (currentSnapshot !== serializedSnapshot) {
    throw new Error(
      `Generated WorkPaper HyperFormula audit is out of date. Run: bun scripts/gen-workpaper-hyperformula-audit.ts`,
    );
  }
} else {
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, serializedSnapshot);
}

console.log(
  JSON.stringify(
    {
      outputPath,
      hyperFormulaVersion: snapshot.hyperFormulaVersion,
      hyperFormulaCommit: snapshot.hyperFormulaCommit,
      instanceMethodCount: snapshot.classSurface.instanceMethods.length,
      configKeyCount: snapshot.configKeys.length,
      mode: isCheckMode ? "check" : "write",
    },
    null,
    2,
  ),
);
