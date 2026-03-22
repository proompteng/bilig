#!/usr/bin/env bun

import { existsSync, readdirSync, rmSync } from "node:fs";
import { join, resolve } from "node:path";

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const packagesDir = join(rootDir, "packages");

for (const entry of readdirSync(packagesDir, { withFileTypes: true })) {
  if (!entry.isDirectory()) {
    continue;
  }

  const packageDir = join(packagesDir, entry.name);
  for (const relativePath of ["dist", "tsconfig.tsbuildinfo"]) {
    const target = join(packageDir, relativePath);
    if (existsSync(target)) {
      rmSync(target, { recursive: true, force: true });
    }
  }
}
