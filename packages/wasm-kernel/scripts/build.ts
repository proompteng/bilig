#!/usr/bin/env bun

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const exec = promisify(execFile);
const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const asc = resolve(rootDir, "../../node_modules/.bin/asc");

await exec(asc, ["assembly/index.ts", "--target", "release"], { cwd: rootDir });
