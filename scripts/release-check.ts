#!/usr/bin/env bun

import { gzipSync } from "node:zlib";
import { readdir } from "node:fs/promises";
import { basename, resolve } from "node:path";

const budgets = {
  mainJsGzipBytes: 350 * 1024,
  runtimeWasmGzipBytes: 250 * 1024,
  sqliteWasmGzipBytes: 400 * 1024,
};

async function findAssets() {
  const assetDir = resolve("apps/web/dist/assets");
  const entries = await readdir(assetDir, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile())
    .map((entry) => resolve(assetDir, entry.name));

  const jsAssets = files.filter((file) => file.endsWith(".js"));
  const wasmAssets = files.filter((file) => file.endsWith(".wasm"));

  if (jsAssets.length === 0) {
    throw new Error("No built JavaScript assets were found in apps/web/dist/assets");
  }

  if (wasmAssets.length === 0) {
    throw new Error("No built WASM assets were found in apps/web/dist/assets");
  }

  return { jsAssets, wasmAssets };
}

async function measureAsset(file) {
  const bytes = new Uint8Array(await Bun.file(file).arrayBuffer());
  return {
    file,
    rawBytes: bytes.byteLength,
    gzipBytes: gzipSync(bytes).byteLength,
  };
}

function assertBudget(label, actual, budget) {
  if (actual > budget) {
    throw new Error(`${label} exceeded budget: ${actual} bytes > ${budget} bytes`);
  }
}

function isSqliteWasmAsset(file) {
  return basename(file).startsWith("sqlite3-");
}

const { jsAssets, wasmAssets } = await findAssets();
const jsMeasurements = await Promise.all(jsAssets.map((file) => measureAsset(file)));
const wasmMeasurements = await Promise.all(wasmAssets.map((file) => measureAsset(file)));
const largestJs = jsMeasurements.reduce((largest, entry) =>
  entry.gzipBytes > largest.gzipBytes ? entry : largest,
);
const runtimeWasmMeasurements = wasmMeasurements.filter((entry) => !isSqliteWasmAsset(entry.file));
const sqliteWasmMeasurements = wasmMeasurements.filter((entry) => isSqliteWasmAsset(entry.file));

if (runtimeWasmMeasurements.length === 0) {
  throw new Error("No first-party runtime WASM assets were found in apps/web/dist/assets");
}

if (sqliteWasmMeasurements.length === 0) {
  throw new Error("No SQLite WASM assets were found in apps/web/dist/assets");
}

const largestRuntimeWasm = runtimeWasmMeasurements.reduce((largest, entry) =>
  entry.gzipBytes > largest.gzipBytes ? entry : largest,
);
const largestSqliteWasm = sqliteWasmMeasurements.reduce((largest, entry) =>
  entry.gzipBytes > largest.gzipBytes ? entry : largest,
);

assertBudget("Main JavaScript gzip size", largestJs.gzipBytes, budgets.mainJsGzipBytes);
assertBudget("Runtime WASM gzip size", largestRuntimeWasm.gzipBytes, budgets.runtimeWasmGzipBytes);
assertBudget("SQLite WASM gzip size", largestSqliteWasm.gzipBytes, budgets.sqliteWasmGzipBytes);

console.log(
  JSON.stringify(
    {
      budgets,
      largestJs,
      largestRuntimeWasm,
      largestSqliteWasm,
    },
    null,
    2,
  ),
);
