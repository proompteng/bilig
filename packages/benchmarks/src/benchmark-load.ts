import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import { seedLoadWorkbook } from "./generate-workbook.js";

export interface LoadBenchmarkResult {
  scenario: "load";
  materializedCells: number;
  elapsedMs: number;
}

export async function runLoadBenchmark(materializedCells = 10_000): Promise<LoadBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: "benchmark-load" });
  await engine.ready();

  const started = performance.now();
  seedLoadWorkbook(engine, materializedCells);
  const elapsed = performance.now() - started;

  return { scenario: "load", materializedCells, elapsedMs: elapsed };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const materializedCells = Number.parseInt(process.argv[2] ?? "10000", 10);
  console.log(JSON.stringify(await runLoadBenchmark(materializedCells), null, 2));
}
