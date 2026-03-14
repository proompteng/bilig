import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import type { RecalcMetrics } from "@bilig/protocol";
import { seedDownstreamWorkbook } from "./generate-workbook.js";
import { measureMemory, sampleMemory, type MemoryMeasurement } from "./metrics.js";

export interface EditBenchmarkResult {
  scenario: "single-edit";
  downstreamCount: number;
  elapsedMs: number;
  metrics: RecalcMetrics;
  memory: MemoryMeasurement;
}

export async function runEditBenchmark(downstreamCount = 10_000): Promise<EditBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: "benchmark-edit" });
  await engine.ready();
  seedDownstreamWorkbook(engine, downstreamCount);

  const memoryBefore = sampleMemory();
  const started = performance.now();
  engine.setCellValue("Sheet1", "A1", 99);
  const elapsed = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "single-edit",
    downstreamCount,
    elapsedMs: elapsed,
    metrics: engine.getLastMetrics(),
    memory: measureMemory(memoryBefore, memoryAfter)
  };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const downstreamCount = Number.parseInt(process.argv[2] ?? "10000", 10);
  console.log(JSON.stringify(await runEditBenchmark(downstreamCount), null, 2));
}
