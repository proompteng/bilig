import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import type { RecalcMetrics } from "@bilig/protocol";
import { seedRangeAggregateWorkbook } from "./generate-workbook.js";
import { measureMemory, sampleMemory, type MemoryMeasurement } from "./metrics.js";

export interface RangeAggregateBenchmarkResult {
  scenario: "range-aggregates";
  sourceCount: number;
  aggregateCount: number;
  elapsedMs: number;
  metrics: RecalcMetrics;
  memory: MemoryMeasurement;
}

export async function runRangeAggregateBenchmark(
  sourceCount = 1_024,
  aggregateCount = 10_000
): Promise<RangeAggregateBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: "benchmark-range-aggregates" });
  await engine.ready();
  seedRangeAggregateWorkbook(engine, sourceCount, aggregateCount);

  const memoryBefore = sampleMemory();
  const started = performance.now();
  engine.setCellValue("Sheet1", "A1", 99);
  const elapsed = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "range-aggregates",
    sourceCount,
    aggregateCount,
    elapsedMs: elapsed,
    metrics: engine.getLastMetrics(),
    memory: measureMemory(memoryBefore, memoryAfter)
  };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const sourceCount = Number.parseInt(process.argv[2] ?? "1024", 10);
  const aggregateCount = Number.parseInt(process.argv[3] ?? "10000", 10);
  console.log(JSON.stringify(await runRangeAggregateBenchmark(sourceCount, aggregateCount), null, 2));
}
