import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "../packages/core/src/index.ts";

const engine = new SpreadsheetEngine({ workbookName: "perf-smoke" });
engine.createSheet("Sheet1");

for (let i = 1; i <= 1000; i += 1) {
  const addr = `A${i}`;
  engine.setCellValue("Sheet1", addr, i);
}

for (let i = 1; i <= 1000; i += 1) {
  const addr = `B${i}`;
  engine.setCellFormula("Sheet1", addr, `A${i}*2`);
}

const started = performance.now();
engine.setCellValue("Sheet1", "A1", 42);
const metrics = engine.getLastMetrics();
const elapsed = performance.now() - started;

if (elapsed > 250) {
  console.error(`perf smoke exceeded threshold: ${elapsed.toFixed(2)}ms`);
  process.exit(1);
}

console.log(JSON.stringify({ elapsedMs: elapsed, metrics }, null, 2));
