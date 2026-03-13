import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import { seedWorkbook } from "./generate-workbook.js";

const engine = new SpreadsheetEngine({ workbookName: "benchmark-edit" });
await engine.ready();
seedWorkbook(engine, 5000);

const started = performance.now();
engine.setCellValue("Sheet1", "A1", 99);
const elapsed = performance.now() - started;

console.log(
  JSON.stringify(
    {
      scenario: "single-edit",
      elapsedMs: elapsed,
      metrics: engine.getLastMetrics()
    },
    null,
    2
  )
);
