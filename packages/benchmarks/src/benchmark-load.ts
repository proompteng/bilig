import { performance } from "node:perf_hooks";
import { SpreadsheetEngine } from "@bilig/core";
import { seedWorkbook } from "./generate-workbook.js";

const engine = new SpreadsheetEngine({ workbookName: "benchmark-load" });
await engine.ready();

const started = performance.now();
seedWorkbook(engine, 5000);
const elapsed = performance.now() - started;

console.log(JSON.stringify({ scenario: "load", elapsedMs: elapsed }, null, 2));
