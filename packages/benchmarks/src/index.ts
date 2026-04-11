import { runEditBenchmark } from "./benchmark-edit.js";
import { runLoadBenchmark } from "./benchmark-load.js";
import { runRangeAggregateBenchmark } from "./benchmark-range-heavy.js";
import { runRenderCommitBenchmark } from "./benchmark-renderer.js";
import { runTopologyEditBenchmark } from "./benchmark-topology-edit.js";
import { runWorkPaperVsHyperFormulaExpandedBenchmarkSuite } from "./benchmark-workpaper-vs-hyperformula-expanded.js";
import { runWorkPaperBenchmarkSuite } from "./benchmark-workpaper.js";
import { runWorkPaperVsHyperFormulaBenchmarkSuite } from "./benchmark-workpaper-vs-hyperformula.js";

export * from "./benchmark-edit.js";
export * from "./benchmark-load.js";
export * from "./benchmark-range-heavy.js";
export * from "./benchmark-renderer.js";
export * from "./benchmark-topology-edit.js";
export * from "./benchmark-workpaper-vs-hyperformula-expanded.js";
export * from "./benchmark-workpaper.js";
export * from "./benchmark-workpaper-vs-hyperformula.js";
export * from "./generate-workbook.js";
export * from "./metrics.js";
export * from "./stats.js";
export * from "./workbook-corpus.js";
export * from "./workpaper-benchmark-fixtures.js";

if (import.meta.url === `file://${process.argv[1]}`) {
  const results = [];
  results.push(await runLoadBenchmark(10_000));
  results.push(await runLoadBenchmark(50_000));
  results.push(await runLoadBenchmark(100_000));
  results.push(await runLoadBenchmark("dense-mixed-250k"));
  results.push(await runEditBenchmark(100));
  results.push(await runEditBenchmark(1_000));
  results.push(await runEditBenchmark(10_000));
  results.push(await runRangeAggregateBenchmark(1_024, 10_000));
  results.push(await runTopologyEditBenchmark(10_000));
  results.push(await runRenderCommitBenchmark(1_000));
  results.push(await runRenderCommitBenchmark(10_000));
  results.push(...(await runWorkPaperBenchmarkSuite()));
  results.push(...runWorkPaperVsHyperFormulaBenchmarkSuite());
  results.push(...runWorkPaperVsHyperFormulaExpandedBenchmarkSuite());

  console.log(JSON.stringify(results, null, 2));
}
