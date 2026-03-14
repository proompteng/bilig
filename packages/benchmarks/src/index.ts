import { runEditBenchmark } from "./benchmark-edit.js";
import { runLoadBenchmark } from "./benchmark-load.js";
import { runRenderCommitBenchmark } from "./benchmark-renderer.js";

export * from "./benchmark-edit.js";
export * from "./benchmark-load.js";
export * from "./benchmark-renderer.js";
export * from "./generate-workbook.js";
export * from "./metrics.js";
export * from "./stats.js";

if (import.meta.url === `file://${process.argv[1]}`) {
  const results = [];
  results.push(await runLoadBenchmark(10_000));
  results.push(await runLoadBenchmark(50_000));
  results.push(await runLoadBenchmark(100_000));
  results.push(await runEditBenchmark(100));
  results.push(await runEditBenchmark(1_000));
  results.push(await runEditBenchmark(10_000));
  results.push(await runRenderCommitBenchmark(1_000));
  results.push(await runRenderCommitBenchmark(10_000));

  console.log(JSON.stringify(results, null, 2));
}
