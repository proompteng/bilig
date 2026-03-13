import { runEditBenchmark } from "./benchmark-edit.js";
import { runLoadBenchmark } from "./benchmark-load.js";
import { runRenderCommitBenchmark } from "./benchmark-renderer.js";

export * from "./benchmark-edit.js";
export * from "./benchmark-load.js";
export * from "./benchmark-renderer.js";
export * from "./generate-workbook.js";

if (import.meta.url === `file://${process.argv[1]}`) {
  const results = await Promise.all([
    runLoadBenchmark(10_000),
    runLoadBenchmark(50_000),
    runLoadBenchmark(100_000),
    runEditBenchmark(100),
    runEditBenchmark(1_000),
    runEditBenchmark(10_000),
    runRenderCommitBenchmark(1_000),
    runRenderCommitBenchmark(10_000)
  ]);

  console.log(JSON.stringify(results, null, 2));
}
