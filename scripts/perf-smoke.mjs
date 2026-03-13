import { runEditBenchmark } from "../packages/benchmarks/src/benchmark-edit.ts";

const { elapsedMs: elapsed, metrics, downstreamCount } = await runEditBenchmark(1_000);

if (elapsed > 250) {
  console.error(`perf smoke exceeded threshold: ${elapsed.toFixed(2)}ms`);
  process.exit(1);
}

if (metrics.dirtyFormulaCount < downstreamCount) {
  console.error(
    `perf smoke failed to mark the expected downstream formulas dirty: expected at least ${downstreamCount}, got ${metrics.dirtyFormulaCount}`
  );
  process.exit(1);
}

if (metrics.wasmFormulaCount === 0) {
  console.error("perf smoke did not exercise the wasm fast path");
  process.exit(1);
}

console.log(JSON.stringify({ elapsedMs: elapsed, downstreamCount, metrics }, null, 2));
