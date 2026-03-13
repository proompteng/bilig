import { runEditBenchmark } from "../packages/benchmarks/src/benchmark-edit.ts";
import { runLoadBenchmark } from "../packages/benchmarks/src/benchmark-load.ts";
import { runRenderCommitBenchmark } from "../packages/benchmarks/src/benchmark-renderer.ts";

const baseBudgets = {
  load100kMs: 1500,
  edit10kElapsedMs: 120,
  edit10kRecalcMs: 50,
  renderCommit10kMs: 50
};
const toleranceMultiplier = Number.parseFloat(process.env.BILIG_BENCH_TOLERANCE ?? (process.env.CI ? "1.5" : "1"));
const budgets = Object.fromEntries(
  Object.entries(baseBudgets).map(([key, value]) => [key, value * toleranceMultiplier])
);

function assertBudget(label, actual, threshold) {
  if (actual > threshold) {
    throw new Error(`${label} exceeded budget: ${actual.toFixed(2)}ms > ${threshold}ms`);
  }
}

const load = await runLoadBenchmark(100_000);
const edit = await runEditBenchmark(10_000);
const renderCommit = await runRenderCommitBenchmark(10_000);

assertBudget("100k snapshot load", load.elapsedMs, budgets.load100kMs);
assertBudget("10k downstream edit", edit.elapsedMs, budgets.edit10kElapsedMs);
assertBudget("10k downstream recalc", edit.metrics.recalcMs, budgets.edit10kRecalcMs);
assertBudget("10k render commit", renderCommit.elapsedMs, budgets.renderCommit10kMs);

console.log(
  JSON.stringify(
    {
      baseBudgets,
      budgets,
      toleranceMultiplier,
      results: {
        load,
        edit,
        renderCommit
      }
    },
    null,
    2
  )
);
