#!/usr/bin/env bun
import { runEditBenchmark } from "../packages/benchmarks/src/benchmark-edit.ts";
import { runLoadBenchmark } from "../packages/benchmarks/src/benchmark-load.ts";
import { runRangeAggregateBenchmark } from "../packages/benchmarks/src/benchmark-range-heavy.ts";
import { runTopologyEditBenchmark } from "../packages/benchmarks/src/benchmark-topology-edit.ts";
import { runRenderCommitBenchmark } from "../packages/benchmarks/src/benchmark-renderer.ts";
import {
  runWorkerReconnectCatchUpBenchmark,
  runWorkerVisibleEditBenchmark,
  runWorkerWarmStartBenchmark,
} from "./bench-worker-runtime.ts";

const baseBudgets = {
  load100kP95Ms: 1500,
  load100kWorkingSetDeltaBytes: 250 * 1024 * 1024,
  edit10kElapsedP95Ms: 120,
  edit10kRecalcMedianMs: 50,
  edit10kRecalcP95Ms: 120,
  rangeAggregates10kElapsedP95Ms: 120,
  rangeAggregates10kRecalcP95Ms: 100,
  topologyEdit10kElapsedP95Ms: 80,
  topologyEdit10kRecalcP95Ms: 80,
  renderCommit10kP95Ms: 50,
  workerWarmStart100kP95Ms: 500,
  workerVisibleEdit10kP95Ms: 16,
  workerReconnectCatchUp100PendingP95Ms: 2000,
};
const toleranceMultiplier = Number.parseFloat(
  process.env.BILIG_BENCH_TOLERANCE ?? (process.env.CI ? "1.5" : "1"),
);
const budgets = Object.fromEntries(
  Object.entries(baseBudgets).map(([key, value]) => [key, value * toleranceMultiplier]),
);

function assertBudget(label, actual, threshold, formatter = formatMs) {
  if (actual > threshold) {
    throw new Error(`${label} exceeded budget: ${formatter(actual)} > ${formatter(threshold)}`);
  }
}

function assertAllRunsUseWasmFastPath(label, runs, expectedWasmFormulaCount) {
  const degradedRun = runs.find((run) => {
    return (
      run.metrics.wasmFormulaCount < expectedWasmFormulaCount || run.metrics.jsFormulaCount !== 0
    );
  });
  if (!degradedRun) {
    return;
  }
  throw new Error(
    `${label} did not stay on the wasm fast path: wasm=${degradedRun.metrics.wasmFormulaCount}, js=${degradedRun.metrics.jsFormulaCount}`,
  );
}

function formatMs(value) {
  return `${value.toFixed(2)}ms`;
}

function formatBytes(value) {
  return `${(value / (1024 * 1024)).toFixed(2)}MB`;
}

function summarizeNumbers(values) {
  if (values.length === 0) {
    throw new Error("Cannot summarize empty benchmark samples");
  }
  const samples = [...values].toSorted((left, right) => left - right);

  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length;
  return {
    samples,
    min: samples[0],
    median: quantile(samples, 0.5),
    p95: quantile(samples, 0.95),
    max: samples[samples.length - 1],
    mean,
  };
}

function quantile(sortedValues, percentile) {
  if (percentile <= 0) {
    return sortedValues[0];
  }
  if (percentile >= 1) {
    return sortedValues[sortedValues.length - 1];
  }
  const index = Math.ceil(percentile * sortedValues.length) - 1;
  return sortedValues[Math.max(0, Math.min(sortedValues.length - 1, index))];
}

async function sampleBenchmark(runner, iterations, { warmupIterations = 0 } = {}) {
  const run = () => runner();

  await Array.from({ length: warmupIterations }).reduce((previous) => {
    return previous.then(() => run());
  }, Promise.resolve());

  const runs = [];
  await Array.from({ length: iterations }).reduce((previous) => {
    return previous.then(() =>
      run().then((result) => {
        return runs.push(result);
      }),
    );
  }, Promise.resolve());
  return runs;
}

const loadRuns = await sampleBenchmark(() => runLoadBenchmark(100_000), 5, {
  warmupIterations: 1,
});
const editRuns = await sampleBenchmark(() => runEditBenchmark(10_000), 5, {
  warmupIterations: 1,
});
const rangeRuns = await sampleBenchmark(() => runRangeAggregateBenchmark(1_024, 10_000), 3, {
  warmupIterations: 1,
});
const topologyRuns = await sampleBenchmark(() => runTopologyEditBenchmark(10_000), 3, {
  warmupIterations: 1,
});
const renderRuns = await sampleBenchmark(() => runRenderCommitBenchmark(10_000), 5, {
  warmupIterations: 1,
});
const workerWarmStartRuns = await sampleBenchmark(() => runWorkerWarmStartBenchmark(100_000), 3, {
  warmupIterations: 1,
});
const workerVisibleEditRuns = await sampleBenchmark(
  () => runWorkerVisibleEditBenchmark(10_000),
  5,
  {
    warmupIterations: 1,
  },
);
const workerReconnectCatchUpRuns = await sampleBenchmark(
  () => runWorkerReconnectCatchUpBenchmark(10_000, 100),
  3,
  {
    warmupIterations: 1,
  },
);

const loadElapsed = summarizeNumbers(loadRuns.map((run) => run.elapsedMs));
const loadWorkingSetDelta = summarizeNumbers(
  loadRuns.map((run) => run.memory.delta.heapUsedBytes + run.memory.delta.externalBytes),
);
const loadHeapUsedAfter = summarizeNumbers(loadRuns.map((run) => run.memory.after.heapUsedBytes));

const editElapsed = summarizeNumbers(editRuns.map((run) => run.elapsedMs));
const editRecalc = summarizeNumbers(editRuns.map((run) => run.metrics.recalcMs));
const editRssAfter = summarizeNumbers(editRuns.map((run) => run.memory.after.rssBytes));

const rangeElapsed = summarizeNumbers(rangeRuns.map((run) => run.elapsedMs));
const rangeRecalc = summarizeNumbers(rangeRuns.map((run) => run.metrics.recalcMs));
const rangeRssAfter = summarizeNumbers(rangeRuns.map((run) => run.memory.after.rssBytes));

const topologyElapsed = summarizeNumbers(topologyRuns.map((run) => run.elapsedMs));
const topologyRecalc = summarizeNumbers(topologyRuns.map((run) => run.metrics.recalcMs));
const topologyRssAfter = summarizeNumbers(topologyRuns.map((run) => run.memory.after.rssBytes));

const renderElapsed = summarizeNumbers(renderRuns.map((run) => run.elapsedMs));
const renderRssAfter = summarizeNumbers(renderRuns.map((run) => run.memory.after.rssBytes));
const workerWarmStartElapsed = summarizeNumbers(workerWarmStartRuns.map((run) => run.elapsedMs));
const workerWarmStartWorkingSetDelta = summarizeNumbers(
  workerWarmStartRuns.map((run) => run.memory.delta.heapUsedBytes + run.memory.delta.externalBytes),
);
const workerVisibleEditElapsed = summarizeNumbers(
  workerVisibleEditRuns.map((run) => run.visiblePatchMs),
);
const workerVisibleEditCommitElapsed = summarizeNumbers(
  workerVisibleEditRuns.map((run) => run.commitMs),
);
const workerReconnectCatchUpElapsed = summarizeNumbers(
  workerReconnectCatchUpRuns.map((run) => run.catchUpMs),
);
const workerReconnectRebaseElapsed = summarizeNumbers(
  workerReconnectCatchUpRuns.map((run) => run.rebaseMs),
);
const workerReconnectSubmitDrainElapsed = summarizeNumbers(
  workerReconnectCatchUpRuns.map((run) => run.submitDrainMs),
);
const workerReconnectAckElapsed = summarizeNumbers(
  workerReconnectCatchUpRuns.map((run) => run.ackMs),
);

assertAllRunsUseWasmFastPath("10k range aggregate benchmark", rangeRuns, 10_000);

assertBudget("100k snapshot load p95", loadElapsed.p95, budgets.load100kP95Ms);
assertBudget(
  "100k snapshot load working-set delta",
  loadWorkingSetDelta.max,
  budgets.load100kWorkingSetDeltaBytes,
  formatBytes,
);
assertBudget("10k downstream edit p95", editElapsed.p95, budgets.edit10kElapsedP95Ms);
assertBudget("10k downstream recalc median", editRecalc.median, budgets.edit10kRecalcMedianMs);
assertBudget("10k downstream recalc p95", editRecalc.p95, budgets.edit10kRecalcP95Ms);
assertBudget(
  "10k range aggregate edit p95",
  rangeElapsed.p95,
  budgets.rangeAggregates10kElapsedP95Ms,
);
assertBudget(
  "10k range aggregate recalc p95",
  rangeRecalc.p95,
  budgets.rangeAggregates10kRecalcP95Ms,
);
assertBudget("10k topology edit p95", topologyElapsed.p95, budgets.topologyEdit10kElapsedP95Ms);
assertBudget("10k topology recalc p95", topologyRecalc.p95, budgets.topologyEdit10kRecalcP95Ms);
assertBudget("10k render commit p95", renderElapsed.p95, budgets.renderCommit10kP95Ms);
assertBudget(
  "100k worker warm-start p95",
  workerWarmStartElapsed.p95,
  budgets.workerWarmStart100kP95Ms,
);
assertBudget(
  "10k worker local visible edit p95",
  workerVisibleEditElapsed.p95,
  budgets.workerVisibleEdit10kP95Ms,
);
assertBudget(
  "100 pending worker reconnect catch-up p95",
  workerReconnectCatchUpElapsed.p95,
  budgets.workerReconnectCatchUp100PendingP95Ms,
);

console.log(
  JSON.stringify(
    {
      baseBudgets,
      budgets,
      toleranceMultiplier,
      sampleCounts: {
        load100k: loadRuns.length,
        edit10k: editRuns.length,
        rangeAggregates10k: rangeRuns.length,
        topologyEdit10k: topologyRuns.length,
        renderCommit10k: renderRuns.length,
        workerWarmStart100k: workerWarmStartRuns.length,
        workerVisibleEdit10k: workerVisibleEditRuns.length,
        workerReconnectCatchUp100Pending: workerReconnectCatchUpRuns.length,
      },
      results: {
        load100k: {
          scenario: "load",
          materializedCells: 100_000,
          elapsedMs: loadElapsed,
          workingSetDeltaBytes: loadWorkingSetDelta,
          heapUsedAfterBytes: loadHeapUsedAfter,
          runs: loadRuns,
        },
        edit10k: {
          scenario: "single-edit",
          downstreamCount: 10_000,
          elapsedMs: editElapsed,
          recalcMs: editRecalc,
          rssAfterBytes: editRssAfter,
          runs: editRuns,
        },
        rangeAggregates10k: {
          scenario: "range-aggregates",
          sourceCount: 1_024,
          aggregateCount: 10_000,
          elapsedMs: rangeElapsed,
          recalcMs: rangeRecalc,
          rssAfterBytes: rangeRssAfter,
          runs: rangeRuns,
        },
        topologyEdit10k: {
          scenario: "topology-edit",
          chainLength: 10_000,
          elapsedMs: topologyElapsed,
          recalcMs: topologyRecalc,
          rssAfterBytes: topologyRssAfter,
          runs: topologyRuns,
        },
        renderCommit10k: {
          scenario: "render-commit",
          declaredCells: 10_000,
          elapsedMs: renderElapsed,
          rssAfterBytes: renderRssAfter,
          runs: renderRuns,
        },
        workerWarmStart100k: {
          scenario: "worker-warm-start",
          materializedCells: 100_000,
          elapsedMs: workerWarmStartElapsed,
          workingSetDeltaBytes: workerWarmStartWorkingSetDelta,
          runs: workerWarmStartRuns,
        },
        workerVisibleEdit10k: {
          scenario: "worker-visible-edit",
          materializedCells: 10_000,
          visiblePatchMs: workerVisibleEditElapsed,
          commitMs: workerVisibleEditCommitElapsed,
          runs: workerVisibleEditRuns,
        },
        workerReconnectCatchUp100Pending: {
          scenario: "worker-reconnect-catch-up",
          materializedCells: 10_000,
          pendingMutationCount: 100,
          catchUpMs: workerReconnectCatchUpElapsed,
          rebaseMs: workerReconnectRebaseElapsed,
          submitDrainMs: workerReconnectSubmitDrainElapsed,
          ackMs: workerReconnectAckElapsed,
          runs: workerReconnectCatchUpRuns,
        },
      },
    },
    null,
    2,
  ),
);
