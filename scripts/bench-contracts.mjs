#!/usr/bin/env bun

import { fileURLToPath } from "node:url";

const rootDir = fileURLToPath(new URL("../", import.meta.url));
const textDecoder = new TextDecoder();

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
  renderCommit10kP95Ms: 50
};
const toleranceMultiplier = Number.parseFloat(process.env.BILIG_BENCH_TOLERANCE ?? (process.env.CI ? "1.5" : "1"));
const budgets = Object.fromEntries(
  Object.entries(baseBudgets).map(([key, value]) => [key, value * toleranceMultiplier])
);

function assertBudget(label, actual, threshold, formatter = formatMs) {
  if (actual > threshold) {
    throw new Error(`${label} exceeded budget: ${formatter(actual)} > ${formatter(threshold)}`);
  }
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
  const samples = [...values].sort((left, right) => left - right);
  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length;
  return {
    samples,
    min: samples[0],
    median: quantile(samples, 0.5),
    p95: quantile(samples, 0.95),
    max: samples[samples.length - 1],
    mean
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

function runBenchmarkScript(scriptRelativePath, arg) {
  const args = Array.isArray(arg) ? arg.map(String) : [String(arg)];
  const result = Bun.spawnSync([process.execPath, scriptRelativePath, ...args], {
    cwd: rootDir,
    env: process.env,
    stdin: "ignore",
    stdout: "pipe",
    stderr: "pipe"
  });
  if (result.exitCode !== 0) {
    throw new Error(
      `Benchmark script failed (${scriptRelativePath} ${arg})\nstdout:\n${textDecoder.decode(result.stdout)}\nstderr:\n${textDecoder.decode(result.stderr)}`
    );
  }
  return JSON.parse(textDecoder.decode(result.stdout).trim());
}

function sampleBenchmark(scriptRelativePath, arg, iterations) {
  const runs = [];
  for (let index = 0; index < iterations; index += 1) {
    runs.push(runBenchmarkScript(scriptRelativePath, arg));
  }
  return runs;
}

const loadRuns = sampleBenchmark("packages/benchmarks/src/benchmark-load.ts", 100_000, 3);
const editRuns = sampleBenchmark("packages/benchmarks/src/benchmark-edit.ts", 10_000, 5);
const rangeRuns = sampleBenchmark("packages/benchmarks/src/benchmark-range-heavy.ts", [1_024, 10_000], 3);
const topologyRuns = sampleBenchmark("packages/benchmarks/src/benchmark-topology-edit.ts", 10_000, 3);
const renderRuns = sampleBenchmark("packages/benchmarks/src/benchmark-renderer.ts", 10_000, 5);

const loadElapsed = summarizeNumbers(loadRuns.map((run) => run.elapsedMs));
const loadWorkingSetDelta = summarizeNumbers(
  loadRuns.map((run) => run.memory.delta.heapUsedBytes + run.memory.delta.externalBytes)
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

assertBudget("100k snapshot load p95", loadElapsed.p95, budgets.load100kP95Ms);
assertBudget(
  "100k snapshot load working-set delta",
  loadWorkingSetDelta.max,
  budgets.load100kWorkingSetDeltaBytes,
  formatBytes
);
assertBudget("10k downstream edit p95", editElapsed.p95, budgets.edit10kElapsedP95Ms);
assertBudget("10k downstream recalc median", editRecalc.median, budgets.edit10kRecalcMedianMs);
assertBudget("10k downstream recalc p95", editRecalc.p95, budgets.edit10kRecalcP95Ms);
assertBudget("10k range aggregate edit p95", rangeElapsed.p95, budgets.rangeAggregates10kElapsedP95Ms);
assertBudget("10k range aggregate recalc p95", rangeRecalc.p95, budgets.rangeAggregates10kRecalcP95Ms);
assertBudget("10k topology edit p95", topologyElapsed.p95, budgets.topologyEdit10kElapsedP95Ms);
assertBudget("10k topology recalc p95", topologyRecalc.p95, budgets.topologyEdit10kRecalcP95Ms);
assertBudget("10k render commit p95", renderElapsed.p95, budgets.renderCommit10kP95Ms);

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
        renderCommit10k: renderRuns.length
      },
      results: {
        load100k: {
          scenario: "load",
          materializedCells: 100_000,
          elapsedMs: loadElapsed,
          workingSetDeltaBytes: loadWorkingSetDelta,
          heapUsedAfterBytes: loadHeapUsedAfter,
          runs: loadRuns
        },
        edit10k: {
          scenario: "single-edit",
          downstreamCount: 10_000,
          elapsedMs: editElapsed,
          recalcMs: editRecalc,
          rssAfterBytes: editRssAfter,
          runs: editRuns
        },
        rangeAggregates10k: {
          scenario: "range-aggregates",
          sourceCount: 1_024,
          aggregateCount: 10_000,
          elapsedMs: rangeElapsed,
          recalcMs: rangeRecalc,
          rssAfterBytes: rangeRssAfter,
          runs: rangeRuns
        },
        topologyEdit10k: {
          scenario: "topology-edit",
          chainLength: 10_000,
          elapsedMs: topologyElapsed,
          recalcMs: topologyRecalc,
          rssAfterBytes: topologyRssAfter,
          runs: topologyRuns
        },
        renderCommit10k: {
          scenario: "render-commit",
          declaredCells: 10_000,
          elapsedMs: renderElapsed,
          rssAfterBytes: renderRssAfter,
          runs: renderRuns
        }
      }
    },
    null,
    2
  )
);
