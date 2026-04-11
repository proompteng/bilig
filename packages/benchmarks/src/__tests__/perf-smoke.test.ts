import { describe, expect, it, vi } from "vitest";
import {
  runPerfSmokeBenchmark,
  runPerfSmokeGate,
  type PerfSmokeBenchmarkResult,
} from "../../../../scripts/perf-smoke.ts";

describe("perf smoke", () => {
  it("runs the benchmark edit scenario through node and tsx", async () => {
    const result = await runPerfSmokeBenchmark(100);

    expect(result.downstreamCount).toBe(100);
    expect(result.elapsedMs).toBeGreaterThanOrEqual(0);
    expect(result.metrics.dirtyFormulaCount).toBeGreaterThanOrEqual(100);
    expect(result.metrics.wasmFormulaCount).toBeGreaterThan(0);
  });

  it("retries once after building wasm when the first pass falls back to js", async () => {
    const jsOnly: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        dirtyFormulaCount: 100,
        wasmFormulaCount: 0,
      },
    };
    const wasmReady: PerfSmokeBenchmarkResult = {
      ...jsOnly,
      metrics: {
        ...jsOnly.metrics,
        wasmFormulaCount: 100,
      },
    };
    const runBenchmark = vi.fn(async () =>
      runBenchmark.mock.calls.length === 1 ? jsOnly : wasmReady,
    );
    const buildWasm = vi.fn(async () => {});

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    });

    expect(result).toEqual(wasmReady);
    expect(runBenchmark).toHaveBeenCalledTimes(2);
    expect(buildWasm).toHaveBeenCalledTimes(1);
  });

  it("does not build wasm when the first pass already uses the fast path", async () => {
    const wasmReady: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        dirtyFormulaCount: 100,
        wasmFormulaCount: 100,
      },
    };
    const runBenchmark = vi.fn(async () => wasmReady);
    const buildWasm = vi.fn(async () => {});

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    });

    expect(result).toEqual(wasmReady);
    expect(runBenchmark).toHaveBeenCalledTimes(1);
    expect(buildWasm).not.toHaveBeenCalled();
  });
});
