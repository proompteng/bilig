import { describe, expect, it } from "vitest";
import { runPerfSmokeBenchmark } from "../../../../scripts/perf-smoke.ts";

describe("perf smoke", () => {
  it("runs the benchmark edit scenario through node and tsx", async () => {
    const result = await runPerfSmokeBenchmark(100);

    expect(result.downstreamCount).toBe(100);
    expect(result.elapsedMs).toBeGreaterThanOrEqual(0);
    expect(result.metrics.dirtyFormulaCount).toBeGreaterThanOrEqual(100);
    expect(result.metrics.wasmFormulaCount).toBeGreaterThan(0);
  });
});
