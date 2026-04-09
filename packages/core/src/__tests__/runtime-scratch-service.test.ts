import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { createEngineRuntimeScratchService } from "../engine/services/runtime-scratch-service.js";

describe("EngineRuntimeScratchService", () => {
  it("grows recalc scratch buffers beyond the initial capacity and preserves seeded values", () => {
    const scratch = createEngineRuntimeScratchService();
    scratch.getPendingKernelSyncNow()[0] = 7;
    scratch.getChangedInputSeenNow()[0] = 11;
    scratch.getImpactedFormulaBufferNow()[0] = 19;

    Effect.runSync(scratch.ensureRecalcCapacity(256));

    expect(scratch.getPendingKernelSyncNow().length).toBeGreaterThanOrEqual(256);
    expect(scratch.getChangedInputSeenNow().length).toBeGreaterThanOrEqual(256);
    expect(scratch.getImpactedFormulaBufferNow().length).toBeGreaterThanOrEqual(256);
    expect(scratch.getPendingKernelSyncNow()[0]).toBe(7);
    expect(scratch.getChangedInputSeenNow()[0]).toBe(11);
    expect(scratch.getImpactedFormulaBufferNow()[0]).toBe(19);
  });

  it("tracks epochs and materialized cell counters through the extracted scratch boundary", () => {
    const scratch = createEngineRuntimeScratchService();

    scratch.setChangedInputEpochNow(5);
    scratch.setChangedFormulaEpochNow(6);
    scratch.setChangedUnionEpochNow(7);
    scratch.setExplicitChangedEpochNow(8);
    scratch.setImpactedFormulaEpochNow(9);
    scratch.setMaterializedCellCountNow(4);

    expect(scratch.getChangedInputEpochNow()).toBe(5);
    expect(scratch.getChangedFormulaEpochNow()).toBe(6);
    expect(scratch.getChangedUnionEpochNow()).toBe(7);
    expect(scratch.getExplicitChangedEpochNow()).toBe(8);
    expect(scratch.getImpactedFormulaEpochNow()).toBe(9);
    expect(scratch.getMaterializedCellCountNow()).toBe(4);
  });
});
