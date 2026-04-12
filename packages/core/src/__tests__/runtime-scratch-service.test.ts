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

  it("roundtrips every scratch buffer setter and getter through the extracted boundary", () => {
    const scratch = createEngineRuntimeScratchService();
    const pendingKernelSync = new Uint32Array([1, 2]);
    const wasmBatch = new Uint32Array([3, 4]);
    const mutationRoots = new Uint32Array([5, 6]);
    const changedInputSeen = new Uint32Array([7, 8]);
    const changedInputBuffer = new Uint32Array([9, 10]);
    const changedFormulaSeen = new Uint32Array([11, 12]);
    const changedFormulaBuffer = new Uint32Array([13, 14]);
    const changedUnionSeen = new Uint32Array([15, 16]);
    const changedUnion = new Uint32Array([17, 18]);
    const materializedCells = new Uint32Array([19, 20]);
    const explicitChangedSeen = new Uint32Array([21, 22]);
    const explicitChangedBuffer = new Uint32Array([23, 24]);
    const impactedFormulaSeen = new Uint32Array([25, 26]);
    const impactedFormulaBuffer = new Uint32Array([27, 28]);

    scratch.setPendingKernelSyncNow(pendingKernelSync);
    scratch.setWasmBatchNow(wasmBatch);
    scratch.setMutationRootsNow(mutationRoots);
    scratch.setChangedInputSeenNow(changedInputSeen);
    scratch.setChangedInputBufferNow(changedInputBuffer);
    scratch.setChangedFormulaSeenNow(changedFormulaSeen);
    scratch.setChangedFormulaBufferNow(changedFormulaBuffer);
    scratch.setChangedUnionSeenNow(changedUnionSeen);
    scratch.setChangedUnionNow(changedUnion);
    scratch.setMaterializedCellsNow(materializedCells);
    scratch.setExplicitChangedSeenNow(explicitChangedSeen);
    scratch.setExplicitChangedBufferNow(explicitChangedBuffer);
    scratch.setImpactedFormulaSeenNow(impactedFormulaSeen);
    scratch.setImpactedFormulaBufferNow(impactedFormulaBuffer);

    expect(scratch.getPendingKernelSyncNow()).toBe(pendingKernelSync);
    expect(scratch.getWasmBatchNow()).toBe(wasmBatch);
    expect(scratch.getMutationRootsNow()).toBe(mutationRoots);
    expect(scratch.getChangedInputSeenNow()).toBe(changedInputSeen);
    expect(scratch.getChangedInputBufferNow()).toBe(changedInputBuffer);
    expect(scratch.getChangedFormulaSeenNow()).toBe(changedFormulaSeen);
    expect(scratch.getChangedFormulaBufferNow()).toBe(changedFormulaBuffer);
    expect(scratch.getChangedUnionSeenNow()).toBe(changedUnionSeen);
    expect(scratch.getChangedUnionNow()).toBe(changedUnion);
    expect(scratch.getMaterializedCellsNow()).toBe(materializedCells);
    expect(scratch.getExplicitChangedSeenNow()).toBe(explicitChangedSeen);
    expect(scratch.getExplicitChangedBufferNow()).toBe(explicitChangedBuffer);
    expect(scratch.getImpactedFormulaSeenNow()).toBe(impactedFormulaSeen);
    expect(scratch.getImpactedFormulaBufferNow()).toBe(impactedFormulaBuffer);
  });

  it("wraps allocation failures when recalc scratch growth cannot be satisfied", () => {
    const scratch = createEngineRuntimeScratchService();

    expect(() => Effect.runSync(scratch.ensureRecalcCapacity(Number.MAX_SAFE_INTEGER))).toThrow(
      "Invalid typed array length",
    );
  });
});
