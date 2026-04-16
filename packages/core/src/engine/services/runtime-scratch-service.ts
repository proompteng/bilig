import { Effect } from 'effect'
import { growUint32 } from '../../engine-buffer-utils.js'
import type { U32 } from '../runtime-state.js'
import { EngineRuntimeScratchError } from '../errors.js'

function scratchErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export interface EngineRuntimeScratchService {
  readonly ensureRecalcCapacity: (size: number) => Effect.Effect<void, EngineRuntimeScratchError>
  readonly ensureRecalcCapacityNow: (size: number) => void
  readonly getPendingKernelSyncNow: () => U32
  readonly setPendingKernelSyncNow: (next: U32) => void
  readonly getDeferredKernelSyncCountNow: () => number
  readonly setDeferredKernelSyncCountNow: (next: number) => void
  readonly getDeferredKernelSyncEpochNow: () => number
  readonly setDeferredKernelSyncEpochNow: (next: number) => void
  readonly getDeferredKernelSyncSeenNow: () => U32
  readonly setDeferredKernelSyncSeenNow: (next: U32) => void
  readonly getWasmBatchNow: () => U32
  readonly setWasmBatchNow: (next: U32) => void
  readonly getMutationRootsNow: () => U32
  readonly setMutationRootsNow: (next: U32) => void
  readonly getChangedInputEpochNow: () => number
  readonly setChangedInputEpochNow: (next: number) => void
  readonly getChangedInputSeenNow: () => U32
  readonly setChangedInputSeenNow: (next: U32) => void
  readonly getChangedInputBufferNow: () => U32
  readonly setChangedInputBufferNow: (next: U32) => void
  readonly getChangedFormulaEpochNow: () => number
  readonly setChangedFormulaEpochNow: (next: number) => void
  readonly getChangedFormulaSeenNow: () => U32
  readonly setChangedFormulaSeenNow: (next: U32) => void
  readonly getChangedFormulaBufferNow: () => U32
  readonly setChangedFormulaBufferNow: (next: U32) => void
  readonly getChangedUnionEpochNow: () => number
  readonly setChangedUnionEpochNow: (next: number) => void
  readonly getChangedUnionSeenNow: () => U32
  readonly setChangedUnionSeenNow: (next: U32) => void
  readonly getChangedUnionNow: () => U32
  readonly setChangedUnionNow: (next: U32) => void
  readonly getMaterializedCellCountNow: () => number
  readonly setMaterializedCellCountNow: (next: number) => void
  readonly getMaterializedCellsNow: () => U32
  readonly setMaterializedCellsNow: (next: U32) => void
  readonly getExplicitChangedEpochNow: () => number
  readonly setExplicitChangedEpochNow: (next: number) => void
  readonly getExplicitChangedSeenNow: () => U32
  readonly setExplicitChangedSeenNow: (next: U32) => void
  readonly getExplicitChangedBufferNow: () => U32
  readonly setExplicitChangedBufferNow: (next: U32) => void
  readonly getImpactedFormulaEpochNow: () => number
  readonly setImpactedFormulaEpochNow: (next: number) => void
  readonly getImpactedFormulaSeenNow: () => U32
  readonly setImpactedFormulaSeenNow: (next: U32) => void
  readonly getImpactedFormulaBufferNow: () => U32
  readonly setImpactedFormulaBufferNow: (next: U32) => void
}

export function createEngineRuntimeScratchService(): EngineRuntimeScratchService {
  let pendingKernelSync: U32 = new Uint32Array(128)
  let deferredKernelSyncCount = 0
  let deferredKernelSyncEpoch = 1
  let deferredKernelSyncSeen: U32 = new Uint32Array(128)
  let wasmBatch: U32 = new Uint32Array(128)
  let mutationRoots: U32 = new Uint32Array(128)
  let changedInputEpoch = 1
  let changedInputSeen: U32 = new Uint32Array(128)
  let changedInputBuffer: U32 = new Uint32Array(128)
  let changedFormulaEpoch = 1
  let changedFormulaSeen: U32 = new Uint32Array(128)
  let changedFormulaBuffer: U32 = new Uint32Array(128)
  let changedUnionEpoch = 1
  let changedUnionSeen: U32 = new Uint32Array(128)
  let changedUnion: U32 = new Uint32Array(128)
  let materializedCellCount = 0
  let materializedCells: U32 = new Uint32Array(128)
  let explicitChangedEpoch = 1
  let explicitChangedSeen: U32 = new Uint32Array(128)
  let explicitChangedBuffer: U32 = new Uint32Array(128)
  let impactedFormulaEpoch = 1
  let impactedFormulaSeen: U32 = new Uint32Array(128)
  let impactedFormulaBuffer: U32 = new Uint32Array(128)

  const ensureRecalcCapacityNow = (size: number): void => {
    if (size > mutationRoots.length) {
      mutationRoots = growUint32(mutationRoots, size)
    }
    if (size > changedInputSeen.length) {
      changedInputSeen = growUint32(changedInputSeen, size)
    }
    if (size > changedInputBuffer.length) {
      changedInputBuffer = growUint32(changedInputBuffer, size)
    }
    if (size > changedFormulaSeen.length) {
      changedFormulaSeen = growUint32(changedFormulaSeen, size)
    }
    if (size > changedFormulaBuffer.length) {
      changedFormulaBuffer = growUint32(changedFormulaBuffer, size)
    }
    if (size > pendingKernelSync.length) {
      pendingKernelSync = growUint32(pendingKernelSync, size)
    }
    if (size > deferredKernelSyncSeen.length) {
      deferredKernelSyncSeen = growUint32(deferredKernelSyncSeen, size)
    }
    if (size > wasmBatch.length) {
      wasmBatch = growUint32(wasmBatch, size)
    }
    if (size > changedUnion.length) {
      changedUnion = growUint32(changedUnion, size)
    }
    if (size > changedUnionSeen.length) {
      changedUnionSeen = growUint32(changedUnionSeen, size)
    }
    if (size > materializedCells.length) {
      materializedCells = growUint32(materializedCells, size)
    }
    if (size > explicitChangedSeen.length) {
      explicitChangedSeen = growUint32(explicitChangedSeen, size)
    }
    if (size > explicitChangedBuffer.length) {
      explicitChangedBuffer = growUint32(explicitChangedBuffer, size)
    }
    if (size > impactedFormulaSeen.length) {
      impactedFormulaSeen = growUint32(impactedFormulaSeen, size)
    }
    if (size > impactedFormulaBuffer.length) {
      impactedFormulaBuffer = growUint32(impactedFormulaBuffer, size)
    }
  }

  return {
    ensureRecalcCapacity(size) {
      return Effect.try({
        try: () => {
          ensureRecalcCapacityNow(size)
        },
        catch: (cause) =>
          new EngineRuntimeScratchError({
            message: scratchErrorMessage('Failed to ensure recalc scratch capacity', cause),
            cause,
          }),
      })
    },
    ensureRecalcCapacityNow,
    getPendingKernelSyncNow: () => pendingKernelSync,
    setPendingKernelSyncNow: (next) => {
      pendingKernelSync = next
    },
    getDeferredKernelSyncCountNow: () => deferredKernelSyncCount,
    setDeferredKernelSyncCountNow: (next) => {
      deferredKernelSyncCount = next
    },
    getDeferredKernelSyncEpochNow: () => deferredKernelSyncEpoch,
    setDeferredKernelSyncEpochNow: (next) => {
      deferredKernelSyncEpoch = next
    },
    getDeferredKernelSyncSeenNow: () => deferredKernelSyncSeen,
    setDeferredKernelSyncSeenNow: (next) => {
      deferredKernelSyncSeen = next
    },
    getWasmBatchNow: () => wasmBatch,
    setWasmBatchNow: (next) => {
      wasmBatch = next
    },
    getMutationRootsNow: () => mutationRoots,
    setMutationRootsNow: (next) => {
      mutationRoots = next
    },
    getChangedInputEpochNow: () => changedInputEpoch,
    setChangedInputEpochNow: (next) => {
      changedInputEpoch = next
    },
    getChangedInputSeenNow: () => changedInputSeen,
    setChangedInputSeenNow: (next) => {
      changedInputSeen = next
    },
    getChangedInputBufferNow: () => changedInputBuffer,
    setChangedInputBufferNow: (next) => {
      changedInputBuffer = next
    },
    getChangedFormulaEpochNow: () => changedFormulaEpoch,
    setChangedFormulaEpochNow: (next) => {
      changedFormulaEpoch = next
    },
    getChangedFormulaSeenNow: () => changedFormulaSeen,
    setChangedFormulaSeenNow: (next) => {
      changedFormulaSeen = next
    },
    getChangedFormulaBufferNow: () => changedFormulaBuffer,
    setChangedFormulaBufferNow: (next) => {
      changedFormulaBuffer = next
    },
    getChangedUnionEpochNow: () => changedUnionEpoch,
    setChangedUnionEpochNow: (next) => {
      changedUnionEpoch = next
    },
    getChangedUnionSeenNow: () => changedUnionSeen,
    setChangedUnionSeenNow: (next) => {
      changedUnionSeen = next
    },
    getChangedUnionNow: () => changedUnion,
    setChangedUnionNow: (next) => {
      changedUnion = next
    },
    getMaterializedCellCountNow: () => materializedCellCount,
    setMaterializedCellCountNow: (next) => {
      materializedCellCount = next
    },
    getMaterializedCellsNow: () => materializedCells,
    setMaterializedCellsNow: (next) => {
      materializedCells = next
    },
    getExplicitChangedEpochNow: () => explicitChangedEpoch,
    setExplicitChangedEpochNow: (next) => {
      explicitChangedEpoch = next
    },
    getExplicitChangedSeenNow: () => explicitChangedSeen,
    setExplicitChangedSeenNow: (next) => {
      explicitChangedSeen = next
    },
    getExplicitChangedBufferNow: () => explicitChangedBuffer,
    setExplicitChangedBufferNow: (next) => {
      explicitChangedBuffer = next
    },
    getImpactedFormulaEpochNow: () => impactedFormulaEpoch,
    setImpactedFormulaEpochNow: (next) => {
      impactedFormulaEpoch = next
    },
    getImpactedFormulaSeenNow: () => impactedFormulaSeen,
    setImpactedFormulaSeenNow: (next) => {
      impactedFormulaSeen = next
    },
    getImpactedFormulaBufferNow: () => impactedFormulaBuffer,
    setImpactedFormulaBufferNow: (next) => {
      impactedFormulaBuffer = next
    },
  }
}
