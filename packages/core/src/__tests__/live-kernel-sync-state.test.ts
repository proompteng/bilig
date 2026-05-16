import { describe, expect, it } from 'vitest'
import { deferKernelSyncNow } from '../engine/services/live-kernel-sync-state.js'
import { createEngineRuntimeScratchService } from '../engine/services/runtime-scratch-service.js'

describe('live kernel sync state', () => {
  it('defers unique kernel sync cells and preserves existing pending cells', () => {
    const scratch = createEngineRuntimeScratchService()
    scratch.getPendingKernelSyncNow()[0] = 2
    scratch.setDeferredKernelSyncCountNow(1)

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: Uint32Array.of(2, 5, 5, 7),
    })

    expect(scratch.getDeferredKernelSyncCountNow()).toBe(3)
    expect(Array.from(scratch.getPendingKernelSyncNow().slice(0, 3))).toEqual([2, 5, 7])
  })

  it('does not advance the epoch for an empty defer request', () => {
    const scratch = createEngineRuntimeScratchService()
    const beforeEpoch = scratch.getDeferredKernelSyncEpochNow()

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: [],
    })

    expect(scratch.getDeferredKernelSyncEpochNow()).toBe(beforeEpoch)
    expect(scratch.getDeferredKernelSyncCountNow()).toBe(0)
  })

  it('wraps the deferred seen epoch and clears stale markers', () => {
    const scratch = createEngineRuntimeScratchService()
    scratch.setDeferredKernelSyncEpochNow(0xffff_fffe)
    scratch.getDeferredKernelSyncSeenNow()[3] = 0xffff_fffe

    deferKernelSyncNow({
      scratch,
      cellStoreSize: 8,
      cellIndices: [3],
    })

    expect(scratch.getDeferredKernelSyncEpochNow()).toBe(1)
    expect(scratch.getDeferredKernelSyncSeenNow()[2]).toBe(0)
    expect(scratch.getDeferredKernelSyncSeenNow()[3]).toBe(1)
    expect(Array.from(scratch.getPendingKernelSyncNow().slice(0, 1))).toEqual([3])
  })
})
