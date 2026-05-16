import { describe, expect, it, vi } from 'vitest'
import type { EngineEvent } from '@bilig/protocol'
import type { EngineTrackedEvent } from '../events.js'
import { createInitialRecalcMetrics } from '../engine/runtime-state.js'
import { createEngineFullInvalidationService, type EngineFullInvalidationEvents } from '../engine/services/full-invalidation-service.js'
import { ENGINE_RANGE_INVALIDATION_PATCH_KIND, type EnginePatch } from '../patches/patch-types.js'

function makeService(options: {
  readonly hasGeneralListeners?: boolean
  readonly hasTrackedListeners?: boolean
  readonly hasWatchedCellListeners?: boolean
  readonly patches?: readonly EnginePatch[]
}) {
  const previousMetrics = {
    ...createInitialRecalcMetrics(),
    batchId: 11,
    changedInputCount: 4,
    dirtyFormulaCount: 3,
    wasmFormulaCount: 2,
    jsFormulaCount: 1,
    rangeNodeVisits: 99,
    recalcMs: 8,
    compileMs: 6,
  }
  const emitAllWatched = vi.fn((_: EngineEvent) => {})
  const emitTracked = vi.fn((_: EngineTrackedEvent) => {})
  const captureChangedPatches = vi.fn((_: readonly number[] | Uint32Array) => options.patches ?? [])
  const setLastMetrics = vi.fn((_: typeof previousMetrics) => {})
  const events: EngineFullInvalidationEvents = {
    hasListeners: () => options.hasGeneralListeners === true,
    hasTrackedListeners: () => options.hasTrackedListeners === true,
    hasCellListeners: () => options.hasWatchedCellListeners === true,
    emitAllWatched,
    emitTracked,
  }

  return {
    service: createEngineFullInvalidationService({
      state: {
        events,
        getLastMetrics: () => previousMetrics,
        setLastMetrics,
      },
      patchEmitter: {
        captureChangedPatches,
      },
    }),
    emitAllWatched,
    emitTracked,
    captureChangedPatches,
    setLastMetrics,
  }
}

describe('EngineFullInvalidationService', () => {
  it('increments and resets batch metrics before notifying watched listeners', () => {
    const { service, emitAllWatched, emitTracked, captureChangedPatches, setLastMetrics } = makeService({
      hasWatchedCellListeners: true,
    })

    service.emitFullSnapshotInvalidation({ incrementMetrics: true })

    expect(setLastMetrics).toHaveBeenCalledWith(
      expect.objectContaining({
        batchId: 12,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      }),
    )
    expect(emitAllWatched).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
        changedCellIndices: new Uint32Array(),
        changedCells: [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        explicitChangedCount: 0,
        metrics: expect.objectContaining({ batchId: 12 }),
      }),
    )
    expect(captureChangedPatches).not.toHaveBeenCalled()
    expect(emitTracked).not.toHaveBeenCalled()
  })

  it('emits tracked full invalidation patches without notifying general listeners', () => {
    const patch: EnginePatch = {
      kind: ENGINE_RANGE_INVALIDATION_PATCH_KIND,
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
    }
    const { service, emitAllWatched, emitTracked, captureChangedPatches, setLastMetrics } = makeService({
      hasTrackedListeners: true,
      patches: [patch],
    })

    service.emitFullSnapshotInvalidation({ incrementMetrics: false })

    expect(setLastMetrics).not.toHaveBeenCalled()
    expect(emitAllWatched).not.toHaveBeenCalled()
    expect(captureChangedPatches).toHaveBeenCalledWith(new Uint32Array(), {
      invalidation: 'full',
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
    })
    expect(emitTracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
        changedCellIndices: new Uint32Array(),
        explicitChangedCount: 0,
        patches: [patch],
        metrics: expect.objectContaining({ batchId: 11 }),
      }),
    )
  })
})
