import type { EngineEvent } from '@bilig/protocol'
import type { EngineTrackedEvent } from '../../events.js'
import type { EnginePatchEmitterService } from '../../patches/patch-emitter.js'
import type { EngineRuntimeState } from '../runtime-state.js'

export interface EngineFullInvalidationEvents {
  readonly hasListeners: () => boolean
  readonly hasTrackedListeners: () => boolean
  readonly hasCellListeners: () => boolean
  readonly emitAllWatched: (event: EngineEvent) => void
  readonly emitTracked: (event: EngineTrackedEvent) => void
}

export interface EngineFullInvalidationService {
  readonly emitFullSnapshotInvalidation: (options: { readonly incrementMetrics: boolean }) => void
}

export function createEngineFullInvalidationService(args: {
  readonly state: {
    readonly events: EngineFullInvalidationEvents
    readonly getLastMetrics: Pick<EngineRuntimeState, 'getLastMetrics'>['getLastMetrics']
    readonly setLastMetrics: Pick<EngineRuntimeState, 'setLastMetrics'>['setLastMetrics']
  }
  readonly patchEmitter: Pick<EnginePatchEmitterService, 'captureChangedPatches'>
}): EngineFullInvalidationService {
  return {
    emitFullSnapshotInvalidation(options) {
      let metrics = args.state.getLastMetrics()
      if (options.incrementMetrics) {
        metrics = {
          ...metrics,
          batchId: metrics.batchId + 1,
          changedInputCount: 0,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        }
        args.state.setLastMetrics(metrics)
      }

      const hasGeneralEventListeners = args.state.events.hasListeners()
      const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
      const hasWatchedCellListeners = args.state.events.hasCellListeners()
      if (!hasGeneralEventListeners && !hasTrackedEventListeners && !hasWatchedCellListeners) {
        return
      }

      const changedCellIndices = new Uint32Array()
      const event: EngineEvent & { explicitChangedCount: number } = {
        kind: 'batch',
        invalidation: 'full',
        changedCellIndices,
        changedCells: [],
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics,
        explicitChangedCount: 0,
      }
      if (hasGeneralEventListeners || hasWatchedCellListeners) {
        args.state.events.emitAllWatched(event)
      }
      if (hasTrackedEventListeners) {
        const patches = args.patchEmitter.captureChangedPatches(changedCellIndices, {
          invalidation: 'full',
          invalidatedRanges: [],
          invalidatedRows: [],
          invalidatedColumns: [],
        })
        args.state.events.emitTracked({
          ...event,
          ...(patches.length > 0 ? { patches } : {}),
        })
      }
    },
  }
}
