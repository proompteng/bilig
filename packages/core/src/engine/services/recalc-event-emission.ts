import type { EngineChangedCell, EngineEvent, RecalcMetrics } from '@bilig/protocol'
import type { EngineTrackedEvent } from '../../events.js'
import type { EnginePatch } from '../../patches/patch-types.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'

const TRACKED_CELL_PATCH_LIMIT = 64

export function captureTrackedRecalcPatchesForCells(input: {
  readonly changed: readonly number[] | U32
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | U32,
    request?: {
      invalidation?: 'cells' | 'full'
      invalidatedRanges?: readonly {
        sheetName: string
        startAddress: string
        endAddress: string
      }[]
      invalidatedRows?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    },
  ) => readonly EnginePatch[]
}): readonly EnginePatch[] | undefined {
  return input.changed.length <= TRACKED_CELL_PATCH_LIMIT
    ? input.captureChangedPatches(input.changed, {
        invalidation: 'cells',
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      })
    : undefined
}

export function emitRecalcBatchEvents(input: {
  readonly state: Pick<EngineRuntimeState, 'events' | 'workbook'>
  readonly changed: readonly number[] | U32
  readonly changedCells: readonly EngineChangedCell[]
  readonly metrics: RecalcMetrics
  readonly explicitChangedCount: number
  readonly captureChangedPatches: Parameters<typeof captureTrackedRecalcPatchesForCells>[0]['captureChangedPatches']
}): void {
  const changedCellIndices: EngineEvent['changedCellIndices'] =
    input.changed instanceof Uint32Array ? input.changed : Array.from(input.changed)
  const event: EngineEvent & {
    explicitChangedCount: number
  } = {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices,
    changedCells: input.changedCells,
    invalidatedRanges: [],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics: input.metrics,
    explicitChangedCount: input.explicitChangedCount,
  }
  input.state.events.emit(event, input.changed, (cellIndex) => input.state.workbook.getQualifiedAddress(cellIndex))
  if (!input.state.events.hasTrackedListeners()) {
    return
  }
  const patches = captureTrackedRecalcPatchesForCells({
    changed: input.changed,
    captureChangedPatches: input.captureChangedPatches,
  })
  const trackedEvent: EngineTrackedEvent = {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices,
    ...(patches ? { patches } : {}),
    invalidatedRanges: [],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics: input.metrics,
    explicitChangedCount: input.explicitChangedCount,
  }
  input.state.events.emitTracked(trackedEvent)
}
