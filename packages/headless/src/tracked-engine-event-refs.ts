import type { CellValue, EngineEvent } from '@bilig/protocol'

export interface TrackedCellPatch {
  readonly kind: 'cell'
  readonly cellIndex: number
  readonly address: {
    readonly sheet: number
    readonly row: number
    readonly col: number
  }
  readonly sheetName: string
  readonly a1: string
  readonly newValue: CellValue
}

interface CoreTrackedEngineEvent {
  kind: EngineEvent['kind']
  invalidation: EngineEvent['invalidation']
  changedCellIndices: EngineEvent['changedCellIndices']
  patches?: readonly TrackedCellPatch[]
  invalidatedRanges: EngineEvent['invalidatedRanges']
  invalidatedRows: EngineEvent['invalidatedRows']
  invalidatedColumns: EngineEvent['invalidatedColumns']
  metrics: EngineEvent['metrics']
  explicitChangedCount?: number
}

export interface TrackedEngineEvent {
  invalidation: CoreTrackedEngineEvent['invalidation']
  changedCellIndices: CoreTrackedEngineEvent['changedCellIndices']
  patches?: readonly TrackedCellPatch[]
  changedInputCount: number
  explicitChangedCount?: number
  hasInvalidatedRanges: boolean
  hasInvalidatedRows: boolean
  hasInvalidatedColumns: boolean
}

function readExplicitChangedCount(event: CoreTrackedEngineEvent): number | undefined {
  const explicitChangedCount = Reflect.get(event, 'explicitChangedCount')
  return typeof explicitChangedCount === 'number' && explicitChangedCount >= 0 ? explicitChangedCount : undefined
}

export function captureTrackedEngineEvent(event: CoreTrackedEngineEvent): TrackedEngineEvent {
  return {
    invalidation: event.invalidation,
    changedCellIndices:
      event.changedCellIndices instanceof Uint32Array ? Uint32Array.from(event.changedCellIndices) : [...event.changedCellIndices],
    ...(event.patches ? { patches: [...event.patches] } : {}),
    changedInputCount: event.metrics.changedInputCount,
    explicitChangedCount: readExplicitChangedCount(event),
    hasInvalidatedRanges: event.invalidatedRanges.length > 0,
    hasInvalidatedRows: event.invalidatedRows.length > 0,
    hasInvalidatedColumns: event.invalidatedColumns.length > 0,
  }
}
