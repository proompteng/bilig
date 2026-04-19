import type { CellValue, EngineEvent } from '@bilig/protocol'
import type { EnginePatch } from '@bilig/core'

const TRACKED_RANGE_INVALIDATION_PATCH_KIND = 'range-invalidation' as const
const TRACKED_ROW_INVALIDATION_PATCH_KIND = 'row-invalidation' as const
const TRACKED_COLUMN_INVALIDATION_PATCH_KIND = 'column-invalidation' as const

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

export interface TrackedRangeInvalidationPatch {
  readonly kind: typeof TRACKED_RANGE_INVALIDATION_PATCH_KIND
  readonly range: {
    readonly sheetName: string
    readonly startAddress: string
    readonly endAddress: string
  }
}

export interface TrackedRowInvalidationPatch {
  readonly kind: typeof TRACKED_ROW_INVALIDATION_PATCH_KIND
  readonly sheetName: string
  readonly startIndex: number
  readonly endIndex: number
}

export interface TrackedColumnInvalidationPatch {
  readonly kind: typeof TRACKED_COLUMN_INVALIDATION_PATCH_KIND
  readonly sheetName: string
  readonly startIndex: number
  readonly endIndex: number
}

export type TrackedPatch = TrackedCellPatch | TrackedRangeInvalidationPatch | TrackedRowInvalidationPatch | TrackedColumnInvalidationPatch

interface CoreTrackedEngineEvent {
  kind: EngineEvent['kind']
  invalidation: EngineEvent['invalidation']
  changedCellIndices: EngineEvent['changedCellIndices']
  patches?: readonly EnginePatch[]
  invalidatedRanges: EngineEvent['invalidatedRanges']
  invalidatedRows: EngineEvent['invalidatedRows']
  invalidatedColumns: EngineEvent['invalidatedColumns']
  metrics: EngineEvent['metrics']
  explicitChangedCount?: number
}

export interface TrackedEngineEvent {
  invalidation: CoreTrackedEngineEvent['invalidation']
  changedCellIndices: CoreTrackedEngineEvent['changedCellIndices']
  patches?: readonly TrackedPatch[]
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

function hasPatchKind(event: CoreTrackedEngineEvent, kind: TrackedPatch['kind']): boolean {
  return event.patches?.some((patch) => patch.kind === kind) ?? false
}

export function captureTrackedEngineEvent(event: CoreTrackedEngineEvent): TrackedEngineEvent {
  return {
    invalidation: event.invalidation,
    changedCellIndices:
      event.changedCellIndices instanceof Uint32Array ? Uint32Array.from(event.changedCellIndices) : [...event.changedCellIndices],
    ...(event.patches ? { patches: [...event.patches] as readonly TrackedPatch[] } : {}),
    changedInputCount: event.metrics.changedInputCount,
    explicitChangedCount: readExplicitChangedCount(event),
    hasInvalidatedRanges: event.invalidatedRanges.length > 0 || hasPatchKind(event, TRACKED_RANGE_INVALIDATION_PATCH_KIND),
    hasInvalidatedRows: event.invalidatedRows.length > 0 || hasPatchKind(event, TRACKED_ROW_INVALIDATION_PATCH_KIND),
    hasInvalidatedColumns: event.invalidatedColumns.length > 0 || hasPatchKind(event, TRACKED_COLUMN_INVALIDATION_PATCH_KIND),
  }
}
