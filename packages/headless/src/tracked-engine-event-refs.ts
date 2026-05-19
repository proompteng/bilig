import type { CellValue, EngineEvent } from '@bilig/protocol'
import type { EnginePatch } from '@bilig/core/headless-runtime'

const TRACKED_RANGE_INVALIDATION_PATCH_KIND = 'range-invalidation' as const
const TRACKED_ROW_INVALIDATION_PATCH_KIND = 'row-invalidation' as const
const TRACKED_COLUMN_INVALIDATION_PATCH_KIND = 'column-invalidation' as const
const TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY = '__biligTrackedPhysicalSheetId'
const TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY = '__biligTrackedPhysicalSortedSliceSplit'

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

export interface CoreTrackedEngineEvent {
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
  changedCellIndicesSortedDisjoint: boolean
  firstChangedCellIndex?: number
  lastChangedCellIndex?: number
  trustedPhysicalSheetId?: number
  trustedSortedSliceSplit?: number
  hasInvalidatedRanges: boolean
  hasInvalidatedRows: boolean
  hasInvalidatedColumns: boolean
}

interface CaptureTrackedEngineEventOptions {
  readonly cloneChangedCellIndices?: boolean
  readonly borrowChangedCellIndexViews?: boolean
}

interface CapturedChangedCellIndices {
  readonly changedCellIndices: CoreTrackedEngineEvent['changedCellIndices']
  readonly sortedDisjoint: boolean
  readonly firstChangedCellIndex?: number
  readonly lastChangedCellIndex?: number
  readonly trustedPhysicalSheetId?: number
  readonly trustedSortedSliceSplit?: number
}

function buildCapturedChangedCellIndices(
  changedCellIndices: CoreTrackedEngineEvent['changedCellIndices'],
  sortedDisjoint: boolean,
  firstChangedCellIndex: number | undefined,
  lastChangedCellIndex: number | undefined,
): CapturedChangedCellIndices {
  const captured: {
    changedCellIndices: CoreTrackedEngineEvent['changedCellIndices']
    sortedDisjoint: boolean
    firstChangedCellIndex?: number
    lastChangedCellIndex?: number
  } = { changedCellIndices, sortedDisjoint }
  if (firstChangedCellIndex !== undefined) {
    captured.firstChangedCellIndex = firstChangedCellIndex
  }
  if (lastChangedCellIndex !== undefined) {
    captured.lastChangedCellIndex = lastChangedCellIndex
  }
  return captured
}

function hasPatchKind(event: CoreTrackedEngineEvent, kind: TrackedPatch['kind']): boolean {
  return event.patches?.some((patch) => patch.kind === kind) ?? false
}

function readTrustedPhysicalSplitMetadata(changedCellIndices: CoreTrackedEngineEvent['changedCellIndices']):
  | {
      readonly trustedPhysicalSheetId: number
      readonly trustedSortedSliceSplit: number
    }
  | undefined {
  if (!(changedCellIndices instanceof Uint32Array)) {
    return undefined
  }
  const trustedPhysicalSheetId = Reflect.get(changedCellIndices, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY)
  const trustedSortedSliceSplit = Reflect.get(changedCellIndices, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY)
  if (
    typeof trustedPhysicalSheetId !== 'number' ||
    !Number.isInteger(trustedPhysicalSheetId) ||
    trustedPhysicalSheetId < 0 ||
    typeof trustedSortedSliceSplit !== 'number' ||
    !Number.isInteger(trustedSortedSliceSplit) ||
    trustedSortedSliceSplit <= 0 ||
    trustedSortedSliceSplit >= changedCellIndices.length
  ) {
    return undefined
  }
  return {
    trustedPhysicalSheetId,
    trustedSortedSliceSplit,
  }
}

function copyTrustedPhysicalSplitMetadata(
  target: Uint32Array,
  trusted: {
    readonly trustedPhysicalSheetId: number
    readonly trustedSortedSliceSplit: number
  },
): void {
  Reflect.set(target, TRUSTED_TRACKED_PHYSICAL_SHEET_ID_PROPERTY, trusted.trustedPhysicalSheetId)
  Reflect.set(target, TRUSTED_TRACKED_PHYSICAL_SORTED_SPLIT_PROPERTY, trusted.trustedSortedSliceSplit)
}

function captureChangedCellIndices(
  changedCellIndices: CoreTrackedEngineEvent['changedCellIndices'],
  cloneChangedCellIndices: boolean,
  borrowChangedCellIndexViews: boolean,
): CapturedChangedCellIndices {
  const length = changedCellIndices.length
  if (!cloneChangedCellIndices && borrowChangedCellIndexViews && changedCellIndices instanceof Uint32Array && length <= 2) {
    if (length === 0) {
      return { changedCellIndices, sortedDisjoint: true }
    }
    const first = changedCellIndices[0]!
    if (length === 1) {
      return {
        changedCellIndices,
        sortedDisjoint: Number.isInteger(first) && first >= 0,
        firstChangedCellIndex: first,
        lastChangedCellIndex: first,
      }
    }
    const last = changedCellIndices[1]!
    return {
      changedCellIndices,
      sortedDisjoint: Number.isInteger(first) && first >= 0 && Number.isInteger(last) && last > first,
      firstChangedCellIndex: first,
      lastChangedCellIndex: last,
    }
  }
  const trustedPhysicalSplit = readTrustedPhysicalSplitMetadata(changedCellIndices)
  if (trustedPhysicalSplit !== undefined && changedCellIndices instanceof Uint32Array) {
    const capturedChangedCellIndices =
      cloneChangedCellIndices || !borrowChangedCellIndexViews ? new Uint32Array(changedCellIndices) : changedCellIndices
    if (capturedChangedCellIndices !== changedCellIndices) {
      copyTrustedPhysicalSplitMetadata(capturedChangedCellIndices, trustedPhysicalSplit)
    }
    return {
      changedCellIndices: capturedChangedCellIndices,
      sortedDisjoint: true,
      ...(length === 0
        ? {}
        : {
            firstChangedCellIndex: capturedChangedCellIndices[0],
            lastChangedCellIndex: capturedChangedCellIndices[length - 1],
          }),
      ...trustedPhysicalSplit,
    }
  }
  let sortedDisjoint = true
  let previous = -1
  let firstChangedCellIndex: number | undefined
  let lastChangedCellIndex: number | undefined
  const noteCellIndex = (index: number, value: number): void => {
    if (index === 0) {
      firstChangedCellIndex = value
    }
    if (!Number.isInteger(value) || value < 0 || value <= previous) {
      sortedDisjoint = false
    }
    previous = value
    lastChangedCellIndex = value
  }
  if (
    !cloneChangedCellIndices &&
    changedCellIndices instanceof Uint32Array &&
    (borrowChangedCellIndexViews ||
      (changedCellIndices.byteOffset === 0 && changedCellIndices.byteLength === changedCellIndices.buffer.byteLength))
  ) {
    for (let index = 0; index < length; index += 1) {
      noteCellIndex(index, changedCellIndices[index]!)
    }
    return buildCapturedChangedCellIndices(changedCellIndices, sortedDisjoint, firstChangedCellIndex, lastChangedCellIndex)
  }
  if (!cloneChangedCellIndices && borrowChangedCellIndexViews) {
    for (let index = 0; index < length; index += 1) {
      noteCellIndex(index, changedCellIndices[index]!)
    }
    return buildCapturedChangedCellIndices(changedCellIndices, sortedDisjoint, firstChangedCellIndex, lastChangedCellIndex)
  }
  if (changedCellIndices instanceof Uint32Array) {
    const copied = new Uint32Array(length)
    for (let index = 0; index < length; index += 1) {
      const value = changedCellIndices[index]!
      copied[index] = value
      noteCellIndex(index, value)
    }
    return buildCapturedChangedCellIndices(copied, sortedDisjoint, firstChangedCellIndex, lastChangedCellIndex)
  }
  const copied = Array.from({ length }, () => 0)
  for (let index = 0; index < length; index += 1) {
    const value = changedCellIndices[index]!
    copied[index] = value
    noteCellIndex(index, value)
  }
  return buildCapturedChangedCellIndices(copied, sortedDisjoint, firstChangedCellIndex, lastChangedCellIndex)
}

export function canUseSortedDisjointTrackedEngineEventChanges(events: readonly TrackedEngineEvent[]): boolean {
  let previousLastChangedCellIndex = -1
  for (const event of events) {
    if (
      event.invalidation === 'full' ||
      (event.patches !== undefined && event.patches.length > 0) ||
      event.hasInvalidatedRanges ||
      event.hasInvalidatedRows ||
      event.hasInvalidatedColumns ||
      !event.changedCellIndicesSortedDisjoint
    ) {
      return false
    }
    if (event.changedCellIndices.length === 0) {
      continue
    }
    const first = event.firstChangedCellIndex ?? event.changedCellIndices[0]!
    const last = event.lastChangedCellIndex ?? event.changedCellIndices[event.changedCellIndices.length - 1]!
    if (first <= previousLastChangedCellIndex) {
      return false
    }
    previousLastChangedCellIndex = last
  }
  return true
}

export function captureTrackedEngineEvent(
  event: CoreTrackedEngineEvent,
  options: CaptureTrackedEngineEventOptions = {},
): TrackedEngineEvent {
  const cloneChangedCellIndices = options.cloneChangedCellIndices ?? true
  const capturedChangedCellIndices = captureChangedCellIndices(
    event.changedCellIndices,
    cloneChangedCellIndices,
    options.borrowChangedCellIndexViews === true,
  )
  const explicitChangedCount =
    typeof event.explicitChangedCount === 'number' && event.explicitChangedCount >= 0 ? event.explicitChangedCount : undefined
  if (
    event.invalidation === 'cells' &&
    (event.patches === undefined || event.patches.length === 0) &&
    event.invalidatedRanges.length === 0 &&
    event.invalidatedRows.length === 0 &&
    event.invalidatedColumns.length === 0
  ) {
    return {
      invalidation: event.invalidation,
      changedCellIndices: capturedChangedCellIndices.changedCellIndices,
      changedInputCount: event.metrics.changedInputCount,
      ...(explicitChangedCount === undefined ? {} : { explicitChangedCount }),
      changedCellIndicesSortedDisjoint: capturedChangedCellIndices.sortedDisjoint,
      ...(capturedChangedCellIndices.firstChangedCellIndex === undefined
        ? {}
        : { firstChangedCellIndex: capturedChangedCellIndices.firstChangedCellIndex }),
      ...(capturedChangedCellIndices.lastChangedCellIndex === undefined
        ? {}
        : { lastChangedCellIndex: capturedChangedCellIndices.lastChangedCellIndex }),
      ...(capturedChangedCellIndices.trustedPhysicalSheetId === undefined
        ? {}
        : { trustedPhysicalSheetId: capturedChangedCellIndices.trustedPhysicalSheetId }),
      ...(capturedChangedCellIndices.trustedSortedSliceSplit === undefined
        ? {}
        : { trustedSortedSliceSplit: capturedChangedCellIndices.trustedSortedSliceSplit }),
      hasInvalidatedRanges: false,
      hasInvalidatedRows: false,
      hasInvalidatedColumns: false,
    }
  }
  const patches = event.patches && event.patches.length > 0 ? ([...event.patches] as readonly TrackedPatch[]) : undefined
  return {
    invalidation: event.invalidation,
    changedCellIndices: capturedChangedCellIndices.changedCellIndices,
    ...(patches ? { patches } : {}),
    changedInputCount: event.metrics.changedInputCount,
    ...(explicitChangedCount === undefined ? {} : { explicitChangedCount }),
    changedCellIndicesSortedDisjoint: capturedChangedCellIndices.sortedDisjoint,
    ...(capturedChangedCellIndices.firstChangedCellIndex === undefined
      ? {}
      : { firstChangedCellIndex: capturedChangedCellIndices.firstChangedCellIndex }),
    ...(capturedChangedCellIndices.lastChangedCellIndex === undefined
      ? {}
      : { lastChangedCellIndex: capturedChangedCellIndices.lastChangedCellIndex }),
    ...(capturedChangedCellIndices.trustedPhysicalSheetId === undefined
      ? {}
      : { trustedPhysicalSheetId: capturedChangedCellIndices.trustedPhysicalSheetId }),
    ...(capturedChangedCellIndices.trustedSortedSliceSplit === undefined
      ? {}
      : { trustedSortedSliceSplit: capturedChangedCellIndices.trustedSortedSliceSplit }),
    hasInvalidatedRanges: event.invalidatedRanges.length > 0 || hasPatchKind(event, TRACKED_RANGE_INVALIDATION_PATCH_KIND),
    hasInvalidatedRows: event.invalidatedRows.length > 0 || hasPatchKind(event, TRACKED_ROW_INVALIDATION_PATCH_KIND),
    hasInvalidatedColumns: event.invalidatedColumns.length > 0 || hasPatchKind(event, TRACKED_COLUMN_INVALIDATION_PATCH_KIND),
  }
}
