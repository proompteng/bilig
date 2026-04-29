import { indexToColumn } from '@bilig/formula'
import { makeCellKey, type SpreadsheetEngine } from '@bilig/core'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkPaperCellChange } from './work-paper-types.js'

const COLUMN_LABEL_CACHE: string[] = []
const DEFERRED_TRACKED_INDEX_CHANGES = new WeakMap<readonly WorkPaperCellChange[], DeferredTrackedIndexChanges>()
const LAZY_PUBLIC_CHANGE_SOURCE_THRESHOLD = 256

export interface MaterializedTrackedIndexChanges {
  readonly changes: WorkPaperCellChange[]
  readonly ordered: boolean
}

interface TrackedIndexMaterializationOptions {
  readonly explicitChangedCount?: number
  readonly lazy?: boolean
  readonly deferLazyDetach?: boolean
  readonly trustedPhysicalSheetId?: number
  readonly trustedSortedSliceSplit?: number
}

export interface TrackedIndexChangeSource {
  readonly invalidation?: 'cells' | 'full'
  readonly changedCellIndices: readonly number[] | Uint32Array
  readonly explicitChangedCount?: number
  readonly changedCellIndicesSortedDisjoint?: boolean
  readonly firstChangedCellIndex?: number
  readonly lastChangedCellIndex?: number
  readonly patches?: readonly unknown[]
  readonly hasInvalidatedRanges?: boolean
  readonly hasInvalidatedRows?: boolean
  readonly hasInvalidatedColumns?: boolean
}

export interface MaterializedTrackedIndexChangeSources extends MaterializedTrackedIndexChanges {
  readonly usedSortedDisjointFastPath: boolean
}

interface TrackedIndexDetachOptions {
  readonly preservePositions?: boolean
}

interface DeferredTrackedIndexChanges {
  readonly forceMaterialize: () => void
  readonly detach: (options?: TrackedIndexDetachOptions) => void
}

interface DetachedPhysicalTrackedIndexChanges {
  readonly cellIndices?: Uint32Array
  readonly rows?: Uint32Array
  readonly cols?: Uint32Array
  readonly tags?: Uint8Array
  readonly constantTag?: ValueTag
  readonly numbers: Float64Array
  readonly errors?: Int32Array
  readonly stringIds?: Int32Array
  readonly strings?: readonly (string | undefined)[]
}

interface OrderedTrackedIndexSource {
  readonly changedCellIndices: readonly number[] | Uint32Array
  readonly sortedSliceSplit?: number
  readonly firstCellIndex: number
  readonly lastCellIndex: number
}

export function materializeTrackedIndexChangesWithMetadata(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  options: TrackedIndexMaterializationOptions = {},
): MaterializedTrackedIndexChanges {
  if (changedCellIndices.length === 0) {
    return { changes: [], ordered: true }
  }
  const workbook = engine.workbook
  const cellStore = workbook.cellStore
  const formatAddressCached = (row: number, col: number): string => {
    let label = COLUMN_LABEL_CACHE[col]
    if (label === undefined) {
      label = indexToColumn(col)
      COLUMN_LABEL_CACHE[col] = label
    }
    return `${label}${row + 1}`
  }
  if (options.lazy && options.trustedPhysicalSheetId !== undefined) {
    const sheet = workbook.getSheetById(options.trustedPhysicalSheetId)
    if (!sheet || sheet.structureVersion === 1) {
      const split =
        options.trustedSortedSliceSplit !== undefined &&
        options.trustedSortedSliceSplit > 0 &&
        options.trustedSortedSliceSplit < changedCellIndices.length
          ? options.trustedSortedSliceSplit
          : undefined
      return {
        changes: createLazyPhysicalTrackedIndexChanges(
          options.trustedPhysicalSheetId,
          sheet?.name ?? workbook.getSheetNameById(options.trustedPhysicalSheetId),
          cellStore,
          engine,
          changedCellIndices,
          formatAddressCached,
          split,
        ),
        ordered: true,
      }
    }
  }
  let firstSheetId: number | undefined
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const sheetId = cellStore.sheetIds[changedCellIndices[index]!]
    if (sheetId !== undefined) {
      firstSheetId = sheetId
      break
    }
  }
  if (firstSheetId === undefined) {
    return { changes: [], ordered: true }
  }
  if (options.lazy) {
    const lazyPhysicalChanges = tryCreateLazyPhysicalTrackedIndexChanges(engine, changedCellIndices, firstSheetId, formatAddressCached, {
      sortedSliceSplit: options.explicitChangedCount,
    })
    if (lazyPhysicalChanges) {
      return { changes: lazyPhysicalChanges, ordered: true }
    }
  }
  let isSingleSheetChangeSet = true
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const sheetId = cellStore.sheetIds[changedCellIndices[index]!]
    if (sheetId !== undefined && sheetId !== firstSheetId) {
      isSingleSheetChangeSet = false
      break
    }
  }
  const changes: WorkPaperCellChange[] = []
  if (isSingleSheetChangeSet) {
    const sheet = workbook.getSheetById(firstSheetId)
    const sheetName = sheet?.name ?? workbook.getSheetNameById(firstSheetId)
    const isPhysicalSheet = !sheet || sheet.structureVersion === 1
    const split = options.explicitChangedCount
    if (
      isPhysicalSheet &&
      split !== undefined &&
      split > 0 &&
      split < changedCellIndices.length &&
      isTrackedIndexSliceSorted(cellStore, changedCellIndices, 0, split) &&
      isTrackedIndexSliceSorted(cellStore, changedCellIndices, split, changedCellIndices.length)
    ) {
      if (options.lazy && trackedIndicesAllBelongToSheet(cellStore, changedCellIndices, firstSheetId)) {
        return {
          changes: createLazyPhysicalTrackedIndexChanges(
            firstSheetId,
            sheetName,
            cellStore,
            engine,
            changedCellIndices,
            formatAddressCached,
            split,
          ),
          ordered: true,
        }
      }
      let explicitIndex = 0
      let recalculatedIndex = split
      while (explicitIndex < split && recalculatedIndex < changedCellIndices.length) {
        const explicitCellIndex = changedCellIndices[explicitIndex]!
        const recalculatedCellIndex = changedCellIndices[recalculatedIndex]!
        if (compareTrackedPhysicalCellIndices(cellStore, explicitCellIndex, recalculatedCellIndex) <= 0) {
          changes.push(readPhysicalTrackedIndexChange(explicitCellIndex, firstSheetId, sheetName, cellStore, engine, formatAddressCached))
          explicitIndex += 1
        } else {
          changes.push(
            readPhysicalTrackedIndexChange(recalculatedCellIndex, firstSheetId, sheetName, cellStore, engine, formatAddressCached),
          )
          recalculatedIndex += 1
        }
      }
      while (explicitIndex < split) {
        changes.push(
          readPhysicalTrackedIndexChange(
            changedCellIndices[explicitIndex]!,
            firstSheetId,
            sheetName,
            cellStore,
            engine,
            formatAddressCached,
          ),
        )
        explicitIndex += 1
      }
      while (recalculatedIndex < changedCellIndices.length) {
        changes.push(
          readPhysicalTrackedIndexChange(
            changedCellIndices[recalculatedIndex]!,
            firstSheetId,
            sheetName,
            cellStore,
            engine,
            formatAddressCached,
          ),
        )
        recalculatedIndex += 1
      }
      return { changes, ordered: true }
    }
    if (
      options.lazy &&
      isPhysicalSheet &&
      trackedIndicesAllBelongToSheet(cellStore, changedCellIndices, firstSheetId) &&
      isTrackedIndexSliceSorted(cellStore, changedCellIndices, 0, changedCellIndices.length)
    ) {
      return {
        changes: createLazyPhysicalTrackedIndexChanges(firstSheetId, sheetName, cellStore, engine, changedCellIndices, formatAddressCached),
        ordered: true,
      }
    }
    let ordered = true
    let previousRow = -1
    let previousCol = -1
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const cellIndex = changedCellIndices[index]!
      if (cellStore.sheetIds[cellIndex] !== firstSheetId) {
        continue
      }
      let row: number
      let col: number
      if (isPhysicalSheet) {
        row = cellStore.rows[cellIndex]!
        col = cellStore.cols[cellIndex]!
      } else {
        const position = workbook.getCellPosition(cellIndex)
        /* v8 ignore next -- defensive guard for stale logical indices without visible positions. */
        if (!position) {
          continue
        }
        row = position.row
        col = position.col
      }
      if (row < previousRow || (row === previousRow && col < previousCol)) {
        ordered = false
      }
      changes.push({
        kind: 'cell',
        address: { sheet: firstSheetId, row, col },
        sheetName,
        a1: formatAddressCached(row, col),
        newValue: readTrackedCellValue(cellStore, cellIndex, engine),
      })
      previousRow = row
      previousCol = col
    }
    return { changes, ordered }
  }

  const sheetNames = new Map<number, string>()
  const physicalSheetIds = new Set<number>()
  let ordered = true
  let previousSheetOrder = -1
  let previousRow = -1
  let previousCol = -1
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    let sheetName = sheetNames.get(sheetId)
    let sheetOrder = 0
    if (sheetName === undefined) {
      const sheet = workbook.getSheetById(sheetId)
      sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
      sheetOrder = sheet?.order ?? 0
      sheetNames.set(sheetId, sheetName)
      if (!sheet || sheet.structureVersion === 1) {
        physicalSheetIds.add(sheetId)
      }
    } else {
      sheetOrder = workbook.getSheetById(sheetId)?.order ?? 0
    }
    let row: number
    let col: number
    if (physicalSheetIds.has(sheetId)) {
      row = cellStore.rows[cellIndex]!
      col = cellStore.cols[cellIndex]!
    } else {
      const position = workbook.getCellPosition(cellIndex)
      /* v8 ignore next -- defensive guard for stale logical indices without visible positions. */
      if (!position) {
        continue
      }
      row = position.row
      col = position.col
    }
    if (
      sheetOrder < previousSheetOrder ||
      (sheetOrder === previousSheetOrder && (row < previousRow || (row === previousRow && col < previousCol)))
    ) {
      ordered = false
    }
    changes.push({
      kind: 'cell',
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: formatAddressCached(row, col),
      newValue: readTrackedCellValue(cellStore, cellIndex, engine),
    })
    previousSheetOrder = sheetOrder
    previousRow = row
    previousCol = col
  }
  return { changes, ordered }
}

export function materializeTrackedIndexChanges(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  options: TrackedIndexMaterializationOptions = {},
): readonly WorkPaperCellChange[] {
  return materializeTrackedIndexChangesWithMetadata(engine, changedCellIndices, options).changes
}

export function materializeTrackedIndexChangeSourcesWithMetadata(
  engine: SpreadsheetEngine,
  sources: readonly TrackedIndexChangeSource[],
  options: TrackedIndexMaterializationOptions = {},
): MaterializedTrackedIndexChangeSources | null {
  if (sources.length === 0) {
    return { changes: [], ordered: true, usedSortedDisjointFastPath: true }
  }
  for (const source of sources) {
    if (
      source.invalidation === 'full' ||
      (source.patches !== undefined && source.patches.length > 0) ||
      source.hasInvalidatedRanges ||
      source.hasInvalidatedRows ||
      source.hasInvalidatedColumns
    ) {
      return null
    }
  }
  const lazySameSheetChanges = tryCreateLazySameSheetTrackedSourceChanges(engine, sources, {
    deferLazyDetach: options.deferLazyDetach === true,
    preferLazyPublicChanges: options.lazy === true,
  })
  if (lazySameSheetChanges) {
    return lazySameSheetChanges
  }
  if (sources.length === 1) {
    const materialized = materializeTrackedIndexChangesWithMetadata(engine, sources[0]!.changedCellIndices, {
      ...options,
      explicitChangedCount: sources[0]!.explicitChangedCount,
    })
    return {
      changes: materialized.changes,
      ordered: materialized.ordered,
      usedSortedDisjointFastPath: materialized.ordered && trackedSourceHasSortedDisjointIndices(sources[0]!),
    }
  }

  const materializedSources = Array.from({ length: sources.length }, () => undefined as MaterializedTrackedIndexChanges | undefined)
  const sheetOrders = sheetOrderLookup(engine)
  const fastChanges: WorkPaperCellChange[] = []
  let previousNumericCellIndex = -1
  let previousPublicChange: WorkPaperCellChange | undefined
  let canUseSortedDisjointFastPath = true
  for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
    const source = sources[sourceIndex]!
    if (!trackedSourceHasSortedDisjointIndices(source)) {
      canUseSortedDisjointFastPath = false
      break
    }
    if (source.changedCellIndices.length > 0) {
      const first = source.firstChangedCellIndex ?? source.changedCellIndices[0]!
      const last = source.lastChangedCellIndex ?? source.changedCellIndices[source.changedCellIndices.length - 1]!
      if (first <= previousNumericCellIndex) {
        canUseSortedDisjointFastPath = false
        break
      }
      previousNumericCellIndex = last
    }
    const materialized = materializeTrackedIndexChangesWithMetadata(engine, source.changedCellIndices, {
      ...options,
      explicitChangedCount: source.explicitChangedCount,
    })
    materializedSources[sourceIndex] = materialized
    if (!materialized.ordered) {
      canUseSortedDisjointFastPath = false
      break
    }
    for (let changeIndex = 0; changeIndex < materialized.changes.length; changeIndex += 1) {
      const change = materialized.changes[changeIndex]!
      if (previousPublicChange !== undefined) {
        const comparison = compareTrackedCellChanges(previousPublicChange, change, sheetOrders)
        if (comparison >= 0) {
          canUseSortedDisjointFastPath = false
          break
        }
      }
      fastChanges.push(change)
      previousPublicChange = change
    }
    if (!canUseSortedDisjointFastPath) {
      break
    }
  }
  if (canUseSortedDisjointFastPath) {
    return { changes: fastChanges, ordered: true, usedSortedDisjointFastPath: true }
  }
  return materializeTrackedIndexChangeSourcesGeneric(engine, sources, materializedSources, options, sheetOrders)
}

export function forceMaterializeTrackedIndexChanges(changes: readonly WorkPaperCellChange[]): boolean {
  const deferred = DEFERRED_TRACKED_INDEX_CHANGES.get(changes)
  if (deferred === undefined) {
    return false
  }
  deferred.forceMaterialize()
  DEFERRED_TRACKED_INDEX_CHANGES.delete(changes)
  return true
}

export function detachTrackedIndexChanges(changes: readonly WorkPaperCellChange[], options: TrackedIndexDetachOptions = {}): boolean {
  const deferred = DEFERRED_TRACKED_INDEX_CHANGES.get(changes)
  if (deferred === undefined) {
    return false
  }
  deferred.detach(options)
  return true
}

export function hasDeferredTrackedIndexChanges(changes: readonly WorkPaperCellChange[]): boolean {
  return DEFERRED_TRACKED_INDEX_CHANGES.has(changes)
}

function readTrackedCellValue(cellStore: TrackedCellStore, cellIndex: number, engine: SpreadsheetEngine): CellValue {
  const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  switch (tag) {
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: cellStore.numbers[cellIndex] ?? 0 }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (cellStore.numbers[cellIndex] ?? 0) !== 0 }
    case ValueTag.String:
      return cellStore.getValue(cellIndex, (stringId) => engine.strings.get(stringId))
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: cellStore.errors[cellIndex]! }
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty }
  }
}

type TrackedCellStore = SpreadsheetEngine['workbook']['cellStore']
type SheetOrderLookup = ReadonlyMap<number, number>

function materializeTrackedIndexChangeSourcesGeneric(
  engine: SpreadsheetEngine,
  sources: readonly TrackedIndexChangeSource[],
  materializedSources: readonly (MaterializedTrackedIndexChanges | undefined)[],
  options: TrackedIndexMaterializationOptions,
  sheetOrders: SheetOrderLookup,
): MaterializedTrackedIndexChangeSources {
  const latestChangesByKey = new Map<number, WorkPaperCellChange>()
  for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
    const source = sources[sourceIndex]!
    const materialized =
      materializedSources[sourceIndex] ??
      materializeTrackedIndexChangesWithMetadata(engine, source.changedCellIndices, {
        ...options,
        explicitChangedCount: source.explicitChangedCount,
      })
    for (let changeIndex = 0; changeIndex < materialized.changes.length; changeIndex += 1) {
      const change = materialized.changes[changeIndex]!
      const key = makeCellKey(change.address.sheet, change.address.row, change.address.col)
      latestChangesByKey.delete(key)
      latestChangesByKey.set(key, change)
    }
  }
  return {
    changes: orderTrackedCellChanges([...latestChangesByKey.values()], sheetOrders),
    ordered: true,
    usedSortedDisjointFastPath: false,
  }
}

function tryCreateLazySameSheetTrackedSourceChanges(
  engine: SpreadsheetEngine,
  sources: readonly TrackedIndexChangeSource[],
  options: { readonly deferLazyDetach: boolean; readonly preferLazyPublicChanges: boolean },
): MaterializedTrackedIndexChangeSources | null {
  let totalLength = 0
  for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
    totalLength += sources[sourceIndex]!.changedCellIndices.length
  }
  if (!options.preferLazyPublicChanges && totalLength < LAZY_PUBLIC_CHANGE_SOURCE_THRESHOLD) {
    return null
  }

  const workbook = engine.workbook
  const cellStore = workbook.cellStore
  const orderedSources: OrderedTrackedIndexSource[] = []
  let sheetId: number | undefined
  let previousNumericCellIndex = -1
  let previousLastCellIndex: number | undefined
  for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
    const source = sources[sourceIndex]!
    if (!trackedSourceHasSortedDisjointIndices(source)) {
      return null
    }
    if (source.changedCellIndices.length === 0) {
      continue
    }

    const firstNumericCellIndex = source.firstChangedCellIndex ?? source.changedCellIndices[0]!
    const lastNumericCellIndex = source.lastChangedCellIndex ?? source.changedCellIndices[source.changedCellIndices.length - 1]!
    if (firstNumericCellIndex <= previousNumericCellIndex) {
      return null
    }
    previousNumericCellIndex = lastNumericCellIndex

    const sourceSheetId = cellStore.sheetIds[source.changedCellIndices[0]!]
    if (sourceSheetId === undefined) {
      return null
    }
    if (sheetId === undefined) {
      sheetId = sourceSheetId
    } else if (sourceSheetId !== sheetId) {
      return null
    }
    const sheet = workbook.getSheetById(sourceSheetId)
    if (sheet && sheet.structureVersion !== 1) {
      return null
    }
    if (!trackedIndicesAllBelongToSheet(cellStore, source.changedCellIndices, sourceSheetId)) {
      return null
    }

    const orderedSource = orderedTrackedIndexSource(cellStore, source)
    if (orderedSource === null) {
      return null
    }
    if (
      previousLastCellIndex !== undefined &&
      compareTrackedPhysicalCellIndices(cellStore, previousLastCellIndex, orderedSource.firstCellIndex) >= 0
    ) {
      return null
    }
    previousLastCellIndex = orderedSource.lastCellIndex
    orderedSources.push(orderedSource)
  }
  if (sheetId === undefined) {
    return { changes: [], ordered: true, usedSortedDisjointFastPath: true }
  }

  const orderedCellIndices = new Uint32Array(totalLength)
  let offset = 0
  for (let sourceIndex = 0; sourceIndex < orderedSources.length; sourceIndex += 1) {
    const source = orderedSources[sourceIndex]!
    offset = copyOrderedTrackedCellIndices(cellStore, source.changedCellIndices, orderedCellIndices, offset, source.sortedSliceSplit)
  }

  const sheet = workbook.getSheetById(sheetId)
  const sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
  const changes = createLazyPhysicalTrackedIndexChanges(sheetId, sheetName, cellStore, engine, orderedCellIndices, formatTrackedAddress)
  if (!options.deferLazyDetach) {
    detachTrackedIndexChanges(changes)
  }
  return { changes, ordered: true, usedSortedDisjointFastPath: true }
}

function orderedTrackedIndexSource(cellStore: TrackedCellStore, source: TrackedIndexChangeSource): OrderedTrackedIndexSource | null {
  const length = source.changedCellIndices.length
  const split = source.explicitChangedCount
  if (split !== undefined && split > 0 && split < length) {
    if (
      !isTrackedIndexSliceSorted(cellStore, source.changedCellIndices, 0, split) ||
      !isTrackedIndexSliceSorted(cellStore, source.changedCellIndices, split, length)
    ) {
      return null
    }
    const firstLeft = source.changedCellIndices[0]!
    const firstRight = source.changedCellIndices[split]!
    const lastLeft = source.changedCellIndices[split - 1]!
    const lastRight = source.changedCellIndices[length - 1]!
    return {
      changedCellIndices: source.changedCellIndices,
      sortedSliceSplit: split,
      firstCellIndex: compareTrackedPhysicalCellIndices(cellStore, firstLeft, firstRight) <= 0 ? firstLeft : firstRight,
      lastCellIndex: compareTrackedPhysicalCellIndices(cellStore, lastLeft, lastRight) >= 0 ? lastLeft : lastRight,
    }
  }
  if (!isTrackedIndexSliceSorted(cellStore, source.changedCellIndices, 0, length)) {
    return null
  }
  return {
    changedCellIndices: source.changedCellIndices,
    firstCellIndex: source.changedCellIndices[0]!,
    lastCellIndex: source.changedCellIndices[length - 1]!,
  }
}

function trackedSourceHasSortedDisjointIndices(source: TrackedIndexChangeSource): boolean {
  if (source.changedCellIndicesSortedDisjoint !== undefined) {
    return source.changedCellIndicesSortedDisjoint
  }
  let previous = -1
  for (let index = 0; index < source.changedCellIndices.length; index += 1) {
    const cellIndex = source.changedCellIndices[index]!
    if (!Number.isInteger(cellIndex) || cellIndex < 0 || cellIndex <= previous) {
      return false
    }
    previous = cellIndex
  }
  return true
}

function sheetOrderLookup(engine: SpreadsheetEngine): SheetOrderLookup {
  const sheetOrders = new Map<number, number>()
  engine.workbook.sheetsById.forEach((sheet, sheetId) => {
    sheetOrders.set(sheetId, sheet.order)
  })
  return sheetOrders
}

function compareTrackedCellChanges(left: WorkPaperCellChange, right: WorkPaperCellChange, sheetOrders: SheetOrderLookup): number {
  return (
    (sheetOrders.get(left.address.sheet) ?? 0) - (sheetOrders.get(right.address.sheet) ?? 0) ||
    left.address.row - right.address.row ||
    left.address.col - right.address.col
  )
}

function orderTrackedCellChanges(changes: WorkPaperCellChange[], sheetOrders: SheetOrderLookup): WorkPaperCellChange[] {
  if (changes.length < 2) {
    return changes
  }
  if (isTrackedCellChangeSliceSorted(changes, sheetOrders, 0, changes.length)) {
    return changes
  }
  if (isTrackedCellChangeSliceReverseSorted(changes, sheetOrders, 0, changes.length)) {
    return changes.toReversed()
  }
  return changes.toSorted((left, right) => compareTrackedCellChanges(left, right, sheetOrders))
}

function isTrackedCellChangeSliceSorted(
  changes: readonly WorkPaperCellChange[],
  sheetOrders: SheetOrderLookup,
  start: number,
  end: number,
): boolean {
  for (let index = start + 1; index < end; index += 1) {
    if (compareTrackedCellChanges(changes[index - 1]!, changes[index]!, sheetOrders) > 0) {
      return false
    }
  }
  return true
}

function isTrackedCellChangeSliceReverseSorted(
  changes: readonly WorkPaperCellChange[],
  sheetOrders: SheetOrderLookup,
  start: number,
  end: number,
): boolean {
  for (let index = start + 1; index < end; index += 1) {
    if (compareTrackedCellChanges(changes[index - 1]!, changes[index]!, sheetOrders) < 0) {
      return false
    }
  }
  return true
}

function compareTrackedPhysicalCellIndices(cellStore: TrackedCellStore, leftCellIndex: number, rightCellIndex: number): number {
  return (
    (cellStore.rows[leftCellIndex] ?? 0) - (cellStore.rows[rightCellIndex] ?? 0) ||
    (cellStore.cols[leftCellIndex] ?? 0) - (cellStore.cols[rightCellIndex] ?? 0)
  )
}

function isTrackedIndexSliceSorted(
  cellStore: TrackedCellStore,
  changedCellIndices: readonly number[] | Uint32Array,
  start: number,
  end: number,
): boolean {
  for (let index = start + 1; index < end; index += 1) {
    if (compareTrackedPhysicalCellIndices(cellStore, changedCellIndices[index - 1]!, changedCellIndices[index]!) > 0) {
      return false
    }
  }
  return true
}

function formatTrackedAddress(row: number, col: number): string {
  let label = COLUMN_LABEL_CACHE[col]
  if (label === undefined) {
    label = indexToColumn(col)
    COLUMN_LABEL_CACHE[col] = label
  }
  return `${label}${row + 1}`
}

function readPhysicalTrackedIndexChange(
  cellIndex: number,
  sheetId: number,
  sheetName: string,
  cellStore: TrackedCellStore,
  engine: SpreadsheetEngine,
  formatAddressCached: (row: number, col: number) => string,
): WorkPaperCellChange {
  const row = cellStore.rows[cellIndex]!
  const col = cellStore.cols[cellIndex]!
  return {
    kind: 'cell',
    address: { sheet: sheetId, row, col },
    sheetName,
    a1: formatAddressCached(row, col),
    newValue: readTrackedCellValue(cellStore, cellIndex, engine),
  }
}

function trackedIndicesAllBelongToSheet(
  cellStore: TrackedCellStore,
  changedCellIndices: readonly number[] | Uint32Array,
  sheetId: number,
): boolean {
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    if (cellStore.sheetIds[changedCellIndices[index]!] !== sheetId) {
      return false
    }
  }
  return true
}

function createLazyPhysicalTrackedIndexChanges(
  sheetId: number,
  sheetName: string,
  cellStore: TrackedCellStore,
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  formatAddressCached: (row: number, col: number) => string,
  sortedSliceSplit?: number,
): WorkPaperCellChange[] {
  const changedCellIndicesForMaterialization =
    changedCellIndices instanceof Uint32Array ? changedCellIndices : copyTrackedCellIndices(changedCellIndices)
  const length = changedCellIndicesForMaterialization.length
  const cache: WorkPaperCellChange[] = []
  cache.length = length
  let detached: DetachedPhysicalTrackedIndexChanges | undefined
  let fullyMaterialized = false
  let orderedCellIndices: Uint32Array | undefined
  const hasMaterializedIndex = (index: number): boolean => Object.prototype.hasOwnProperty.call(cache, index)
  const orderedCellIndexAt = (index: number): number => {
    if (sortedSliceSplit === undefined) {
      return changedCellIndicesForMaterialization[index]!
    }
    if (orderedCellIndices === undefined) {
      orderedCellIndices = new Uint32Array(length)
      copyOrderedTrackedCellIndices(cellStore, changedCellIndicesForMaterialization, orderedCellIndices, 0, sortedSliceSplit)
    }
    return orderedCellIndices[index]!
  }
  const forEachOrderedCellIndex = (fn: (index: number, cellIndex: number) => void): void => {
    if (sortedSliceSplit === undefined) {
      for (let index = 0; index < length; index += 1) {
        fn(index, changedCellIndicesForMaterialization[index]!)
      }
      return
    }
    if (orderedCellIndices === undefined) {
      orderedCellIndices = new Uint32Array(length)
      copyOrderedTrackedCellIndices(cellStore, changedCellIndicesForMaterialization, orderedCellIndices, 0, sortedSliceSplit)
    }
    for (let index = 0; index < length; index += 1) {
      fn(index, orderedCellIndices[index]!)
    }
  }
  const materialize = (index: number): WorkPaperCellChange => {
    if (hasMaterializedIndex(index)) {
      return cache[index]!
    }
    if (detached) {
      cache[index] = readDetachedPhysicalTrackedIndexChange(index, sheetId, sheetName, cellStore, detached, formatAddressCached)
      return cache[index]
    }
    cache[index] = readPhysicalTrackedIndexChange(orderedCellIndexAt(index), sheetId, sheetName, cellStore, engine, formatAddressCached)
    return cache[index]
  }
  const detach = (options: TrackedIndexDetachOptions = {}): void => {
    if (fullyMaterialized) {
      return
    }
    const preservePositions = options.preservePositions ?? true
    if (detached !== undefined) {
      if (preservePositions && detached.rows === undefined) {
        if (detached.cellIndices === undefined) {
          return
        }
        const rows = new Uint32Array(length)
        const cols = new Uint32Array(length)
        for (let index = 0; index < length; index += 1) {
          const cellIndex = detached.cellIndices[index]!
          rows[index] = cellStore.rows[cellIndex] ?? 0
          cols[index] = cellStore.cols[cellIndex] ?? 0
        }
        detached = { ...detached, rows, cols }
      }
      return
    }
    const cellIndices = preservePositions ? undefined : new Uint32Array(length)
    const rows = preservePositions ? new Uint32Array(length) : undefined
    const cols = preservePositions ? new Uint32Array(length) : undefined
    const numbers = new Float64Array(length)
    let tags: Uint8Array | undefined
    let constantTag: ValueTag | undefined
    let errors: Int32Array | undefined
    let stringIds: Int32Array | undefined
    let strings: (string | undefined)[] | undefined
    const ensureTags = (currentIndex: number): Uint8Array => {
      if (tags !== undefined) {
        return tags
      }
      const nextTags = new Uint8Array(length)
      nextTags.fill(constantTag ?? ValueTag.Empty, 0, currentIndex)
      tags = nextTags
      return nextTags
    }
    forEachOrderedCellIndex((index, cellIndex) => {
      if (cellIndices !== undefined) {
        cellIndices[index] = cellIndex
      }
      if (rows !== undefined && cols !== undefined) {
        rows[index] = cellStore.rows[cellIndex] ?? 0
        cols[index] = cellStore.cols[cellIndex] ?? 0
      }
      const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
      if (tags !== undefined) {
        tags[index] = tag
      } else if (constantTag === undefined) {
        constantTag = tag
      } else if (tag !== constantTag) {
        ensureTags(index)[index] = tag
      }
      switch (tag) {
        case ValueTag.Number:
        case ValueTag.Boolean:
          numbers[index] = cellStore.numbers[cellIndex] ?? 0
          break
        case ValueTag.String:
          {
            stringIds ??= new Int32Array(length)
            strings ??= []
            stringIds[index] = cellStore.stringIds[cellIndex] ?? 0
            const value = cellStore.getValue(cellIndex, (stringId) => engine.strings.get(stringId))
            strings[index] = value.tag === ValueTag.String ? value.value : ''
          }
          break
        case ValueTag.Error:
          errors ??= new Int32Array(length)
          errors[index] = cellStore.errors[cellIndex] ?? ErrorCode.None
          break
        case ValueTag.Empty:
        default:
          break
      }
    })
    detached = {
      ...(cellIndices === undefined ? {} : { cellIndices }),
      ...(rows !== undefined && cols !== undefined ? { rows, cols } : {}),
      ...(tags !== undefined ? { tags } : { constantTag: constantTag ?? ValueTag.Empty }),
      numbers,
      ...(errors === undefined ? {} : { errors }),
      ...(stringIds === undefined ? {} : { stringIds }),
      ...(strings === undefined ? {} : { strings }),
    }
  }
  const forceMaterialize = (): void => {
    if (fullyMaterialized) {
      return
    }
    if (detached !== undefined) {
      for (let index = 0; index < length; index += 1) {
        if (!hasMaterializedIndex(index)) {
          cache[index] = readDetachedPhysicalTrackedIndexChange(index, sheetId, sheetName, cellStore, detached, formatAddressCached)
        }
      }
      fullyMaterialized = true
      return
    }
    for (let index = 0; index < length; index += 1) {
      if (!hasMaterializedIndex(index)) {
        cache[index] = readPhysicalTrackedIndexChange(orderedCellIndexAt(index), sheetId, sheetName, cellStore, engine, formatAddressCached)
      }
    }
    fullyMaterialized = true
  }
  const numericIndexOf = (property: string | symbol): number | undefined => {
    if (typeof property !== 'string' || property.length === 0) {
      return undefined
    }
    const index = Number(property)
    return Number.isInteger(index) && index >= 0 && index < length && String(index) === property ? index : undefined
  }
  const proxy = new Proxy(cache, {
    get(target, property, receiver) {
      const index = numericIndexOf(property)
      return index === undefined ? Reflect.get(target, property, receiver) : materialize(index)
    },
    getOwnPropertyDescriptor(target, property) {
      const index = numericIndexOf(property)
      if (index === undefined) {
        return Reflect.getOwnPropertyDescriptor(target, property)
      }
      return {
        configurable: true,
        enumerable: true,
        value: materialize(index),
        writable: true,
      }
    },
    has(target, property) {
      return numericIndexOf(property) !== undefined || Reflect.has(target, property)
    },
    ownKeys(target) {
      return [
        ...Array.from({ length }, (_value, index) => String(index)),
        ...Reflect.ownKeys(target).filter((key) => typeof key !== 'string' || numericIndexOf(key) === undefined),
      ]
    },
  })
  DEFERRED_TRACKED_INDEX_CHANGES.set(proxy, { forceMaterialize, detach })
  return proxy
}

function tryCreateLazyPhysicalTrackedIndexChanges(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  sheetId: number,
  formatAddressCached: (row: number, col: number) => string,
  options: { readonly sortedSliceSplit?: number } = {},
): WorkPaperCellChange[] | null {
  const workbook = engine.workbook
  const sheet = workbook.getSheetById(sheetId)
  if (sheet && sheet.structureVersion !== 1) {
    return null
  }
  const cellStore = workbook.cellStore
  const length = changedCellIndices.length
  const split = options.sortedSliceSplit
  const sortedSliceSplit = split !== undefined && split > 0 && split < length ? split : undefined
  let previousRow = -1
  let previousCol = -1
  let previousRightRow = -1
  let previousRightCol = -1
  for (let index = 0; index < length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    if (cellStore.sheetIds[cellIndex] !== sheetId) {
      return null
    }
    const row = cellStore.rows[cellIndex] ?? 0
    const col = cellStore.cols[cellIndex] ?? 0
    if (sortedSliceSplit !== undefined && index >= sortedSliceSplit) {
      if (row < previousRightRow || (row === previousRightRow && col < previousRightCol)) {
        return null
      }
      previousRightRow = row
      previousRightCol = col
      continue
    }
    if (row < previousRow || (row === previousRow && col < previousCol)) {
      return null
    }
    previousRow = row
    previousCol = col
  }
  const sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
  return createLazyPhysicalTrackedIndexChanges(
    sheetId,
    sheetName,
    cellStore,
    engine,
    changedCellIndices,
    formatAddressCached,
    sortedSliceSplit,
  )
}

function copyOrderedTrackedCellIndices(
  cellStore: TrackedCellStore,
  changedCellIndices: readonly number[] | Uint32Array,
  target: Uint32Array,
  offset: number,
  sortedSliceSplit?: number,
): number {
  if (sortedSliceSplit === undefined) {
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      target[offset + index] = changedCellIndices[index]!
    }
    return offset + changedCellIndices.length
  }

  let explicitIndex = 0
  let recalculatedIndex = sortedSliceSplit
  let outputIndex = offset
  while (explicitIndex < sortedSliceSplit && recalculatedIndex < changedCellIndices.length) {
    const explicitCellIndex = changedCellIndices[explicitIndex]!
    const recalculatedCellIndex = changedCellIndices[recalculatedIndex]!
    if (compareTrackedPhysicalCellIndices(cellStore, explicitCellIndex, recalculatedCellIndex) <= 0) {
      target[outputIndex] = explicitCellIndex
      explicitIndex += 1
    } else {
      target[outputIndex] = recalculatedCellIndex
      recalculatedIndex += 1
    }
    outputIndex += 1
  }
  while (explicitIndex < sortedSliceSplit) {
    target[outputIndex] = changedCellIndices[explicitIndex]!
    explicitIndex += 1
    outputIndex += 1
  }
  while (recalculatedIndex < changedCellIndices.length) {
    target[outputIndex] = changedCellIndices[recalculatedIndex]!
    recalculatedIndex += 1
    outputIndex += 1
  }
  return outputIndex
}

function copyTrackedCellIndices(changedCellIndices: readonly number[] | Uint32Array): Uint32Array {
  return Uint32Array.from(changedCellIndices)
}

function readDetachedPhysicalTrackedIndexChange(
  index: number,
  sheetId: number,
  sheetName: string,
  cellStore: TrackedCellStore,
  detached: DetachedPhysicalTrackedIndexChanges,
  formatAddressCached: (row: number, col: number) => string,
): WorkPaperCellChange {
  const cellIndex = detached.cellIndices?.[index]
  const row = detached.rows?.[index] ?? (cellIndex === undefined ? 0 : (cellStore.rows[cellIndex] ?? 0))
  const col = detached.cols?.[index] ?? (cellIndex === undefined ? 0 : (cellStore.cols[cellIndex] ?? 0))
  return {
    kind: 'cell',
    address: { sheet: sheetId, row, col },
    sheetName,
    a1: formatAddressCached(row, col),
    newValue: readDetachedCellValue(detached, index),
  }
}

function readDetachedCellValue(detached: DetachedPhysicalTrackedIndexChanges, index: number): CellValue {
  const tag = (detached.tags?.[index] ?? detached.constantTag ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: detached.numbers[index] ?? 0 }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: (detached.numbers[index] ?? 0) !== 0 }
    case ValueTag.String:
      return { tag: ValueTag.String, value: detached.strings?.[index] ?? '', stringId: detached.stringIds?.[index] ?? 0 }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: detached.errors?.[index] ?? ErrorCode.None }
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty }
  }
}
