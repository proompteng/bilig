import { indexToColumn } from '@bilig/formula'
import { makeCellKey, type SpreadsheetEngine } from '@bilig/core/headless-runtime'
import {
  compareTrackedPhysicalCellIndices,
  copyOrderedTrackedCellIndices,
  createLazyPhysicalTrackedIndexChanges,
  createPrefixedLazyTrackedIndexChanges,
  detachTrackedIndexChanges,
  formatTrackedAddress,
  isTrackedIndexSliceSorted,
  readPhysicalTrackedIndexChange,
  trackedIndicesAllBelongToSheet,
  trackedIndicesUseVisiblePhysicalPositions,
  tryCreateLazyPhysicalTrackedIndexChanges,
} from './tracked-cell-lazy-physical-changes.js'
import { readTrackedCellValue } from './tracked-cell-value-read.js'
import { trackedSourceHasSortedDisjointIndices } from './tracked-index-source-order.js'
import type { WorkPaperCellChange } from './work-paper-types.js'

const COLUMN_LABEL_CACHE: string[] = []
const LAZY_PUBLIC_CHANGE_SOURCE_THRESHOLD = 256

export {
  detachTrackedIndexChanges,
  forceMaterializeTrackedIndexChanges,
  hasDeferredTrackedIndexChanges,
} from './tracked-cell-lazy-physical-changes.js'

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
  readonly trustedPhysicalSheetId?: number
  readonly trustedSortedSliceSplit?: number
  readonly patches?: readonly unknown[]
  readonly hasInvalidatedRanges?: boolean
  readonly hasInvalidatedRows?: boolean
  readonly hasInvalidatedColumns?: boolean
}

export interface MaterializedTrackedIndexChangeSources extends MaterializedTrackedIndexChanges {
  readonly usedSortedDisjointFastPath: boolean
}

interface OrderedTrackedIndexSource {
  readonly changedCellIndices: readonly number[] | Uint32Array
  readonly sortedSliceSplit?: number
  readonly firstCellIndex: number
  readonly lastCellIndex: number
}

function withSourceExplicitChangedCount(
  options: TrackedIndexMaterializationOptions,
  explicitChangedCount: number | undefined,
): TrackedIndexMaterializationOptions {
  return explicitChangedCount === undefined ? options : { ...options, explicitChangedCount }
}

function withTrackedSourceOptions(
  options: TrackedIndexMaterializationOptions,
  source: TrackedIndexChangeSource,
): TrackedIndexMaterializationOptions {
  return {
    ...withSourceExplicitChangedCount(options, source.explicitChangedCount),
    ...(source.trustedPhysicalSheetId === undefined ? {} : { trustedPhysicalSheetId: source.trustedPhysicalSheetId }),
    ...(source.trustedSortedSliceSplit === undefined ? {} : { trustedSortedSliceSplit: source.trustedSortedSliceSplit }),
  }
}

function trackedSourceUsesTrustedPhysicalPositions(source: TrackedIndexChangeSource, sheetId: number): boolean {
  return source.trustedPhysicalSheetId === sheetId
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
    const lazyPhysicalChanges = tryCreateLazyPhysicalTrackedIndexChanges(
      engine,
      changedCellIndices,
      firstSheetId,
      formatAddressCached,
      options.explicitChangedCount === undefined ? {} : { sortedSliceSplit: options.explicitChangedCount },
    )
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
  if (options.lazy && options.explicitChangedCount === 1 && changedCellIndices.length > LAZY_PUBLIC_CHANGE_SOURCE_THRESHOLD) {
    const tailStart = 1
    const tailFirstCellIndex = changedCellIndices[tailStart]
    const tailSheetId = tailFirstCellIndex === undefined ? undefined : cellStore.sheetIds[tailFirstCellIndex]
    const firstSheet = workbook.getSheetById(firstSheetId)
    const tailSheet = tailSheetId === undefined ? undefined : workbook.getSheetById(tailSheetId)
    if (tailSheetId !== undefined && tailSheetId !== firstSheetId && firstSheet && tailSheet && firstSheet.order <= tailSheet.order) {
      const tailIndices =
        changedCellIndices instanceof Uint32Array ? changedCellIndices.subarray(tailStart) : changedCellIndices.slice(tailStart)
      const tailChanges = tryCreateLazyPhysicalTrackedIndexChanges(engine, tailIndices, tailSheetId, formatAddressCached)
      if (tailChanges !== null) {
        const prefix = materializeTrackedIndexChangesWithMetadata(
          engine,
          changedCellIndices instanceof Uint32Array ? changedCellIndices.subarray(0, tailStart) : changedCellIndices.slice(0, tailStart),
        )
        if (prefix.changes.length === tailStart) {
          return {
            changes: createPrefixedLazyTrackedIndexChanges(prefix.changes, tailChanges),
            ordered: true,
          }
        }
      }
    }
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
  const cellChangeSources = sources.filter(sourceHasCellChanges)
  if (cellChangeSources.length === 0) {
    return { changes: [], ordered: true, usedSortedDisjointFastPath: true }
  }
  for (const source of cellChangeSources) {
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
  if (options.lazy === true && cellChangeSources.length === 1) {
    const source = cellChangeSources[0]!
    if (source.trustedPhysicalSheetId !== undefined) {
      const materialized = materializeTrackedIndexChangesWithMetadata(
        engine,
        source.changedCellIndices,
        withTrackedSourceOptions(options, source),
      )
      return {
        changes: materialized.changes,
        ordered: materialized.ordered,
        usedSortedDisjointFastPath: true,
      }
    }
  }
  const lazySameSheetChanges = tryCreateLazySameSheetTrackedSourceChanges(engine, cellChangeSources, {
    deferLazyDetach: options.deferLazyDetach === true,
    preferLazyPublicChanges: options.lazy === true,
  })
  if (lazySameSheetChanges) {
    return lazySameSheetChanges
  }
  const orderedLazySameSheetChanges = tryCreateOrderedLazySameSheetTrackedSourceChanges(engine, cellChangeSources, {
    deferLazyDetach: options.deferLazyDetach === true,
    preferLazyPublicChanges: options.lazy === true,
  })
  if (orderedLazySameSheetChanges) {
    return orderedLazySameSheetChanges
  }
  if (cellChangeSources.length === 1) {
    const source = cellChangeSources[0]!
    const materialized = materializeTrackedIndexChangesWithMetadata(
      engine,
      source.changedCellIndices,
      withTrackedSourceOptions(options, source),
    )
    return {
      changes: materialized.changes,
      ordered: materialized.ordered,
      usedSortedDisjointFastPath: materialized.ordered && trackedSourceHasSortedDisjointIndices(source),
    }
  }

  const materializedSources = Array.from(
    { length: cellChangeSources.length },
    () => undefined as MaterializedTrackedIndexChanges | undefined,
  )
  const sheetOrders = sheetOrderLookup(engine)
  const fastChanges: WorkPaperCellChange[] = []
  let previousNumericCellIndex = -1
  let previousPublicChange: WorkPaperCellChange | undefined
  let canUseSortedDisjointFastPath = true
  for (let sourceIndex = 0; sourceIndex < cellChangeSources.length; sourceIndex += 1) {
    const source = cellChangeSources[sourceIndex]!
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
    const materialized = materializeTrackedIndexChangesWithMetadata(
      engine,
      source.changedCellIndices,
      withTrackedSourceOptions(options, source),
    )
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
  return materializeTrackedIndexChangeSourcesGeneric(engine, cellChangeSources, materializedSources, options, sheetOrders)
}

function sourceHasCellChanges(source: TrackedIndexChangeSource): boolean {
  return (
    source.changedCellIndices.length > 0 ||
    source.patches?.some((patch) => typeof patch === 'object' && patch !== null && Reflect.get(patch, 'kind') === 'cell') === true
  )
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
      materializeTrackedIndexChangesWithMetadata(engine, source.changedCellIndices, withTrackedSourceOptions(options, source))
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
    if (
      sheet &&
      sheet.structureVersion !== 1 &&
      !trackedIndicesUseVisiblePhysicalPositions(workbook, source.changedCellIndices, sourceSheetId)
    ) {
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

function tryCreateOrderedLazySameSheetTrackedSourceChanges(
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
  let sheetId: number | undefined
  let mustDetachBeforeReturn = false
  const latestCellIndexByAddress = new Map<number, number>()
  for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
    const source = sources[sourceIndex]!
    for (let index = 0; index < source.changedCellIndices.length; index += 1) {
      const cellIndex = source.changedCellIndices[index]!
      const sourceSheetId = cellStore.sheetIds[cellIndex]
      if (sourceSheetId === undefined) {
        return null
      }
      if (sheetId === undefined) {
        sheetId = sourceSheetId
        const sheet = workbook.getSheetById(sheetId)
        mustDetachBeforeReturn =
          sheet !== undefined &&
          sheet.structureVersion !== 1 &&
          !trackedSourceUsesTrustedPhysicalPositions(source, sheetId) &&
          !trackedIndicesUseVisiblePhysicalPositions(workbook, source.changedCellIndices, sheetId)
      } else if (sourceSheetId !== sheetId) {
        return null
      }
      const row = cellStore.rows[cellIndex]
      const col = cellStore.cols[cellIndex]
      if (row === undefined || col === undefined) {
        return null
      }
      latestCellIndexByAddress.set(makeCellKey(sourceSheetId, row, col), cellIndex)
    }
  }
  if (sheetId === undefined) {
    return { changes: [], ordered: true, usedSortedDisjointFastPath: true }
  }
  const sheet = workbook.getSheetById(sheetId)
  if (sheet && sheet.structureVersion !== 1 && !mustDetachBeforeReturn) {
    for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
      const source = sources[sourceIndex]!
      if (
        !trackedSourceUsesTrustedPhysicalPositions(source, sheetId) &&
        !trackedIndicesUseVisiblePhysicalPositions(workbook, source.changedCellIndices, sheetId)
      ) {
        mustDetachBeforeReturn = true
        break
      }
    }
  }

  const orderedCellIndices = Uint32Array.from(latestCellIndexByAddress.values())
  orderedCellIndices.sort((left, right) => compareTrackedPhysicalCellIndices(cellStore, left, right))
  const changes = createLazyPhysicalTrackedIndexChanges(
    sheetId,
    sheet?.name ?? workbook.getSheetNameById(sheetId),
    cellStore,
    engine,
    orderedCellIndices,
    formatTrackedAddress,
  )
  if (mustDetachBeforeReturn || !options.deferLazyDetach) {
    detachTrackedIndexChanges(changes)
  }
  return { changes, ordered: true, usedSortedDisjointFastPath: false }
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
