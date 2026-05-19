import { indexToColumn } from '@bilig/formula'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'
import {
  readDetachedPhysicalTrackedIndexChange,
  readTrackedCellValue,
  type DetachedPhysicalTrackedIndexChanges,
} from './tracked-cell-value-read.js'
import type { WorkPaperCellChange } from './work-paper-types.js'

const COLUMN_LABEL_CACHE: string[] = []
const DEFERRED_TRACKED_INDEX_CHANGES = new WeakMap<readonly WorkPaperCellChange[], DeferredTrackedIndexChanges>()

export interface TrackedIndexDetachOptions {
  readonly preservePositions?: boolean
}

interface DeferredTrackedIndexChanges {
  readonly forceMaterialize: () => void
  readonly detach: (options?: TrackedIndexDetachOptions) => void
}

type TrackedCellStore = SpreadsheetEngine['workbook']['cellStore']

export function compareTrackedPhysicalCellIndices(cellStore: TrackedCellStore, leftCellIndex: number, rightCellIndex: number): number {
  return (
    (cellStore.rows[leftCellIndex] ?? 0) - (cellStore.rows[rightCellIndex] ?? 0) ||
    (cellStore.cols[leftCellIndex] ?? 0) - (cellStore.cols[rightCellIndex] ?? 0)
  )
}

export function isTrackedIndexSliceSorted(
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

export function formatTrackedAddress(row: number, col: number): string {
  let label = COLUMN_LABEL_CACHE[col]
  if (label === undefined) {
    label = indexToColumn(col)
    COLUMN_LABEL_CACHE[col] = label
  }
  return `${label}${row + 1}`
}

export function readPhysicalTrackedIndexChange(
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

export function trackedIndicesAllBelongToSheet(
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

export function trackedIndicesUseVisiblePhysicalPositions(
  workbook: SpreadsheetEngine['workbook'],
  changedCellIndices: readonly number[] | Uint32Array,
  sheetId: number,
): boolean {
  const cellStore = workbook.cellStore
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    if (cellStore.sheetIds[cellIndex] !== sheetId) {
      return false
    }
    const row = cellStore.rows[cellIndex]
    const col = cellStore.cols[cellIndex]
    if (row === undefined || col === undefined) {
      return false
    }
    const position = workbook.getCellPosition(cellIndex)
    if (!position || position.row !== row || position.col !== col) {
      return false
    }
  }
  return true
}

export function createLazyPhysicalTrackedIndexChanges(
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

export function createPrefixedLazyTrackedIndexChanges(
  prefixChanges: readonly WorkPaperCellChange[],
  lazyTailChanges: WorkPaperCellChange[],
): WorkPaperCellChange[] {
  const prefixLength = prefixChanges.length
  const tailLength = lazyTailChanges.length
  const length = prefixLength + tailLength
  const cache: WorkPaperCellChange[] = []
  cache.length = length
  for (let index = 0; index < prefixLength; index += 1) {
    cache[index] = prefixChanges[index]!
  }
  let fullyMaterialized = false
  const hasMaterializedIndex = (index: number): boolean => Object.prototype.hasOwnProperty.call(cache, index)
  const materialize = (index: number): WorkPaperCellChange => {
    if (hasMaterializedIndex(index)) {
      return cache[index]!
    }
    cache[index] = lazyTailChanges[index - prefixLength]!
    return cache[index]
  }
  const forceMaterialize = (): void => {
    if (fullyMaterialized) {
      return
    }
    forceMaterializeTrackedIndexChanges(lazyTailChanges)
    for (let index = prefixLength; index < length; index += 1) {
      if (!hasMaterializedIndex(index)) {
        cache[index] = lazyTailChanges[index - prefixLength]!
      }
    }
    fullyMaterialized = true
  }
  const detach = (options: TrackedIndexDetachOptions = {}): void => {
    detachTrackedIndexChanges(lazyTailChanges, options)
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

export function tryCreateLazyPhysicalTrackedIndexChanges(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  sheetId: number,
  formatAddressCached: (row: number, col: number) => string,
  options: { readonly sortedSliceSplit?: number } = {},
): WorkPaperCellChange[] | null {
  const workbook = engine.workbook
  const sheet = workbook.getSheetById(sheetId)
  if (sheet && sheet.structureVersion !== 1 && !trackedIndicesUseVisiblePhysicalPositions(workbook, changedCellIndices, sheetId)) {
    return null
  }
  const cellStore = workbook.cellStore
  const length = changedCellIndices.length
  const split = options.sortedSliceSplit
  const sortedSliceSplit = split !== undefined && split > 0 && split < length ? split : undefined
  let isSorted = true
  let isReverseSorted = sortedSliceSplit === undefined
  let previousRow = -1
  let previousCol = -1
  let previousRightRow = -1
  let previousRightCol = -1
  let previousReverseRow = Number.POSITIVE_INFINITY
  let previousReverseCol = Number.POSITIVE_INFINITY
  for (let index = 0; index < length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    if (cellStore.sheetIds[cellIndex] !== sheetId) {
      return null
    }
    const row = cellStore.rows[cellIndex] ?? 0
    const col = cellStore.cols[cellIndex] ?? 0
    if (isReverseSorted && (row > previousReverseRow || (row === previousReverseRow && col >= previousReverseCol))) {
      isReverseSorted = false
    }
    previousReverseRow = row
    previousReverseCol = col
    if (sortedSliceSplit !== undefined && index >= sortedSliceSplit) {
      if (row < previousRightRow || (row === previousRightRow && col < previousRightCol)) {
        return null
      }
      previousRightRow = row
      previousRightCol = col
      continue
    }
    if (row < previousRow || (row === previousRow && col < previousCol)) {
      if (sortedSliceSplit !== undefined) {
        return null
      }
      isSorted = false
      previousRow = row
      previousCol = col
      continue
    }
    previousRow = row
    previousCol = col
  }
  const sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
  if (sortedSliceSplit !== undefined) {
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
  if (!isSorted) {
    if (!isReverseSorted) {
      return null
    }
    return createLazyPhysicalTrackedIndexChanges(
      sheetId,
      sheetName,
      cellStore,
      engine,
      copyReverseSortedTrackedCellIndices(changedCellIndices),
      formatAddressCached,
    )
  }
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

function copyReverseSortedTrackedCellIndices(changedCellIndices: readonly number[] | Uint32Array): Uint32Array {
  const length = changedCellIndices.length
  const reversed = new Uint32Array(length)
  for (let index = 0; index < length; index += 1) {
    reversed[length - index - 1] = changedCellIndices[index]!
  }
  return reversed
}

export function copyOrderedTrackedCellIndices(
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

export function copyTrackedCellIndices(changedCellIndices: readonly number[] | Uint32Array): Uint32Array {
  return Uint32Array.from(changedCellIndices)
}

export function forceMaterializeTrackedIndexChanges(changes: readonly WorkPaperCellChange[]): boolean {
  const deferred = DEFERRED_TRACKED_INDEX_CHANGES.get(changes)
  if (!deferred) {
    return false
  }
  deferred.forceMaterialize()
  DEFERRED_TRACKED_INDEX_CHANGES.delete(changes)
  return true
}

export function detachTrackedIndexChanges(changes: readonly WorkPaperCellChange[], options: TrackedIndexDetachOptions = {}): boolean {
  const deferred = DEFERRED_TRACKED_INDEX_CHANGES.get(changes)
  if (!deferred) {
    return false
  }
  deferred.detach(options)
  return true
}

export function hasDeferredTrackedIndexChanges(changes: readonly WorkPaperCellChange[]): boolean {
  return DEFERRED_TRACKED_INDEX_CHANGES.has(changes)
}
