import { indexToColumn } from '@bilig/formula'
import type { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkPaperCellChange } from './work-paper-types.js'

export interface MaterializedTrackedIndexChanges {
  readonly changes: readonly WorkPaperCellChange[]
  readonly ordered: boolean
}

export function materializeTrackedIndexChangesWithMetadata(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  options: { readonly explicitChangedCount?: number } = {},
): MaterializedTrackedIndexChanges {
  if (changedCellIndices.length === 0) {
    return { changes: [], ordered: true }
  }
  const workbook = engine.workbook
  const cellStore = workbook.cellStore
  const columnLabels: string[] = []
  const formatAddressCached = (row: number, col: number): string => {
    let label = columnLabels[col]
    if (label === undefined) {
      label = indexToColumn(col)
      columnLabels[col] = label
    }
    return `${label}${row + 1}`
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
  options: { readonly explicitChangedCount?: number } = {},
): readonly WorkPaperCellChange[] {
  return materializeTrackedIndexChangesWithMetadata(engine, changedCellIndices, options).changes
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
