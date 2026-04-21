import { indexToColumn } from '@bilig/formula'
import type { SpreadsheetEngine } from '@bilig/core'
import type { CellValue } from '@bilig/protocol'
import type { WorkPaperCellChange } from './work-paper-types.js'

export function materializeTrackedIndexChanges(
  engine: SpreadsheetEngine,
  changedCellIndices: readonly number[] | Uint32Array,
  options: { readonly explicitChangedCount?: number } = {},
): readonly WorkPaperCellChange[] {
  if (changedCellIndices.length === 0) {
    return []
  }
  const workbook = engine.workbook
  const cellStore = workbook.cellStore
  const readValue = (cellIndex: number): CellValue => cellStore.getValue(cellIndex, (id) => engine.strings.get(id))
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
    return []
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
          changes.push(
            readPhysicalTrackedIndexChange(explicitCellIndex, firstSheetId, sheetName, cellStore, readValue, formatAddressCached),
          )
          explicitIndex += 1
        } else {
          changes.push(
            readPhysicalTrackedIndexChange(recalculatedCellIndex, firstSheetId, sheetName, cellStore, readValue, formatAddressCached),
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
            readValue,
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
            readValue,
            formatAddressCached,
          ),
        )
        recalculatedIndex += 1
      }
      return changes
    }
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
      changes.push({
        kind: 'cell',
        address: { sheet: firstSheetId, row, col },
        sheetName,
        a1: formatAddressCached(row, col),
        newValue: readValue(cellIndex),
      })
    }
    return changes
  }

  const sheetNames = new Map<number, string>()
  const physicalSheetIds = new Set<number>()
  for (let index = 0; index < changedCellIndices.length; index += 1) {
    const cellIndex = changedCellIndices[index]!
    const sheetId = cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      continue
    }
    let sheetName = sheetNames.get(sheetId)
    if (sheetName === undefined) {
      const sheet = workbook.getSheetById(sheetId)
      sheetName = sheet?.name ?? workbook.getSheetNameById(sheetId)
      sheetNames.set(sheetId, sheetName)
      if (!sheet || sheet.structureVersion === 1) {
        physicalSheetIds.add(sheetId)
      }
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
    changes.push({
      kind: 'cell',
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: formatAddressCached(row, col),
      newValue: readValue(cellIndex),
    })
  }
  return changes
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
  readValue: (cellIndex: number) => CellValue,
  formatAddressCached: (row: number, col: number) => string,
): WorkPaperCellChange {
  const row = cellStore.rows[cellIndex]!
  const col = cellStore.cols[cellIndex]!
  return {
    kind: 'cell',
    address: { sheet: sheetId, row, col },
    sheetName,
    a1: formatAddressCached(row, col),
    newValue: readValue(cellIndex),
  }
}
