import type { CellStyleRecord, SheetStyleRangeSnapshot } from '@bilig/protocol'
import { internImportedStyle } from './xlsx-import-cell-styles.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'

type StyleRun = { row: number; startColumn: number; endColumn: number; styleId: string }
type StyledCell = { readonly row: number; readonly column: number; readonly styleId: string }

export function buildLargeSimpleStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  return styleEntriesAreRowMajor(cellScan, stylesByIndex)
    ? buildRowMajorStyleRanges(sheetName, cellScan, stylesByIndex, styleCatalog)
    : buildSortedStyleRanges(sheetName, cellScan, stylesByIndex, styleCatalog)
}

function buildRowMajorStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  const ranges: SheetStyleRangeSnapshot[] = []
  let active: StyleRun | undefined
  cellScan.styleIndexes.forEach((row, column, styleIndex) => {
    const style = stylesByIndex.get(styleIndex)
    if (!style) {
      return
    }
    const styleId = internImportedStyle(style, styleCatalog)
    active = appendStyleRun(ranges, sheetName, active, row, column, styleId)
  })
  if (active) {
    ranges.push(styleRunToRange(sheetName, active))
  }
  return ranges
}

function buildSortedStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  const styledCells: StyledCell[] = []
  cellScan.styleIndexes.forEach((row, column, styleIndex) => {
    const style = stylesByIndex.get(styleIndex)
    if (!style) {
      return
    }
    styledCells.push({
      row,
      column,
      styleId: internImportedStyle(style, styleCatalog),
    })
  })
  styledCells.sort((left, right) => left.row - right.row || left.column - right.column || left.styleId.localeCompare(right.styleId))
  const ranges: SheetStyleRangeSnapshot[] = []
  let active: StyleRun | undefined
  for (const cell of styledCells) {
    active = appendStyleRun(ranges, sheetName, active, cell.row, cell.column, cell.styleId)
  }
  if (active) {
    ranges.push(styleRunToRange(sheetName, active))
  }
  return ranges
}

function styleEntriesAreRowMajor(
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
): boolean {
  let ordered = true
  let lastRow = -1
  let lastColumn = -1
  cellScan.styleIndexes.forEach((row, column, styleIndex) => {
    if (!ordered || !stylesByIndex.has(styleIndex)) {
      return
    }
    if (row < lastRow || (row === lastRow && column <= lastColumn)) {
      ordered = false
      return
    }
    lastRow = row
    lastColumn = column
  })
  return ordered
}

function appendStyleRun(
  ranges: SheetStyleRangeSnapshot[],
  sheetName: string,
  active: StyleRun | undefined,
  row: number,
  column: number,
  styleId: string,
): StyleRun {
  if (active && active.row === row && active.endColumn + 1 === column && active.styleId === styleId) {
    active.endColumn = column
    return active
  }
  if (active) {
    ranges.push(styleRunToRange(sheetName, active))
  }
  return { row, startColumn: column, endColumn: column, styleId }
}

function styleRunToRange(
  sheetName: string,
  run: { readonly row: number; readonly startColumn: number; readonly endColumn: number; readonly styleId: string },
): SheetStyleRangeSnapshot {
  return {
    range: {
      sheetName,
      startAddress: encodeCellAddress(run.row, run.startColumn),
      endAddress: encodeCellAddress(run.row, run.endColumn),
    },
    styleId: run.styleId,
  }
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}
