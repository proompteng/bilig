import type { CellStyleRecord, SheetStyleRangeSnapshot } from '@bilig/protocol'
import { internImportedStyle } from './xlsx-import-cell-styles.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'

export function buildLargeSimpleStyleRanges(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  stylesByIndex: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>,
  styleCatalog: Map<string, CellStyleRecord>,
): SheetStyleRangeSnapshot[] {
  const styleIdsByIndex = new Map<number, string>()
  const builder = new StyleRangeBuilder(sheetName)
  const append = (row: number, column: number, styleIndex: number): void => {
    let styleId = styleIdsByIndex.get(styleIndex)
    if (!styleId) {
      const style = stylesByIndex.get(styleIndex)
      if (!style) {
        return
      }
      styleId = internImportedStyle(style, styleCatalog)
      styleIdsByIndex.set(styleIndex, styleId)
    }
    builder.add(row, column, styleId)
  }
  if (cellScan.styleIndexes.isRowMajorOrdered) {
    cellScan.styleIndexes.forEach(append)
  } else {
    const entries: { row: number; column: number; styleIndex: number }[] = []
    cellScan.styleIndexes.forEach((row, column, styleIndex) => {
      entries.push({ row, column, styleIndex })
    })
    entries
      .toSorted((left, right) => left.row - right.row || left.column - right.column || left.styleIndex - right.styleIndex)
      .forEach((entry) => append(entry.row, entry.column, entry.styleIndex))
  }
  return builder.finish()
}

export function buildLargeSimpleStyleRangesForCells(
  sheetName: string,
  cells: readonly { readonly row: number; readonly column: number; readonly styleId: string }[],
): SheetStyleRangeSnapshot[] {
  const builder = new StyleRangeBuilder(sheetName)
  for (const cell of cells) {
    builder.add(cell.row, cell.column, cell.styleId)
  }
  return builder.finish()
}

function styleRunToRange(
  sheetName: string,
  run: {
    readonly startRow: number
    readonly endRow: number
    readonly startColumn: number
    readonly endColumn: number
    readonly styleId: string
  },
): SheetStyleRangeSnapshot {
  return {
    range: {
      sheetName,
      startAddress: encodeCellAddress(run.startRow, run.startColumn),
      endAddress: encodeCellAddress(run.endRow, run.endColumn),
    },
    styleId: run.styleId,
  }
}

class StyleRangeBuilder {
  private readonly ranges: SheetStyleRangeSnapshot[] = []
  private readonly openRects = new Map<
    string,
    { startRow: number; endRow: number; startColumn: number; endColumn: number; styleId: string }
  >()
  private activeRun: { row: number; startColumn: number; endColumn: number; styleId: string } | undefined
  private currentRow = -1

  constructor(private readonly sheetName: string) {}

  add(row: number, column: number, styleId: string): void {
    if (this.activeRun && this.activeRun.row === row && this.activeRun.endColumn + 1 === column && this.activeRun.styleId === styleId) {
      this.activeRun.endColumn = column
      return
    }
    this.flushActiveRun()
    this.activeRun = { row, startColumn: column, endColumn: column, styleId }
  }

  finish(): SheetStyleRangeSnapshot[] {
    this.flushActiveRun()
    for (const key of this.openRects.keys()) {
      this.flushOpenRect(key)
    }
    return this.ranges
  }

  private flushActiveRun(): void {
    if (!this.activeRun) {
      return
    }
    this.appendRowRun(this.activeRun)
    this.activeRun = undefined
  }

  private appendRowRun(run: {
    readonly row: number
    readonly startColumn: number
    readonly endColumn: number
    readonly styleId: string
  }): void {
    if (this.currentRow !== run.row) {
      if (this.currentRow >= 0) {
        this.flushOpenRectsNotEndingAt(this.currentRow)
      }
      this.currentRow = run.row
    }
    const key = `${String(run.startColumn)}\t${String(run.endColumn)}\t${run.styleId}`
    const rect = this.openRects.get(key)
    if (rect && rect.endRow === run.row - 1) {
      rect.endRow = run.row
      return
    }
    if (rect) {
      this.flushOpenRect(key)
    }
    this.openRects.set(key, {
      startRow: run.row,
      endRow: run.row,
      startColumn: run.startColumn,
      endColumn: run.endColumn,
      styleId: run.styleId,
    })
  }

  private flushOpenRectsNotEndingAt(row: number): void {
    for (const [key, rect] of this.openRects) {
      if (rect.endRow !== row) {
        this.flushOpenRect(key)
      }
    }
  }

  private flushOpenRect(key: string): void {
    const rect = this.openRects.get(key)
    if (!rect) {
      return
    }
    this.ranges.push(styleRunToRange(this.sheetName, rect))
    this.openRects.delete(key)
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
