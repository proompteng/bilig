import { attachRuntimeSnapshot } from '@bilig/core/headless-runtime'
import type { CellValue, WorkbookSnapshot } from '@bilig/protocol'
import { assertRange } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperCellRange, WorkPaperSheetDimensions } from './work-paper-types.js'

export interface WorkPaperSheetListRecord {
  readonly id: number
  readonly name: string
}

export function buildWorkPaperDenseRange<Value>(range: WorkPaperCellRange, read: (address: WorkPaperCellAddress) => Value): Value[][] {
  assertRange(range)
  const height = range.end.row - range.start.row + 1
  const width = range.end.col - range.start.col + 1
  return Array.from({ length: height }, (_row, rowOffset) =>
    Array.from({ length: width }, (_column, colOffset) =>
      read({
        sheet: range.start.sheet,
        row: range.start.row + rowOffset,
        col: range.start.col + colOffset,
      }),
    ),
  )
}

export function workPaperRangeForSheetDimensions(sheetId: number, dimensions: WorkPaperSheetDimensions): WorkPaperCellRange | undefined {
  if (dimensions.width === 0 || dimensions.height === 0) {
    return undefined
  }
  return {
    start: { sheet: sheetId, row: 0, col: 0 },
    end: { sheet: sheetId, row: dimensions.height - 1, col: dimensions.width - 1 },
  }
}

export function readWorkPaperSheetRange<Value>(args: {
  readonly sheetId: number
  readonly dimensions: WorkPaperSheetDimensions
  readonly readRange: (range: WorkPaperCellRange) => Value[][]
}): Value[][] {
  const range = workPaperRangeForSheetDimensions(args.sheetId, args.dimensions)
  return range === undefined ? [] : args.readRange(range)
}

export function collectWorkPaperSheetsByName<Value>(
  sheets: readonly WorkPaperSheetListRecord[],
  readSheet: (sheet: WorkPaperSheetListRecord) => Value,
): Record<string, Value> {
  return Object.fromEntries(sheets.map((sheet) => [sheet.name, readSheet(sheet)]))
}

export function collectSerializedWorkPaperSheets(args: {
  readonly sheets: readonly WorkPaperSheetListRecord[]
  readonly readSheet: (sheet: WorkPaperSheetListRecord) => RawCellContent[][]
  readonly runtimeSnapshot: WorkbookSnapshot
}): Record<string, RawCellContent[][]> {
  return attachRuntimeSnapshot(collectWorkPaperSheetsByName(args.sheets, args.readSheet), args.runtimeSnapshot)
}

export type WorkPaperSheetValues = CellValue[][]
export type WorkPaperSheetFormulas = Array<Array<string | undefined>>
