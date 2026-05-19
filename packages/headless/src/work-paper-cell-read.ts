import type { SheetRecord, SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { parseFormula } from '@bilig/formula'
import { readFastRangeValues } from './fast-range-read.js'
import { readTrackedRuntimeCellValue } from './work-paper-tracked-event-helpers.js'
import { stripLeadingEquals } from './work-paper-runtime-helpers.js'
import {
  classifyWorkPaperCell,
  doesWorkPaperCellHaveSimpleValue,
  isWorkPaperCellPartOfArray,
  workPaperCellValueDetailedType,
  workPaperCellValueType,
} from './work-paper-sheet-inspection.js'
import type {
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
} from './work-paper-types.js'

export function getVisibleWorkPaperCellIndexInSheet(sheet: SheetRecord, row: number, col: number): number | undefined {
  if (sheet.structureVersion === 1) {
    const cellIndex = sheet.grid.getPhysical(row, col)
    if (cellIndex !== -1 && sheet.logical.cellIdentityMatchesVisiblePosition(cellIndex, row, col)) {
      return cellIndex
    }
  }
  return sheet.logical.getVisibleCell(row, col)
}

export function readWorkPaperCellValue(args: {
  readonly engine: SpreadsheetEngine
  readonly sheet: SheetRecord
  readonly address: WorkPaperCellAddress
}): CellValue {
  const cellIndex = getVisibleWorkPaperCellIndexInSheet(args.sheet, args.address.row, args.address.col)
  return cellIndex === undefined
    ? { tag: ValueTag.Empty }
    : readTrackedRuntimeCellValue(args.engine.workbook.cellStore, cellIndex, args.engine.strings)
}

export function readWorkPaperCellFormula(args: {
  readonly engine: SpreadsheetEngine
  readonly sheetName: string
  readonly a1: string
  readonly ownerSheetId: number
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
}): string | undefined {
  const cell = args.engine.getCell(args.sheetName, args.a1)
  return cell.formula ? `=${args.restorePublicFormula(cell.formula, args.ownerSheetId)}` : undefined
}

export function readWorkPaperCellHyperlink(formula: string | undefined): string | undefined {
  if (!formula) {
    return undefined
  }
  const parsed = parseFormula(stripLeadingEquals(formula))
  if (parsed.kind !== 'CallExpr' || parsed.callee.trim().toUpperCase() !== 'HYPERLINK') {
    return undefined
  }
  const firstArgument = parsed.args[0]
  return firstArgument?.kind === 'StringLiteral' ? firstArgument.value : undefined
}

export function readWorkPaperRangeValues(args: {
  readonly engine: SpreadsheetEngine
  readonly range: WorkPaperCellRange
  readonly rangeRef: () => Parameters<SpreadsheetEngine['getRangeValues']>[0]
}): CellValue[][] {
  const fastValues = readFastRangeValues(args.engine, args.range)
  return fastValues ?? args.engine.getRangeValues(args.rangeRef())
}

export function readWorkPaperCellClassification(args: {
  readonly hasFormula: boolean
  readonly isEmpty: boolean
  readonly isPartOfArray: boolean
}): WorkPaperCellType {
  return classifyWorkPaperCell(args)
}

export function readWorkPaperCellHasSimpleValue(args: { readonly hasFormula: boolean; readonly isEmpty: boolean }): boolean {
  return doesWorkPaperCellHaveSimpleValue(args)
}

export function readWorkPaperCellPartOfArray(args: {
  readonly address: WorkPaperCellAddress
  readonly spillRanges: ReturnType<SpreadsheetEngine['getSpillRanges']>
  readonly requireSheetId: (sheetName: string) => number
}): boolean {
  return isWorkPaperCellPartOfArray(args)
}

export function readWorkPaperCellValueType(value: CellValue): WorkPaperCellValueType {
  return workPaperCellValueType(value)
}

export function readWorkPaperCellValueDetailedType(args: {
  readonly value: CellValue
  readonly format: string | undefined
}): WorkPaperCellValueDetailedType {
  return workPaperCellValueDetailedType(args)
}
