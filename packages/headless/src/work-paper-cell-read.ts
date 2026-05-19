import type { SheetRecord, SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { parseFormula } from '@bilig/formula'
import { readFastRangeValueBlock, readFastRangeValues } from './fast-range-read.js'
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
  WorkPaperRangeValueBlock,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
} from './work-paper-types.js'

export function getVisibleWorkPaperCellIndexInSheet(sheet: SheetRecord, row: number, col: number): number | undefined {
  if (sheet.structureVersion === 1) {
    const cellIndex = sheet.grid.getPhysical(row, col)
    return cellIndex === -1 ? undefined : cellIndex
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

export function readWorkPaperCellFormulaByIndex(args: {
  readonly engine: SpreadsheetEngine
  readonly cellIndex: number | undefined
  readonly ownerSheetId: number
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
}): string | undefined {
  if (args.cellIndex === undefined || (args.engine.workbook.cellStore.formulaIds[args.cellIndex] ?? 0) === 0) {
    return undefined
  }
  const formula = args.engine.getCellByIndex(args.cellIndex).formula
  return formula ? `=${args.restorePublicFormula(formula, args.ownerSheetId)}` : undefined
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

export function readWorkPaperRangeValueBlock(args: {
  readonly engine: SpreadsheetEngine
  readonly range: WorkPaperCellRange
  readonly rangeRef: () => Parameters<SpreadsheetEngine['getRangeValues']>[0]
}): WorkPaperRangeValueBlock {
  const fastBlock = readFastRangeValueBlock(args.engine, args.range)
  if (fastBlock) {
    return fastBlock
  }
  const values = args.engine.getRangeValues(args.rangeRef())
  return cellValueMatrixToValueBlock(values)
}

function cellValueMatrixToValueBlock(values: CellValue[][]): WorkPaperRangeValueBlock {
  const rowCount = values.length
  const colCount = values[0]?.length ?? 0
  const area = rowCount * colCount
  const block: WorkPaperRangeValueBlock & { strings?: Map<number, string> } = {
    rowCount,
    colCount,
    tags: new Uint8Array(area),
    numbers: new Float64Array(area),
    stringIds: new Uint32Array(area),
    errors: new Uint16Array(area),
  }
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const row = values[rowOffset]!
    const outputRowOffset = rowOffset * colCount
    for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
      const index = outputRowOffset + colOffset
      const value = row[colOffset] ?? { tag: ValueTag.Empty }
      block.tags[index] = value.tag
      switch (value.tag) {
        case ValueTag.Number:
          block.numbers[index] = value.value ?? 0
          break
        case ValueTag.Boolean:
          block.numbers[index] = value.value ? 1 : 0
          break
        case ValueTag.String:
          if (value.stringId !== undefined) {
            block.stringIds[index] = value.stringId
            block.strings ??= new Map()
            block.strings.set(value.stringId, value.value ?? '')
          }
          break
        case ValueTag.Error:
          block.errors[index] = value.code ?? ErrorCode.None
          break
        case ValueTag.Empty:
          break
      }
    }
  }
  return block
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
