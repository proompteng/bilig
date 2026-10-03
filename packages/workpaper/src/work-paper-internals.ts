import { ValueTag, type CellValue } from '@bilig/protocol'
import type {
  RawCellContent,
  WorkPaperAddressLike,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperDependencyRef,
  WorkPaperInternals,
  WorkPaperSheetDimensions,
} from './work-paper-types.js'

export interface WorkPaperInternalsHooks {
  readonly getCellDependents: (reference: WorkPaperAddressLike) => WorkPaperDependencyRef[]
  readonly getCellFormula: (address: WorkPaperCellAddress) => string | undefined
  readonly getCellPrecedents: (reference: WorkPaperAddressLike) => WorkPaperDependencyRef[]
  readonly getCellValue: (address: WorkPaperCellAddress) => CellValue
  readonly getNamedExpressionsFromFormula: (formula: string) => string[]
  readonly getRangeSerialized: (range: WorkPaperCellRange) => RawCellContent[][]
  readonly getRangeValues: (range: WorkPaperCellRange) => CellValue[][]
  readonly getSheetDimensions: (sheetId: number) => WorkPaperSheetDimensions
  readonly getSheetId: (name: string) => number | undefined
  readonly getSheetName: (sheetId: number) => string | undefined
  readonly getSheetNames: () => string[]
  readonly hasCellValueOrFormula: (address: WorkPaperCellAddress) => boolean
  readonly isCellPartOfArray: (address: WorkPaperCellAddress) => boolean
  readonly normalizeFormula: (formula: string) => string
  readonly recalculate: () => WorkPaperChange[]
  readonly calculateFormula: (formula: string, scope?: number) => CellValue | CellValue[][]
  readonly countSheets: () => number
  readonly validateFormula: (formula: string) => boolean
}

export function createWorkPaperInternals(hooks: WorkPaperInternalsHooks): WorkPaperInternals {
  return Object.freeze({
    graph: Object.freeze({
      getDependents: hooks.getCellDependents,
      getPrecedents: hooks.getCellPrecedents,
    }),
    rangeMapping: Object.freeze({
      getValues: hooks.getRangeValues,
      getSerialized: hooks.getRangeSerialized,
    }),
    arrayMapping: Object.freeze({
      isPartOfArray: hooks.isCellPartOfArray,
      getFormula: hooks.getCellFormula,
    }),
    sheetMapping: Object.freeze({
      getSheetName: hooks.getSheetName,
      getSheetId: hooks.getSheetId,
      getSheetNames: hooks.getSheetNames,
      countSheets: hooks.countSheets,
    }),
    addressMapping: Object.freeze({
      has: hooks.hasCellValueOrFormula,
      getValue: hooks.getCellValue,
      getFormula: hooks.getCellFormula,
    }),
    dependencyGraph: Object.freeze({
      getCellDependents: hooks.getCellDependents,
      getCellPrecedents: hooks.getCellPrecedents,
    }),
    evaluator: Object.freeze({
      recalculate: hooks.recalculate,
      calculateFormula: hooks.calculateFormula,
    }),
    columnSearch: Object.freeze({
      find: (sheetId: number, column: number, matcher: string | ((value: CellValue) => boolean)): WorkPaperCellAddress[] => {
        const dimensions = hooks.getSheetDimensions(sheetId)
        const matches: WorkPaperCellAddress[] = []
        for (let row = 0; row < dimensions.height; row += 1) {
          const address = { sheet: sheetId, row, col: column }
          const value = hooks.getCellValue(address)
          if (doesWorkPaperColumnSearchValueMatch(value, matcher)) {
            matches.push(address)
          }
        }
        return matches
      },
    }),
    lazilyTransformingAstService: Object.freeze({
      normalizeFormula: hooks.normalizeFormula,
      validateFormula: hooks.validateFormula,
      getNamedExpressionsFromFormula: hooks.getNamedExpressionsFromFormula,
    }),
  })
}

export function doesWorkPaperColumnSearchValueMatch(value: CellValue, matcher: string | ((value: CellValue) => boolean)): boolean {
  return typeof matcher === 'string' ? value.tag === ValueTag.String && value.value === matcher : matcher(value)
}
