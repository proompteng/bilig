import { formatAddress } from '@bilig/formula'
import type { LiteralInput } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import type { WorkbookStore } from './workbook-store.js'

export type EngineCellMutationAt =
  | { kind: 'setCellValue'; row: number; col: number; value: LiteralInput }
  | { kind: 'setCellFormula'; row: number; col: number; formula: string }
  | { kind: 'clearCell'; row: number; col: number }

export interface EngineCellMutationRef {
  sheetId: number
  mutation: EngineCellMutationAt
  cellIndex?: number
}

export interface EngineFormulaSourceRef {
  sheetId: number
  row: number
  col: number
  source: string
  cellIndex?: number
}

export interface EngineExistingNumericCellMutationRef {
  sheetId: number
  row: number
  col: number
  cellIndex: number
  value: number
  emitTracked?: boolean
  trustedExistingNumericLiteral?: boolean
  oldNumericValue?: number
}

export interface EngineExistingNumericCellMutationResult {
  readonly changedCellIndices?: Uint32Array
  readonly changedCellCount?: number
  readonly firstChangedCellIndex?: number
  readonly secondChangedCellIndex?: number
  readonly secondChangedRow?: number
  readonly secondChangedCol?: number
  readonly secondChangedNumericValue?: number
  readonly explicitChangedCount: number
}

export function cloneCellMutationAt(mutation: EngineCellMutationAt): EngineCellMutationAt {
  switch (mutation.kind) {
    case 'setCellValue':
      return {
        kind: 'setCellValue',
        row: mutation.row,
        col: mutation.col,
        value: mutation.value,
      }
    case 'setCellFormula':
      return {
        kind: 'setCellFormula',
        row: mutation.row,
        col: mutation.col,
        formula: mutation.formula,
      }
    case 'clearCell':
      return {
        kind: 'clearCell',
        row: mutation.row,
        col: mutation.col,
      }
  }
}

export function cloneCellMutationRef(ref: EngineCellMutationRef): EngineCellMutationRef {
  return {
    sheetId: ref.sheetId,
    mutation: cloneCellMutationAt(ref.mutation),
    ...(ref.cellIndex !== undefined ? { cellIndex: ref.cellIndex } : {}),
  }
}

export function cellMutationRefToEngineOp(workbook: Pick<WorkbookStore, 'getSheetById'>, ref: EngineCellMutationRef): EngineOp {
  const sheet = workbook.getSheetById(ref.sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${ref.sheetId}`)
  }
  const address = formatAddress(ref.mutation.row, ref.mutation.col)
  switch (ref.mutation.kind) {
    case 'setCellValue':
      return {
        kind: 'setCellValue',
        sheetName: sheet.name,
        address,
        value: ref.mutation.value,
      }
    case 'setCellFormula':
      return {
        kind: 'setCellFormula',
        sheetName: sheet.name,
        address,
        formula: ref.mutation.formula,
      }
    case 'clearCell':
      return {
        kind: 'clearCell',
        sheetName: sheet.name,
        address,
      }
  }
}

export function countPotentialNewCellsForMutationRefs(refs: readonly EngineCellMutationRef[]): number {
  let count = 0
  for (let index = 0; index < refs.length; index += 1) {
    if (refs[index]?.mutation.kind !== 'clearCell') {
      count += 1
    }
  }
  return count
}
