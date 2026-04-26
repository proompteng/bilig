import { evaluateAst, formatAddress, parseFormula, parseRangeAddress, type EvaluationContext } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS, ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { ParsedEditorInput } from './worker-workbook-app-model.js'

const MAX_OPTIMISTIC_FORMULA_RANGE_CELLS = 10_000

export function optimisticCellKey(sheetName: string, address: string): string {
  return `${sheetName}:${address}`
}

export function createOptimisticCellSnapshot(input: {
  readonly sheetName: string
  readonly address: string
  readonly current: CellSnapshot
  readonly parsed: ParsedEditorInput
  readonly evaluateFormula?: (formula: string) => CellSnapshot['value'] | null
}): CellSnapshot {
  const version = nextOptimisticVersion(input.current.version)
  const { formula: _formula, input: _input, ...base } = input.current
  switch (input.parsed.kind) {
    case 'clear':
      return {
        ...base,
        sheetName: input.sheetName,
        address: input.address,
        value: { tag: ValueTag.Empty },
        version,
      }
    case 'formula':
      return {
        ...base,
        sheetName: input.sheetName,
        address: input.address,
        formula: input.parsed.formula,
        value: input.evaluateFormula?.(input.parsed.formula) ?? {
          tag: ValueTag.String,
          value: `=${input.parsed.formula}`,
          stringId: 0,
        },
        version,
      }
    case 'value':
      return {
        ...base,
        sheetName: input.sheetName,
        address: input.address,
        input: input.parsed.value,
        value: valueFromLiteral(input.parsed.value),
        version,
      }
  }
}

export function evaluateOptimisticFormula(input: {
  readonly sheetName: string
  readonly address: string
  readonly formula: string
  readonly getCell: (sheetName: string, address: string) => CellSnapshot
  readonly listSheetNames?: () => string[]
}): CellSnapshot['value'] | null {
  try {
    const context: EvaluationContext = {
      sheetName: input.sheetName,
      currentAddress: input.address,
      resolveCell: (sheetName, address) => input.getCell(sheetName, address).value,
      resolveRange: (sheetName, start, end, refKind) => resolveOptimisticRange(input.getCell, sheetName, start, end, refKind),
      resolveFormula: (sheetName, address) => input.getCell(sheetName, address).formula,
      ...(input.listSheetNames ? { listSheetNames: input.listSheetNames } : {}),
    }
    return evaluateAst(parseFormula(input.formula), context)
  } catch {
    return null
  }
}

export function createSupersedingCellSnapshot(snapshot: CellSnapshot, version: number): CellSnapshot {
  return {
    ...snapshot,
    version: Math.max(nextOptimisticVersion(snapshot.version), version),
  }
}

function valueFromLiteral(value: string | number | boolean | null): CellSnapshot['value'] {
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  if (typeof value === 'string') {
    return { tag: ValueTag.String, value, stringId: 0 }
  }
  return { tag: ValueTag.Empty }
}

function resolveOptimisticRange(
  getCell: (sheetName: string, address: string) => CellSnapshot,
  sheetName: string,
  start: string,
  end: string,
  refKind: 'cells' | 'rows' | 'cols',
): CellSnapshot['value'][] {
  const range = parseRangeAddress(`${start}:${end}`, sheetName)
  const resolvedSheetName = range.sheetName ?? sheetName

  switch (range.kind) {
    case 'cells': {
      assertOptimisticRangeSize((range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1))
      const values: CellSnapshot['value'][] = []
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let col = range.start.col; col <= range.end.col; col += 1) {
          values.push(getCell(resolvedSheetName, formatAddress(row, col)).value)
        }
      }
      return values
    }
    case 'rows': {
      if (refKind !== 'rows') {
        return []
      }
      assertOptimisticRangeSize((range.end.row - range.start.row + 1) * MAX_COLS)
      const values: CellSnapshot['value'][] = []
      for (let row = range.start.row; row <= range.end.row; row += 1) {
        for (let col = 0; col < MAX_COLS; col += 1) {
          values.push(getCell(resolvedSheetName, formatAddress(row, col)).value)
        }
      }
      return values
    }
    case 'cols': {
      if (refKind !== 'cols') {
        return []
      }
      assertOptimisticRangeSize((range.end.col - range.start.col + 1) * MAX_ROWS)
      const values: CellSnapshot['value'][] = []
      for (let row = 0; row < MAX_ROWS; row += 1) {
        for (let col = range.start.col; col <= range.end.col; col += 1) {
          values.push(getCell(resolvedSheetName, formatAddress(row, col)).value)
        }
      }
      return values
    }
  }
}

function assertOptimisticRangeSize(cellCount: number): void {
  if (cellCount > MAX_OPTIMISTIC_FORMULA_RANGE_CELLS) {
    throw new Error(`Optimistic formula range is too large: ${cellCount}`)
  }
}

function nextOptimisticVersion(version: number): number {
  return Number.isInteger(version) && version >= 0 ? version + 1 : 1
}
