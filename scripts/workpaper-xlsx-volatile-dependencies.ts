import * as XLSX from 'xlsx'

import { parseFormula, type FormulaNode } from '@bilig/formula'
import type { FormulaCellRecord, WorkPaperXlsxFormulaSkipReason } from './check-workpaper-xlsx-corpus-types.ts'

interface FormulaDependencyRange {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export function markVolatileDependentFormulaCells(
  formulaCells: FormulaCellRecord[],
  skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
): FormulaCellRecord[] {
  const volatileRangesBySheet = new Map<string, FormulaDependencyRange[]>()
  const dependenciesByFormula = new Map<string, readonly FormulaDependencyRange[]>()
  let records = formulaCells
  let changed = false

  for (const record of records) {
    if (record.skipReason === 'volatile-or-environment-dependent-formula') {
      addVolatileRange(volatileRangesBySheet, recordCellRange(record))
    }
  }

  do {
    changed = false
    records = records.map((record) => {
      if (record.skipReason) {
        return record
      }
      const key = formulaRecordKey(record)
      let dependencies = dependenciesByFormula.get(key)
      if (!dependencies) {
        dependencies = collectFormulaDependencyRanges(record.formula, record.sheetName)
        dependenciesByFormula.set(key, dependencies)
      }
      if (!dependencies.some((dependency) => overlapsAnyVolatileRange(dependency, volatileRangesBySheet))) {
        return record
      }

      changed = true
      skippedByReason['volatile-or-environment-dependent-formula'] += 1
      const skippedRecord: FormulaCellRecord = {
        ...record,
        skipReason: 'volatile-or-environment-dependent-formula',
      }
      addVolatileRange(volatileRangesBySheet, recordCellRange(record))
      return skippedRecord
    })
  } while (changed)

  return records
}

function formulaRecordKey(record: FormulaCellRecord): string {
  return `${record.sheetName}!${record.address}`
}

function recordCellRange(record: FormulaCellRecord): FormulaDependencyRange {
  return {
    sheetName: record.sheetName,
    rowStart: record.row,
    rowEnd: record.row,
    colStart: record.col,
    colEnd: record.col,
  }
}

function addVolatileRange(rangesBySheet: Map<string, FormulaDependencyRange[]>, range: FormulaDependencyRange): void {
  const ranges = rangesBySheet.get(range.sheetName)
  if (ranges) {
    ranges.push(range)
    return
  }
  rangesBySheet.set(range.sheetName, [range])
}

function overlapsAnyVolatileRange(
  dependency: FormulaDependencyRange,
  volatileRangesBySheet: ReadonlyMap<string, readonly FormulaDependencyRange[]>,
): boolean {
  return (volatileRangesBySheet.get(dependency.sheetName) ?? []).some((volatile) => rangesOverlap(dependency, volatile))
}

function rangesOverlap(left: FormulaDependencyRange, right: FormulaDependencyRange): boolean {
  return left.rowStart <= right.rowEnd && right.rowStart <= left.rowEnd && left.colStart <= right.colEnd && right.colStart <= left.colEnd
}

function collectFormulaDependencyRanges(formula: string, ownerSheetName: string): FormulaDependencyRange[] {
  try {
    const ranges: FormulaDependencyRange[] = []
    collectFormulaNodeDependencyRanges(parseFormula(formula), ownerSheetName, ranges)
    return ranges
  } catch {
    return []
  }
}

function collectFormulaNodeDependencyRanges(node: FormulaNode, ownerSheetName: string, ranges: FormulaDependencyRange[]): void {
  switch (node.kind) {
    case 'CellRef': {
      const cell = decodeFormulaCellAddress(node.ref)
      if (cell) {
        ranges.push({
          sheetName: node.sheetName ?? ownerSheetName,
          rowStart: cell.r,
          rowEnd: cell.r,
          colStart: cell.c,
          colEnd: cell.c,
        })
      }
      return
    }
    case 'RangeRef': {
      const range = dependencyRangeFromRangeRef(node, ownerSheetName)
      if (range) {
        ranges.push(range)
      }
      return
    }
    case 'RowRef': {
      const row = Number.parseInt(node.ref.replaceAll('$', ''), 10)
      if (Number.isInteger(row) && row > 0) {
        ranges.push({
          sheetName: node.sheetName ?? ownerSheetName,
          rowStart: row - 1,
          rowEnd: row - 1,
          colStart: 0,
          colEnd: Number.MAX_SAFE_INTEGER,
        })
      }
      return
    }
    case 'ColumnRef': {
      const col = decodeFormulaColumn(node.ref)
      if (col !== undefined) {
        ranges.push({
          sheetName: node.sheetName ?? ownerSheetName,
          rowStart: 0,
          rowEnd: Number.MAX_SAFE_INTEGER,
          colStart: col,
          colEnd: col,
        })
      }
      return
    }
    case 'ArrayConstant':
      for (const row of node.rows) {
        for (const entry of row) {
          collectFormulaNodeDependencyRanges(entry, ownerSheetName, ranges)
        }
      }
      return
    case 'UnaryExpr':
      collectFormulaNodeDependencyRanges(node.argument, ownerSheetName, ranges)
      return
    case 'BinaryExpr':
      collectFormulaNodeDependencyRanges(node.left, ownerSheetName, ranges)
      collectFormulaNodeDependencyRanges(node.right, ownerSheetName, ranges)
      return
    case 'CallExpr':
      for (const arg of node.args) {
        collectFormulaNodeDependencyRanges(arg, ownerSheetName, ranges)
      }
      return
    case 'InvokeExpr':
      collectFormulaNodeDependencyRanges(node.callee, ownerSheetName, ranges)
      for (const arg of node.args) {
        collectFormulaNodeDependencyRanges(arg, ownerSheetName, ranges)
      }
      return
    case 'BooleanLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return
  }
}

function dependencyRangeFromRangeRef(
  node: Extract<FormulaNode, { kind: 'RangeRef' }>,
  ownerSheetName: string,
): FormulaDependencyRange | undefined {
  if (node.sheetEndName !== undefined) {
    return undefined
  }
  const sheetName = node.sheetName ?? ownerSheetName
  if (node.refKind === 'cells') {
    const start = decodeFormulaCellAddress(node.start)
    const end = decodeFormulaCellAddress(node.end)
    return start && end
      ? {
          sheetName,
          rowStart: Math.min(start.r, end.r),
          rowEnd: Math.max(start.r, end.r),
          colStart: Math.min(start.c, end.c),
          colEnd: Math.max(start.c, end.c),
        }
      : undefined
  }
  if (node.refKind === 'rows') {
    const startRow = Number.parseInt(node.start.replaceAll('$', ''), 10)
    const endRow = Number.parseInt(node.end.replaceAll('$', ''), 10)
    return Number.isInteger(startRow) && Number.isInteger(endRow) && startRow > 0 && endRow > 0
      ? {
          sheetName,
          rowStart: Math.min(startRow, endRow) - 1,
          rowEnd: Math.max(startRow, endRow) - 1,
          colStart: 0,
          colEnd: Number.MAX_SAFE_INTEGER,
        }
      : undefined
  }
  const startCol = decodeFormulaColumn(node.start)
  const endCol = decodeFormulaColumn(node.end)
  return startCol !== undefined && endCol !== undefined
    ? {
        sheetName,
        rowStart: 0,
        rowEnd: Number.MAX_SAFE_INTEGER,
        colStart: Math.min(startCol, endCol),
        colEnd: Math.max(startCol, endCol),
      }
    : undefined
}

function decodeFormulaCellAddress(address: string): { r: number; c: number } | undefined {
  try {
    return XLSX.utils.decode_cell(address.replaceAll('$', ''))
  } catch {
    return undefined
  }
}

function decodeFormulaColumn(column: string): number | undefined {
  try {
    return XLSX.utils.decode_col(column.replaceAll('$', ''))
  } catch {
    return undefined
  }
}
