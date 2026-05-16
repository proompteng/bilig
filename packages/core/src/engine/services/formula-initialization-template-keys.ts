import type { CompiledFormula, FormulaNode, ParsedDependencyReference, ParsedRangeReferenceInfo } from '@bilig/formula'
import { parseA1RowNumber } from '../../formula/a1-row-number.js'
import { buildFormulaFamilyShapeKey } from '../../formula/formula-family-deps.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import type { RuntimeFormula } from '../runtime-state.js'

export interface InitialTemplateFormulaCacheEntry {
  readonly resolution: FormulaTemplateResolution
  readonly anchorRow: number
  readonly anchorCol: number
  readonly anchorCompiled: CompiledFormula
}

export interface InitialPrefixSumTemplateKey {
  readonly key: number
  readonly rangeCol: number
  readonly rangeColumn: string
}

const INITIAL_PREFIX_SUM_RE = /^SUM\(([A-Z]+)1:\1([1-9]\d*)\)$/

function initialColumnToIndex(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    const code = column.charCodeAt(index)
    value = value * 26 + (code >= 97 && code <= 122 ? code - 96 : code - 64)
  }
  return value - 1
}

function initialReadColumn(source: string, start: number): { readonly column: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (!((code >= 65 && code <= 90) || (code >= 97 && code <= 122))) {
      break
    }
    next += 1
  }
  return next === start ? undefined : { column: source.slice(start, next), next }
}

function initialReadRowNumber(source: string, start: number): { readonly row: number; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    next += 1
  }
  const row = parseA1RowNumber(source.slice(start, next))
  return row === undefined ? undefined : { row, next }
}

function initialReadNumberLiteral(source: string, start: number): { readonly text: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    next += 1
  }
  if (next < source.length && source.charCodeAt(next) === 46) {
    const fractionStart = next + 1
    next = fractionStart
    while (next < source.length) {
      const code = source.charCodeAt(next)
      if (code < 48 || code > 57) {
        break
      }
      next += 1
    }
    if (next === fractionStart) {
      return undefined
    }
  }
  return next === start ? undefined : { text: source.slice(start, next), next }
}

function initialReadRelativeCellToken(
  source: string,
  start: number,
  ownerRow: number,
  ownerCol: number,
): { readonly token: string; readonly next: number } | undefined {
  const column = initialReadColumn(source, start)
  if (!column) {
    return undefined
  }
  const row = initialReadRowNumber(source, column.next)
  if (!row || row.row - 1 !== ownerRow) {
    return undefined
  }
  const col = initialColumnToIndex(column.column)
  return col < 0 ? undefined : { token: `c${col - ownerCol}`, next: row.next }
}

export function tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  let index = source.charCodeAt(0) === 61 ? 1 : 0
  const left = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (!left) {
    return undefined
  }
  index = left.next
  const operator = source[index]
  if (operator !== '+' && operator !== '-' && operator !== '*' && operator !== '/') {
    return undefined
  }
  index += 1
  const rightCell = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (rightCell) {
    return rightCell.next === source.length ? `${left.token}${operator}${rightCell.token}` : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  return rightNumber && rightNumber.next === source.length ? `${left.token}${operator}n${rightNumber.text}` : undefined
}

export function tryBuildInitialPrefixSumTemplateKey(
  source: string,
  ownerRow: number,
  ownerCol: number,
): InitialPrefixSumTemplateKey | undefined {
  const match = INITIAL_PREFIX_SUM_RE.exec(source)
  if (!match) {
    return undefined
  }
  const endRow = parseA1RowNumber(match[2]!)
  if (endRow === undefined || endRow !== ownerRow + 1) {
    return undefined
  }
  const rangeColumn = match[1]!
  const rangeCol = initialColumnToIndex(rangeColumn)
  if (rangeCol < 0) {
    return undefined
  }
  return {
    key: rangeCol - ownerCol,
    rangeCol,
    rangeColumn,
  }
}

export function translateInitialPrefixSumFormula(
  entry: InitialTemplateFormulaCacheEntry,
  source: string,
  ownerRow: number,
  ownerCol: number,
  templateKey: InitialPrefixSumTemplateKey,
): FormulaTemplateResolution {
  const column = templateKey.rangeColumn
  const startAddress = `${column}1`
  const endAddress = `${column}${ownerRow + 1}`
  const rangeAddress = `${startAddress}:${endAddress}`
  const range: ParsedRangeReferenceInfo = {
    address: rangeAddress,
    kind: 'range',
    refKind: 'cells',
    startAddress,
    endAddress,
    startRow: 0,
    endRow: ownerRow,
    startCol: templateKey.rangeCol,
    endCol: templateKey.rangeCol,
  }
  const rangeNode: FormulaNode = {
    kind: 'RangeRef',
    refKind: 'cells',
    start: startAddress,
    end: endAddress,
  }
  const ast: FormulaNode = {
    kind: 'CallExpr',
    callee: entry.anchorCompiled.directAggregateCandidate!.callee,
    args: [rangeNode],
  }
  const rowDelta = ownerRow - entry.anchorRow
  const colDelta = ownerCol - entry.anchorCol
  return {
    ...entry.resolution,
    compiled: {
      ...entry.anchorCompiled,
      source,
      ast,
      optimizedAst: ast,
      astMatchesSource: true,
      deps: [rangeAddress],
      parsedDeps: [range satisfies ParsedDependencyReference],
      symbolicRanges: [rangeAddress],
      parsedSymbolicRanges: [range],
    },
    translated: entry.resolution.translated || rowDelta !== 0 || colDelta !== 0,
    rowDelta: entry.resolution.rowDelta + rowDelta,
    colDelta: entry.resolution.colDelta + colDelta,
  }
}

export function initialFormulaFamilyShapeKey(formula: RuntimeFormula): string {
  return buildFormulaFamilyShapeKey({
    compiled: formula.compiled,
    dependencyCount: formula.dependencyIndices.length,
    rangeDependencyCount: formula.rangeDependencies.length,
    directAggregateKind: formula.directAggregate?.aggregateKind,
    directLookupKind: formula.directLookup?.kind,
    directScalarKind: formula.directScalar?.kind,
    directCriteriaKind: formula.directCriteria?.aggregateKind,
  })
}
