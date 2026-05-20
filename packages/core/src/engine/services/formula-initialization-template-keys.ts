import type {
  CompiledFormula,
  FormulaNode,
  ParsedCellReferenceInfo,
  ParsedDependencyReference,
  ParsedRangeReferenceInfo,
} from '@bilig/formula'
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

export interface InitialSimpleRowRelativeBinaryTemplate {
  readonly key: string
  readonly parsedRefs: {
    readonly symbolicRefs: string[]
    readonly parsedDeps: ParsedDependencyReference[]
    readonly parsedSymbolicRefs: ParsedCellReferenceInfo[]
  }
  readonly usesRowLiteralSuffix: boolean
}

const INITIAL_PREFIX_SUM_RE = /^SUM\(([A-Z]+)1:\1([1-9]\d*)\)$/

export interface InitialSimpleRowRelativeBinaryTemplateKey {
  readonly key: string
  readonly usesRowLiteralSuffix: boolean
}

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
  if (start >= source.length) {
    return undefined
  }
  const firstCode = source.charCodeAt(start)
  if (firstCode < 49 || firstCode > 57) {
    return undefined
  }
  let row = firstCode - 48
  let next = start + 1
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    row = row * 10 + (code - 48)
    if (!Number.isSafeInteger(row)) {
      return undefined
    }
    next += 1
  }
  return { row, next }
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

function initialReadRowLiteralSuffix(
  source: string,
  start: number,
  ownerRow: number,
): { readonly token: string; readonly next: number } | undefined {
  if (source.charCodeAt(start) !== 43) {
    return undefined
  }
  const row = initialReadRowNumber(source, start + 1)
  if (!row || row.row !== ownerRow + 1) {
    return undefined
  }
  return {
    token: '+r0',
    next: row.next,
  }
}

function initialReadRelativeCellToken(
  source: string,
  start: number,
  ownerRow: number,
  ownerCol: number,
):
  | {
      readonly token: string
      readonly next: number
      readonly ref: ParsedCellReferenceInfo
    }
  | undefined {
  const column = initialReadColumn(source, start)
  if (!column) {
    return undefined
  }
  const row = initialReadRowNumber(source, column.next)
  if (!row || row.row - 1 !== ownerRow) {
    return undefined
  }
  const col = initialColumnToIndex(column.column)
  if (col < 0) {
    return undefined
  }
  const normalizedColumn = column.column.toUpperCase()
  return {
    token: `c${col - ownerCol}`,
    next: row.next,
    ref: {
      address: `${normalizedColumn}${row.row}`,
      row: row.row - 1,
      col,
      rowAbsolute: false,
      colAbsolute: false,
    },
  }
}

export function tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  return tryBuildInitialSimpleRowRelativeBinaryTemplateKeyInfo(source, ownerRow, ownerCol)?.key
}

export function tryBuildInitialSimpleRowRelativeBinaryTemplateKeyInfo(
  source: string,
  ownerRow: number,
  ownerCol: number,
): InitialSimpleRowRelativeBinaryTemplateKey | undefined {
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
    if (rightCell.next === source.length) {
      return { key: `${left.token}${operator}${rightCell.token}`, usesRowLiteralSuffix: false }
    }
    const suffix = initialReadRowLiteralSuffix(source, rightCell.next, ownerRow)
    return suffix && suffix.next === source.length
      ? { key: `${left.token}${operator}${rightCell.token}${suffix.token}`, usesRowLiteralSuffix: true }
      : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  if (!rightNumber) {
    return undefined
  }
  if (rightNumber.next === source.length) {
    return { key: `${left.token}${operator}n${rightNumber.text}`, usesRowLiteralSuffix: false }
  }
  const suffix = initialReadRowLiteralSuffix(source, rightNumber.next, ownerRow)
  return suffix && suffix.next === source.length
    ? { key: `${left.token}${operator}n${rightNumber.text}${suffix.token}`, usesRowLiteralSuffix: true }
    : undefined
}

export function tryBuildInitialSimpleRowRelativeBinaryTemplate(
  source: string,
  ownerRow: number,
  ownerCol: number,
): InitialSimpleRowRelativeBinaryTemplate | undefined {
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
    if (rightCell.next === source.length) {
      return createInitialSimpleBinaryTemplate(`${left.token}${operator}${rightCell.token}`, [left.ref, rightCell.ref], false)
    }
    const suffix = initialReadRowLiteralSuffix(source, rightCell.next, ownerRow)
    return suffix && suffix.next === source.length
      ? createInitialSimpleBinaryTemplate(`${left.token}${operator}${rightCell.token}${suffix.token}`, [left.ref, rightCell.ref], true)
      : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  if (!rightNumber) {
    return undefined
  }
  if (rightNumber.next === source.length) {
    return createInitialSimpleBinaryTemplate(`${left.token}${operator}n${rightNumber.text}`, [left.ref], false)
  }
  const suffix = initialReadRowLiteralSuffix(source, rightNumber.next, ownerRow)
  return suffix && suffix.next === source.length
    ? createInitialSimpleBinaryTemplate(`${left.token}${operator}n${rightNumber.text}${suffix.token}`, [left.ref], true)
    : undefined
}

function createInitialSimpleBinaryTemplate(
  key: string,
  parsedSymbolicRefs: ParsedCellReferenceInfo[],
  usesRowLiteralSuffix: boolean,
): InitialSimpleRowRelativeBinaryTemplate {
  const symbolicRefs: string[] = Array(parsedSymbolicRefs.length)
  const parsedDeps: ParsedDependencyReference[] = Array(parsedSymbolicRefs.length)
  for (let index = 0; index < parsedSymbolicRefs.length; index += 1) {
    const ref = parsedSymbolicRefs[index]!
    symbolicRefs[index] = ref.address
    parsedDeps[index] = { kind: 'cell', ...ref }
  }
  return {
    key,
    parsedRefs: {
      symbolicRefs,
      parsedDeps,
      parsedSymbolicRefs,
    },
    usesRowLiteralSuffix,
  }
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
