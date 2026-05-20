import { FormulaMode, type FormulaRecord } from '@bilig/protocol'
import {
  formatAddress,
  parseRangeAddress,
  type CompiledFormula,
  type DirectAggregateCandidate,
  type FormulaNode,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
} from '@bilig/formula'
import { parseA1RowNumber } from './a1-row-number.js'

const SIMPLE_DIRECT_AGGREGATE_RE =
  /^(?<callee>SUM|AVERAGE|AVG|COUNT|MIN|MAX)\s*\(\s*(?<range>[^(),]+:[^(),]+)\s*\)(?:\s*\+\s*(?<offset>[+-]?(?:\d+|\d*\.\d+)))?$/i
const SIMPLE_COLUMN_RANGE_RE = /^([A-Za-z]+)([1-9][0-9]*):([A-Za-z]+)([1-9][0-9]*)$/
const EMPTY_STRINGS: string[] = []
const EMPTY_PROGRAM = new Uint32Array()
const EMPTY_CONSTANTS = new Float64Array()

type DirectAggregateKind = DirectAggregateCandidate['aggregateKind']

const DIRECT_AGGREGATE_KIND_BY_CALLEE: Record<string, DirectAggregateKind> = {
  SUM: 'sum',
  AVERAGE: 'average',
  AVG: 'average',
  COUNT: 'count',
  MIN: 'min',
  MAX: 'max',
}

interface SimpleColumnRangeInfo {
  readonly address: string
  readonly startAddress: string
  readonly endAddress: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

function columnToIndex(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    const code = column.charCodeAt(index)
    value = value * 26 + (code - 64)
  }
  return value - 1
}

function tryParseSimpleColumnRange(rawRange: string): SimpleColumnRangeInfo | null | undefined {
  const match = SIMPLE_COLUMN_RANGE_RE.exec(rawRange)
  if (!match) {
    return undefined
  }
  const startColumn = match[1]!.toUpperCase()
  const endColumn = match[3]!.toUpperCase()
  const startCol = columnToIndex(startColumn)
  const endCol = columnToIndex(endColumn)
  if (endCol < startCol) {
    return undefined
  }
  const startRowNumber = parseA1RowNumber(match[2]!)
  const endRowNumber = parseA1RowNumber(match[4]!)
  if (startRowNumber === undefined || endRowNumber === undefined) {
    return null
  }
  if (endRowNumber < startRowNumber) {
    return undefined
  }
  const startAddress = `${startColumn}${startRowNumber}`
  const endAddress = `${endColumn}${endRowNumber}`
  return {
    address: `${startAddress}:${endAddress}`,
    startAddress,
    endAddress,
    startRow: startRowNumber - 1,
    endRow: endRowNumber - 1,
    startCol,
    endCol,
  }
}

export function tryCompileSimpleDirectAggregateFormula(source: string): CompiledFormula | undefined {
  const trimmedSource = source.trim()
  const normalizedSource = trimmedSource.startsWith('=') ? trimmedSource.slice(1).trim() : trimmedSource
  const match = SIMPLE_DIRECT_AGGREGATE_RE.exec(normalizedSource)
  if (!match?.groups) {
    return undefined
  }

  const callee = match.groups['callee']!.toUpperCase()
  const aggregateKind = DIRECT_AGGREGATE_KIND_BY_CALLEE[callee]!
  const resultOffset = match.groups['offset'] === undefined ? 0 : Number(match.groups['offset'])
  if (!Number.isFinite(resultOffset)) {
    return undefined
  }

  const rawRange = match.groups['range']!.trim()
  const fastRange = tryParseSimpleColumnRange(rawRange)
  if (fastRange === null) {
    return undefined
  }
  let rangeInfo: SimpleColumnRangeInfo
  if (fastRange) {
    rangeInfo = fastRange
  } else {
    let parsedRange: ReturnType<typeof parseRangeAddress>
    try {
      parsedRange = parseRangeAddress(rawRange)
    } catch {
      return undefined
    }
    if (parsedRange.kind !== 'cells' || parsedRange.sheetName !== undefined) {
      return undefined
    }
    const start = parsedRange.start as { readonly text: string; readonly row: number; readonly col: number }
    const end = parsedRange.end as { readonly text: string; readonly row: number; readonly col: number }
    if (start.col !== end.col) {
      return undefined
    }
    rangeInfo = {
      address: rawRange,
      startAddress: start.text,
      endAddress: end.text,
      startRow: start.row,
      endRow: end.row,
      startCol: start.col,
      endCol: end.col,
    }
  }

  const parsedRangeInfo: ParsedRangeReferenceInfo = {
    ...rangeInfo,
    kind: 'range',
    refKind: 'cells',
  }

  const rangeNode: FormulaNode = {
    kind: 'RangeRef',
    refKind: 'cells',
    start: rangeInfo.startAddress,
    end: rangeInfo.endAddress,
  }

  const aggregateNode: FormulaNode = {
    kind: 'CallExpr',
    callee,
    args: [rangeNode],
  }
  const ast: FormulaNode =
    resultOffset === 0
      ? aggregateNode
      : {
          kind: 'BinaryExpr',
          operator: '+',
          left: aggregateNode,
          right: { kind: 'NumberLiteral', value: resultOffset },
        }

  const directAggregateCandidate: DirectAggregateCandidate = {
    callee,
    aggregateKind,
    symbolicRangeIndex: 0,
    ...(resultOffset !== 0 ? { resultOffset } : {}),
  }

  const baseRecord: FormulaRecord = {
    id: 0,
    source: normalizedSource,
    mode: aggregateKind === 'average' ? FormulaMode.JsOnly : FormulaMode.WasmFastPath,
    depsPtr: 0,
    depsLen: 1,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 1,
    maxStackDepth: 0,
  }

  return {
    ...baseRecord,
    ast,
    optimizedAst: ast,
    astMatchesSource: true,
    directAggregateCandidate,
    deps: [rangeInfo.address],
    parsedDeps: [parsedRangeInfo satisfies ParsedDependencyReference],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan: [],
    program: EMPTY_PROGRAM,
    constants: EMPTY_CONSTANTS,
    symbolicRefs: EMPTY_STRINGS,
    parsedSymbolicRefs: [],
    symbolicRanges: [rangeInfo.address],
    parsedSymbolicRanges: [parsedRangeInfo],
    symbolicStrings: EMPTY_STRINGS,
  }
}

export function translateSimpleDirectAggregateFormula(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  source: string,
): CompiledFormula | undefined {
  const candidate = compiled.directAggregateCandidate
  const range = compiled.parsedSymbolicRanges?.[candidate?.symbolicRangeIndex ?? -1]
  if (
    candidate === undefined ||
    range === undefined ||
    compiled.parsedSymbolicRanges?.length !== 1 ||
    compiled.symbolicRanges.length !== 1 ||
    (compiled.parsedSymbolicRefs?.length ?? 0) !== 0 ||
    compiled.symbolicRefs.length !== 0 ||
    range.kind !== 'range' ||
    range.refKind !== 'cells' ||
    range.sheetName !== undefined
  ) {
    return undefined
  }
  const startRow = range.startRow + rowDelta
  const endRow = range.endRow + rowDelta
  const startCol = range.startCol + colDelta
  const endCol = range.endCol + colDelta
  if (startRow < 0 || endRow < startRow || startCol < 0 || endCol < startCol) {
    return undefined
  }

  const startAddress = formatAddress(startRow, startCol)
  const endAddress = formatAddress(endRow, endCol)
  const address = `${startAddress}:${endAddress}`
  const translatedRange: ParsedRangeReferenceInfo = {
    ...range,
    address,
    startAddress,
    endAddress,
    startRow,
    endRow,
    startCol,
    endCol,
  }
  const rangeNode: FormulaNode = {
    kind: 'RangeRef',
    refKind: 'cells',
    start: startAddress,
    end: endAddress,
  }
  const aggregateNode: FormulaNode = {
    kind: 'CallExpr',
    callee: candidate.callee,
    args: [rangeNode],
  }
  const ast: FormulaNode =
    candidate.resultOffset === undefined
      ? aggregateNode
      : {
          kind: 'BinaryExpr',
          operator: '+',
          left: aggregateNode,
          right: { kind: 'NumberLiteral', value: candidate.resultOffset },
        }

  return {
    ...compiled,
    source,
    ast,
    optimizedAst: ast,
    astMatchesSource: false,
    deps: [address],
    parsedDeps: [translatedRange satisfies ParsedDependencyReference],
    symbolicRanges: [address],
    parsedSymbolicRanges: [translatedRange],
  }
}

export function translateAnchoredPrefixDirectAggregateFormula(
  compiled: CompiledFormula,
  ownerRow: number,
  colDelta: number,
  source: string,
): CompiledFormula | undefined {
  const candidate = compiled.directAggregateCandidate
  const range = compiled.parsedSymbolicRanges?.[candidate?.symbolicRangeIndex ?? -1]
  if (
    candidate === undefined ||
    range === undefined ||
    compiled.parsedSymbolicRanges?.length !== 1 ||
    compiled.symbolicRanges.length !== 1 ||
    (compiled.parsedSymbolicRefs?.length ?? 0) !== 0 ||
    compiled.symbolicRefs.length !== 0 ||
    range.kind !== 'range' ||
    range.refKind !== 'cells' ||
    range.sheetName !== undefined ||
    range.startRow !== 0 ||
    range.startCol !== range.endCol
  ) {
    return undefined
  }
  const startCol = range.startCol + colDelta
  const endCol = range.endCol + colDelta
  if (ownerRow < 0 || startCol < 0 || endCol < startCol) {
    return undefined
  }

  const startAddress = formatAddress(0, startCol)
  const endAddress = formatAddress(ownerRow, endCol)
  const address = `${startAddress}:${endAddress}`
  const translatedRange: ParsedRangeReferenceInfo = {
    ...range,
    address,
    startAddress,
    endAddress,
    startRow: 0,
    endRow: ownerRow,
    startCol,
    endCol,
  }
  const rangeNode: FormulaNode = {
    kind: 'RangeRef',
    refKind: 'cells',
    start: startAddress,
    end: endAddress,
  }
  const aggregateNode: FormulaNode = {
    kind: 'CallExpr',
    callee: candidate.callee,
    args: [rangeNode],
  }
  const ast: FormulaNode =
    candidate.resultOffset === undefined
      ? aggregateNode
      : {
          kind: 'BinaryExpr',
          operator: '+',
          left: aggregateNode,
          right: { kind: 'NumberLiteral', value: candidate.resultOffset },
        }

  return {
    ...compiled,
    source,
    ast,
    optimizedAst: ast,
    astMatchesSource: false,
    deps: [address],
    parsedDeps: [translatedRange satisfies ParsedDependencyReference],
    symbolicRanges: [address],
    parsedSymbolicRanges: [translatedRange],
  }
}
