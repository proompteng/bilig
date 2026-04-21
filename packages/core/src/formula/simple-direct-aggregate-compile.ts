import { FormulaMode, type FormulaRecord } from '@bilig/protocol'
import {
  parseRangeAddress,
  type CompiledFormula,
  type DirectAggregateCandidate,
  type FormulaNode,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
} from '@bilig/formula'

const SIMPLE_DIRECT_AGGREGATE_RE = /^(?<callee>SUM|AVERAGE|AVG|COUNT|MIN|MAX)\s*\(\s*(?<range>[^(),]+:[^(),]+)\s*\)$/i
const SIMPLE_COLUMN_RANGE_RE = /^([A-Za-z]+)([1-9][0-9]*):([A-Za-z]+)([1-9][0-9]*)$/

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
    if (code < 65 || code > 90) {
      return -1
    }
    value = value * 26 + (code - 64)
  }
  return value - 1
}

function tryParseSimpleColumnRange(rawRange: string): SimpleColumnRangeInfo | undefined {
  const match = SIMPLE_COLUMN_RANGE_RE.exec(rawRange)
  if (!match) {
    return undefined
  }
  const startColumn = match[1]!.toUpperCase()
  const endColumn = match[3]!.toUpperCase()
  if (startColumn !== endColumn) {
    return undefined
  }
  const startCol = columnToIndex(startColumn)
  if (startCol < 0) {
    return undefined
  }
  const startRowNumber = Number.parseInt(match[2]!, 10)
  const endRowNumber = Number.parseInt(match[4]!, 10)
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
    endCol: startCol,
  }
}

export function tryCompileSimpleDirectAggregateFormula(source: string): CompiledFormula | undefined {
  const match = SIMPLE_DIRECT_AGGREGATE_RE.exec(source.trim())
  if (!match?.groups) {
    return undefined
  }

  const callee = match.groups['callee']!.toUpperCase()
  const aggregateKind = DIRECT_AGGREGATE_KIND_BY_CALLEE[callee]
  if (!aggregateKind) {
    return undefined
  }

  const rawRange = match.groups['range']!.trim()
  const fastRange = tryParseSimpleColumnRange(rawRange)
  let rangeInfo: SimpleColumnRangeInfo
  if (fastRange) {
    rangeInfo = fastRange
  } else {
    const parsedRange = parseRangeAddress(rawRange)
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

  const ast: FormulaNode = {
    kind: 'CallExpr',
    callee,
    args: [rangeNode],
  }

  const directAggregateCandidate: DirectAggregateCandidate = {
    callee,
    aggregateKind,
    symbolicRangeIndex: 0,
  }

  const baseRecord: FormulaRecord = {
    id: 0,
    source,
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
    program: new Uint32Array(),
    constants: new Float64Array(),
    symbolicRefs: [],
    parsedSymbolicRefs: [],
    symbolicRanges: [rangeInfo.address],
    parsedSymbolicRanges: [parsedRangeInfo],
    symbolicStrings: [],
  }
}
