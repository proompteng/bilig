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

type DirectAggregateKind = DirectAggregateCandidate['aggregateKind']

const DIRECT_AGGREGATE_KIND_BY_CALLEE: Record<string, DirectAggregateKind> = {
  SUM: 'sum',
  AVERAGE: 'average',
  AVG: 'average',
  COUNT: 'count',
  MIN: 'min',
  MAX: 'max',
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
  const parsedRange = parseRangeAddress(rawRange)
  if (parsedRange.kind !== 'cells' || parsedRange.start.col !== parsedRange.end.col || parsedRange.sheetName !== undefined) {
    return undefined
  }

  const parsedRangeInfo: ParsedRangeReferenceInfo = {
    address: rawRange,
    kind: 'range',
    refKind: 'cells',
    startAddress: parsedRange.start.text,
    endAddress: parsedRange.end.text,
    startRow: parsedRange.start.row,
    endRow: parsedRange.end.row,
    startCol: parsedRange.start.col,
    endCol: parsedRange.end.col,
  }

  const rangeNode: FormulaNode = {
    kind: 'RangeRef',
    refKind: 'cells',
    start: parsedRange.start.text,
    end: parsedRange.end.text,
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
    deps: [rawRange],
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
    symbolicRanges: [rawRange],
    parsedSymbolicRanges: [parsedRangeInfo],
    symbolicStrings: [],
  }
}
