import { BuiltinId, FormulaMode, Opcode, type FormulaRecord } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import { columnToIndex, formatRangeAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import { getNativeGroupedArrayKind } from './binder-wasm-rules.js'
import { bindFormula, encodeBuiltin } from './binder.js'
import { lowerToPlan, type JsPlanInstruction } from './js-evaluator.js'
import { optimizeFormula } from './optimizer.js'
import { parseFormula } from './parser.js'
import { rewriteSpecialCall } from './special-call-rewrites.js'

function encodeInstruction(opcode: Opcode, operand = 0): number {
  return (opcode << 24) | (operand & 0x00ff_ffff)
}

interface CompilerState {
  program: number[]
  constants: number[]
  refs: string[]
  ranges: string[]
  strings: string[]
}

const SIMPLE_CELL_REF_RE = /^\$?[A-Z]+\$?[1-9][0-9]*$/
const SIMPLE_NUMBER_RE = /^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?$/
const SIMPLE_BINARY_RE =
  /^\s*(\$?[A-Z]+\$?[1-9][0-9]*|[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?)\s*([+*])\s*(\$?[A-Z]+\$?[1-9][0-9]*|[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?)\s*$/

const VOLATILE_BUILTINS = new Set(['TODAY', 'NOW', 'RAND', 'RANDBETWEEN', 'RANDARRAY'])

interface VolatileMetadata {
  volatile: boolean
  randCallCount: number
}

function producesSpillResult(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'InvokeExpr':
      return false
    case 'RangeRef':
      return true
    case 'UnaryExpr':
      return producesSpillResult(node.argument)
    case 'BinaryExpr':
      return producesSpillResult(node.left) || producesSpillResult(node.right)
    case 'CallExpr':
      if (node.callee.toUpperCase() === 'CHOOSE') {
        return node.args.slice(1).some((arg) => producesSpillResult(arg))
      }
      if (node.callee.toUpperCase() === 'TREND' || node.callee.toUpperCase() === 'GROWTH') {
        const shapeArg = node.args[2] ?? node.args[1] ?? node.args[0]
        if (shapeArg === undefined) {
          return false
        }
        return producesSpillResult(shapeArg)
      }
      return [
        'SEQUENCE',
        'EXPAND',
        'LINEST',
        'LOGEST',
        'OFFSET',
        'TAKE',
        'DROP',
        'CHOOSECOLS',
        'CHOOSEROWS',
        'SORT',
        'SORTBY',
        'TOCOL',
        'TOROW',
        'WRAPROWS',
        'WRAPCOLS',
        'FILTER',
        'UNIQUE',
        'FREQUENCY',
        'MODE.MULT',
        'TEXTSPLIT',
        'TRIMRANGE',
        'GROUPBY',
        'PIVOTBY',
        'MAKEARRAY',
        'MAP',
        'SCAN',
        'BYROW',
        'BYCOL',
        'RANDARRAY',
        'MUNIT',
        'MINVERSE',
        'MMULT',
      ].includes(node.callee.toUpperCase())
  }
}

function analyzeVolatileMetadata(node: FormulaNode): VolatileMetadata {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return { volatile: false, randCallCount: 0 }
    case 'UnaryExpr':
      return analyzeVolatileMetadata(node.argument)
    case 'BinaryExpr': {
      const left = analyzeVolatileMetadata(node.left)
      const right = analyzeVolatileMetadata(node.right)
      return {
        volatile: left.volatile || right.volatile,
        randCallCount: left.randCallCount + right.randCallCount,
      }
    }
    case 'CallExpr': {
      const rewritten = rewriteSpecialCall(node)
      if (rewritten) {
        return analyzeVolatileMetadata(rewritten)
      }
      const callee = node.callee.toUpperCase()
      let volatile = VOLATILE_BUILTINS.has(callee)
      let randCallCount = callee === 'RAND' ? 1 : 0
      node.args.forEach((arg) => {
        const child = analyzeVolatileMetadata(arg)
        volatile = volatile || child.volatile
        randCallCount += child.randCallCount
      })
      return { volatile, randCallCount }
    }
    case 'InvokeExpr': {
      const callee = analyzeVolatileMetadata(node.callee)
      const args = node.args.map(analyzeVolatileMetadata)
      return {
        volatile: callee.volatile || args.some((child) => child.volatile),
        randCallCount: callee.randCallCount + args.reduce((sum, child) => sum + child.randCallCount, 0),
      }
    }
  }
}

function emitCellRef(ref: string, sheetName: string | undefined, state: CompilerState): void {
  const qualifiedRef = sheetName ? `${sheetName}!${ref}` : ref
  let index = state.refs.indexOf(qualifiedRef)
  if (index === -1) index = state.refs.push(qualifiedRef) - 1
  state.program.push(encodeInstruction(Opcode.PushCell, index))
}

function emitRangeRef(node: Extract<FormulaNode, { kind: 'RangeRef' }>, state: CompilerState): void {
  const qualifiedRange = formatRangeAddress(
    parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`),
  )
  let index = state.ranges.indexOf(qualifiedRange)
  if (index === -1) {
    index = state.ranges.push(qualifiedRange) - 1
  }
  state.program.push(encodeInstruction(Opcode.PushRange, index))
}

function isCellRangeNode(node: FormulaNode): node is Extract<FormulaNode, { kind: 'RangeRef' }> {
  if (node.kind !== 'RangeRef') {
    return false
  }
  try {
    const sheetPrefix = node.sheetName ? `${node.sheetName}!` : ''
    return parseRangeAddress(`${sheetPrefix}${node.start}:${node.end}`).kind === 'cells'
  } catch {
    return false
  }
}

function emitArgument(node: FormulaNode, state: CompilerState): number {
  if (node.kind === 'RangeRef') {
    emitRangeRef(node, state)
    return 1
  }

  emitNode(node, state)
  return 1
}

const AXIS_AGGREGATE_CODES = new Map<string, number>([
  ['SUM', 1],
  ['AVERAGE', 2],
  ['AVG', 2],
  ['MIN', 3],
  ['MAX', 4],
  ['COUNT', 5],
  ['COUNTA', 6],
])

function getNativeAxisAggregateCode(node: FormulaNode): number | null {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 2) {
    return null
  }
  const [param, body] = node.args
  if (param?.kind !== 'NameRef' || body?.kind !== 'CallExpr' || body.args.length !== 1) {
    return null
  }
  const aggregateCode = AXIS_AGGREGATE_CODES.get(body.callee.toUpperCase())
  if (aggregateCode === undefined) {
    return null
  }
  return body.args[0]?.kind === 'NameRef' && body.args[0].name.trim().toUpperCase() === param.name.trim().toUpperCase()
    ? aggregateCode
    : null
}

function getNativeRunningFoldCode(node: FormulaNode): number | null {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 3) {
    return null
  }
  const [acc, value, body] = node.args
  if (acc?.kind !== 'NameRef' || value?.kind !== 'NameRef' || body?.kind !== 'BinaryExpr') {
    return null
  }
  const foldCode = body.operator === '+' ? 1 : body.operator === '*' ? 2 : null
  if (foldCode === null) {
    return null
  }
  const left = body.left
  const right = body.right
  const accName = acc.name.trim().toUpperCase()
  const valueName = value.name.trim().toUpperCase()
  return left.kind === 'NameRef' &&
    right.kind === 'NameRef' &&
    ((left.name.trim().toUpperCase() === accName && right.name.trim().toUpperCase() === valueName) ||
      (left.name.trim().toUpperCase() === valueName && right.name.trim().toUpperCase() === accName))
    ? foldCode
    : null
}

function isNativeMakearraySumLambda(node: FormulaNode): boolean {
  if (node.kind !== 'CallExpr' || node.callee.toUpperCase() !== 'LAMBDA' || node.args.length !== 3) {
    return false
  }
  const [rowParam, colParam, body] = node.args
  if (rowParam?.kind !== 'NameRef' || colParam?.kind !== 'NameRef' || body?.kind !== 'BinaryExpr') {
    return false
  }
  if (body.operator !== '+') {
    return false
  }
  const left = body.left
  const right = body.right
  const rowName = rowParam.name.trim().toUpperCase()
  const colName = colParam.name.trim().toUpperCase()
  return (
    left.kind === 'NameRef' &&
    right.kind === 'NameRef' &&
    ((left.name.trim().toUpperCase() === rowName && right.name.trim().toUpperCase() === colName) ||
      (left.name.trim().toUpperCase() === colName && right.name.trim().toUpperCase() === rowName))
  )
}

function emitNode(node: FormulaNode, state: CompilerState): void {
  switch (node.kind) {
    case 'NumberLiteral': {
      const index = state.constants.push(node.value) - 1
      state.program.push(encodeInstruction(Opcode.PushNumber, index))
      return
    }
    case 'BooleanLiteral':
      state.program.push(encodeInstruction(Opcode.PushBoolean, node.value ? 1 : 0))
      return
    case 'StringLiteral':
      {
        let index = state.strings.indexOf(node.value)
        if (index === -1) index = state.strings.push(node.value) - 1
        state.program.push(encodeInstruction(Opcode.PushString, index))
      }
      return
    case 'ErrorLiteral':
      state.program.push(encodeInstruction(Opcode.PushError, node.code))
      return
    case 'NameRef':
    case 'StructuredRef':
    case 'SpillRef':
      throw new Error('Defined names are not supported on the wasm fast path')
    case 'CellRef': {
      emitCellRef(node.ref, node.sheetName, state)
      return
    }
    case 'RangeRef':
      emitRangeRef(node, state)
      return
    case 'RowRef':
    case 'ColumnRef':
      throw new Error('Row and column references must appear inside a range')
    case 'UnaryExpr':
      emitNode(node.argument, state)
      if (node.operator === '-') {
        state.program.push(encodeInstruction(Opcode.Neg))
      }
      return
    case 'BinaryExpr':
      emitNode(node.left, state)
      emitNode(node.right, state)
      state.program.push(
        encodeInstruction(
          {
            '+': Opcode.Add,
            '-': Opcode.Sub,
            '*': Opcode.Mul,
            '/': Opcode.Div,
            '^': Opcode.Pow,
            '&': Opcode.Concat,
            '=': Opcode.Eq,
            '<>': Opcode.Neq,
            '>': Opcode.Gt,
            '>=': Opcode.Gte,
            '<': Opcode.Lt,
            '<=': Opcode.Lte,
          }[node.operator],
        ),
      )
      return
    case 'CallExpr':
      {
        const rewritten = rewriteSpecialCall(node)
        if (rewritten) {
          emitNode(rewritten, state)
          return
        }
        const callee = node.callee.toUpperCase()
        const nativeGroupedArrayKind = getNativeGroupedArrayKind(node)
        if (nativeGroupedArrayKind === 'groupby-sum-canonical') {
          emitArgument(node.args[0]!, state)
          emitArgument(node.args[1]!, state)
          state.program.push(encodeInstruction(Opcode.CallBuiltin, (BuiltinId.GroupbySumCanonical << 8) | 2))
          return
        }
        if (nativeGroupedArrayKind === 'pivotby-sum-canonical') {
          emitArgument(node.args[0]!, state)
          emitArgument(node.args[1]!, state)
          emitArgument(node.args[2]!, state)
          state.program.push(encodeInstruction(Opcode.CallBuiltin, (BuiltinId.PivotbySumCanonical << 8) | 3))
          return
        }
        if (callee === 'IF') {
          if (node.args.length !== 3) {
            throw new Error('IF requires exactly three arguments on the wasm fast path')
          }
          emitNode(node.args[0]!, state)
          const jumpIfFalseIndex = state.program.push(encodeInstruction(Opcode.JumpIfFalse, 0)) - 1
          emitNode(node.args[1]!, state)
          const jumpIndex = state.program.push(encodeInstruction(Opcode.Jump, 0)) - 1
          const falseBranchStart = state.program.length
          emitNode(node.args[2]!, state)
          const end = state.program.length
          state.program[jumpIfFalseIndex] = encodeInstruction(Opcode.JumpIfFalse, falseBranchStart)
          state.program[jumpIndex] = encodeInstruction(Opcode.Jump, end)
          return
        }
        if ((callee === 'BYROW' || callee === 'BYCOL') && node.args.length === 2) {
          const lambda = node.args[1]!
          const aggregateCode = getNativeAxisAggregateCode(lambda)
          if (aggregateCode !== null) {
            const aggregateIndex = state.constants.push(aggregateCode) - 1
            state.program.push(encodeInstruction(Opcode.PushNumber, aggregateIndex))
            emitArgument(node.args[0]!, state)
            state.program.push(
              encodeInstruction(Opcode.CallBuiltin, ((callee === 'BYROW' ? BuiltinId.ByrowAggregate : BuiltinId.BycolAggregate) << 8) | 2),
            )
            return
          }
        }
        if (callee === 'MAKEARRAY' && node.args.length === 3 && isNativeMakearraySumLambda(node.args[2]!)) {
          emitArgument(node.args[0]!, state)
          emitArgument(node.args[1]!, state)
          state.program.push(encodeInstruction(Opcode.CallBuiltin, (BuiltinId.MakearraySum << 8) | 2))
          return
        }
        if (callee === 'REDUCE' || callee === 'SCAN') {
          const sourceArg = node.args.length === 3 ? node.args[1] : node.args[0]
          const lambdaArg = node.args.length === 3 ? node.args[2] : node.args[1]
          const initialArg = node.args.length === 3 ? node.args[0] : undefined
          const foldCode = lambdaArg ? getNativeRunningFoldCode(lambdaArg) : null
          if (
            (node.args.length === 2 || node.args.length === 3) &&
            sourceArg !== undefined &&
            lambdaArg !== undefined &&
            foldCode !== null
          ) {
            let argc = 0
            if (initialArg !== undefined) {
              argc += emitArgument(initialArg, state)
            }
            argc += emitArgument(sourceArg, state)
            state.program.push(
              encodeInstruction(
                Opcode.CallBuiltin,
                ((callee === 'REDUCE'
                  ? foldCode === 1
                    ? BuiltinId.ReduceSum
                    : BuiltinId.ReduceProduct
                  : foldCode === 1
                    ? BuiltinId.ScanSum
                    : BuiltinId.ScanProduct) <<
                  8) |
                  argc,
              ),
            )
            return
          }
        }
        const rangeArg = node.args[0]
        if (callee === 'PHONETIC' && node.args.length === 1 && rangeArg && isCellRangeNode(rangeArg)) {
          emitCellRef(rangeArg.start, rangeArg.sheetName, state)
          state.program.push(encodeInstruction(Opcode.CallBuiltin, (BuiltinId.Phonetic << 8) | 1))
          return
        }
        let argc = 0
        node.args.forEach((arg) => {
          argc += emitArgument(arg, state)
        })
        state.program.push(encodeInstruction(Opcode.CallBuiltin, (encodeBuiltin(callee) << 8) | argc))
      }
      return
    case 'InvokeExpr':
      throw new Error('Lambda invocation is not supported on the wasm fast path')
  }
}

export interface CompiledFormula extends FormulaRecord {
  ast: FormulaNode
  optimizedAst: FormulaNode
  astMatchesSource?: boolean
  directAggregateCandidate?: DirectAggregateCandidate
  deps: string[]
  parsedDeps?: ParsedDependencyReference[]
  symbolicNames: string[]
  symbolicTables: string[]
  symbolicSpills: string[]
  volatile: boolean
  randCallCount: number
  producesSpill: boolean
  jsPlan: JsPlanInstruction[]
  program: Uint32Array
  constants: Float64Array
  symbolicRefs: string[]
  parsedSymbolicRefs?: ParsedCellReferenceInfo[]
  symbolicRanges: string[]
  parsedSymbolicRanges?: ParsedRangeReferenceInfo[]
  symbolicStrings: string[]
}

export interface ParsedCellReferenceInfo {
  address: string
  sheetName?: string
  explicitSheet?: boolean
  row?: number
  col?: number
  rowAbsolute?: boolean
  colAbsolute?: boolean
}

export interface ParsedRangeReferenceInfo {
  address: string
  kind: 'range'
  refKind: 'cells' | 'rows' | 'cols'
  sheetName?: string
  explicitSheet?: boolean
  startAddress: string
  endAddress: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
  startRowAbsolute?: boolean
  endRowAbsolute?: boolean
  startColAbsolute?: boolean
  endColAbsolute?: boolean
}

export type ParsedDependencyReference = ({ kind: 'cell' } & ParsedCellReferenceInfo) | ParsedRangeReferenceInfo

export interface DirectAggregateCandidate {
  callee: string
  aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
  symbolicRangeIndex: number
}

interface CompileFormulaAstOptions {
  originalAst?: FormulaNode
  symbolicNames?: string[]
  symbolicTables?: string[]
  symbolicSpills?: string[]
}

const CELL_REFERENCE_PARTS_RE = /^(\$?)([A-Z]+)(\$?)([1-9][0-9]*)$/i
const COLUMN_REFERENCE_PARTS_RE = /^(\$?)([A-Z]+)$/i
const ROW_REFERENCE_PARTS_RE = /^(\$?)([1-9][0-9]*)$/

interface CellReferenceParts {
  row: number
  col: number
  rowAbsolute: boolean
  colAbsolute: boolean
}

interface AxisReferenceParts {
  index: number
  absolute: boolean
}

function computeMaxStackDepth(plan: readonly JsPlanInstruction[]): number {
  let current = 0
  let max = 0
  for (const instruction of plan) {
    switch (instruction.opcode) {
      case 'push-number':
      case 'push-boolean':
      case 'push-string':
      case 'push-error':
      case 'push-name':
      case 'push-cell':
      case 'push-range':
      case 'push-lambda':
        current += 1
        break
      case 'lookup-exact-match':
      case 'lookup-approximate-match':
        break
      case 'binary':
        current -= 1
        break
      case 'call':
        current -= instruction.argc
        current += 1
        break
      case 'invoke':
        current -= instruction.argc
        current += 1
        break
      case 'jump-if-false':
        current -= 1
        break
      case 'bind-name':
        current -= 1
        break
      case 'unary':
      case 'begin-scope':
      case 'end-scope':
      case 'jump':
      case 'return':
        break
    }
    max = Math.max(max, current)
  }
  return max
}

function parseSimpleOperand(source: string): FormulaNode | null {
  const trimmed = source.trim()
  if (SIMPLE_CELL_REF_RE.test(trimmed)) {
    return {
      kind: 'CellRef',
      ref: trimmed,
    }
  }
  if (SIMPLE_NUMBER_RE.test(trimmed)) {
    return {
      kind: 'NumberLiteral',
      value: Number(trimmed),
    }
  }
  return null
}

function stripSheetQualifier(reference: string): string {
  const bang = reference.lastIndexOf('!')
  return bang === -1 ? reference.trim() : reference.slice(bang + 1).trim()
}

function parseCellReferenceParts(reference: string): CellReferenceParts | undefined {
  const match = CELL_REFERENCE_PARTS_RE.exec(stripSheetQualifier(reference))
  if (!match) {
    return undefined
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match
  return {
    row: Number.parseInt(rowText!, 10) - 1,
    col: columnToIndex(columnText!),
    rowAbsolute: rowAbsolute === '$',
    colAbsolute: colAbsolute === '$',
  }
}

function parseAxisReferenceParts(reference: string, kind: 'row' | 'column'): AxisReferenceParts | undefined {
  const match = (kind === 'row' ? ROW_REFERENCE_PARTS_RE : COLUMN_REFERENCE_PARTS_RE).exec(stripSheetQualifier(reference))
  if (!match) {
    return undefined
  }
  return kind === 'row'
    ? {
        index: Number.parseInt(match[2]!, 10) - 1,
        absolute: match[1] === '$',
      }
    : {
        index: columnToIndex(match[2]!),
        absolute: match[1] === '$',
      }
}

function buildParsedCellReferenceInfo(reference: string): ParsedCellReferenceInfo {
  const parsedCell = parseCellAddress(reference)
  const parts = parseCellReferenceParts(reference)
  return {
    address: reference,
    ...(parsedCell.sheetName !== undefined ? { sheetName: parsedCell.sheetName } : {}),
    ...(reference.includes('!') ? { explicitSheet: true } : {}),
    row: parsedCell.row,
    col: parsedCell.col,
    ...(parts
      ? {
          rowAbsolute: parts.rowAbsolute,
          colAbsolute: parts.colAbsolute,
        }
      : {}),
  }
}

function buildParsedRangeReferenceInfo(reference: string): ParsedRangeReferenceInfo {
  const parsedRange = parseRangeAddress(reference)
  const bounds =
    parsedRange.kind === 'cells'
      ? {
          startRow: parsedRange.start.row,
          endRow: parsedRange.end.row,
          startCol: parsedRange.start.col,
          endCol: parsedRange.end.col,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRow: parsedRange.start.row,
            endRow: parsedRange.end.row,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: parsedRange.start.col,
            endCol: parsedRange.end.col,
          }
  const separator = reference.indexOf(':')
  const rawStart = reference.slice(0, separator).trim()
  const rawEnd = reference.slice(separator + 1).trim()
  const cellStart = parsedRange.kind === 'cells' ? parseCellReferenceParts(rawStart) : undefined
  const cellEnd = parsedRange.kind === 'cells' ? parseCellReferenceParts(rawEnd) : undefined
  const rowStart = parsedRange.kind === 'rows' ? parseAxisReferenceParts(rawStart, 'row') : undefined
  const rowEnd = parsedRange.kind === 'rows' ? parseAxisReferenceParts(rawEnd, 'row') : undefined
  const colStart = parsedRange.kind === 'cols' ? parseAxisReferenceParts(rawStart, 'column') : undefined
  const colEnd = parsedRange.kind === 'cols' ? parseAxisReferenceParts(rawEnd, 'column') : undefined
  return {
    address: reference,
    kind: 'range',
    refKind: parsedRange.kind,
    ...(parsedRange.sheetName !== undefined ? { sheetName: parsedRange.sheetName } : {}),
    ...(rawStart.includes('!') ? { explicitSheet: true } : {}),
    startAddress: parsedRange.start.text,
    endAddress: parsedRange.end.text,
    ...bounds,
    ...(parsedRange.kind === 'cells'
      ? {
          startRowAbsolute: cellStart?.rowAbsolute ?? false,
          endRowAbsolute: cellEnd?.rowAbsolute ?? false,
          startColAbsolute: cellStart?.colAbsolute ?? false,
          endColAbsolute: cellEnd?.colAbsolute ?? false,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRowAbsolute: rowStart?.absolute ?? false,
            endRowAbsolute: rowEnd?.absolute ?? false,
          }
        : {
            startColAbsolute: colStart?.absolute ?? false,
            endColAbsolute: colEnd?.absolute ?? false,
          }),
  }
}

function parseDependencyReference(reference: string): ParsedDependencyReference {
  if (reference.includes(':')) {
    return buildParsedRangeReferenceInfo(reference)
  }
  return {
    kind: 'cell',
    ...buildParsedCellReferenceInfo(reference),
  }
}

function buildSimpleCompiledFormula(source: string): CompiledFormula | null {
  const trimmed = source.trim()
  const singleOperand = parseSimpleOperand(trimmed)
  const refs: string[] = []
  const parsedRefs: ParsedCellReferenceInfo[] = []
  const deps: string[] = []
  const parsedDeps: ParsedDependencyReference[] = []
  const constants: number[] = []
  const program: number[] = []

  const registerCellRef = (ref: string): number => {
    const existing = refs.indexOf(ref)
    if (existing !== -1) {
      return existing
    }
    const index = refs.push(ref) - 1
    parsedRefs.push(buildParsedCellReferenceInfo(ref))
    deps.push(ref)
    parsedDeps.push({ kind: 'cell', ...buildParsedCellReferenceInfo(ref) })
    return index
  }

  const emitOperand = (operand: FormulaNode, plan: JsPlanInstruction[]): void => {
    if (operand.kind === 'CellRef') {
      const index = registerCellRef(operand.ref)
      program.push(encodeInstruction(Opcode.PushCell, index))
      plan.push({ opcode: 'push-cell', address: operand.ref })
      return
    }
    if (operand.kind === 'NumberLiteral') {
      const index = constants.push(operand.value) - 1
      program.push(encodeInstruction(Opcode.PushNumber, index))
      plan.push({ opcode: 'push-number', value: operand.value })
      return
    }
    throw new Error(`Unsupported simple operand '${operand.kind}'`)
  }

  if (singleOperand?.kind === 'CellRef') {
    const jsPlan: JsPlanInstruction[] = []
    emitOperand(singleOperand, jsPlan)
    jsPlan.push({ opcode: 'return' })
    program.push(encodeInstruction(Opcode.Ret))
    return {
      id: 0,
      source,
      mode: FormulaMode.WasmFastPath,
      depsPtr: 0,
      depsLen: 0,
      programOffset: 0,
      programLength: program.length,
      constNumberOffset: 0,
      constNumberLength: 0,
      rangeListOffset: 0,
      rangeListLength: 0,
      program: Uint32Array.from(program),
      constants: Float64Array.from(constants),
      symbolicRefs: refs,
      parsedSymbolicRefs: parsedRefs,
      symbolicRanges: [],
      symbolicStrings: [],
      ast: singleOperand,
      optimizedAst: singleOperand,
      astMatchesSource: true,
      deps,
      parsedDeps,
      symbolicNames: [],
      symbolicTables: [],
      symbolicSpills: [],
      volatile: false,
      randCallCount: 0,
      producesSpill: false,
      jsPlan,
      maxStackDepth: computeMaxStackDepth(jsPlan),
    }
  }

  const binaryMatch = trimmed.match(SIMPLE_BINARY_RE)
  if (!binaryMatch) {
    return null
  }
  const left = parseSimpleOperand(binaryMatch[1] ?? '')
  const operatorCandidate = binaryMatch[2]
  const operator = operatorCandidate === '+' || operatorCandidate === '*' ? operatorCandidate : undefined
  const right = parseSimpleOperand(binaryMatch[3] ?? '')
  if (!left || !right || !operator) {
    return null
  }
  const ast: FormulaNode = {
    kind: 'BinaryExpr',
    operator,
    left,
    right,
  }
  const jsPlan: JsPlanInstruction[] = []
  emitOperand(left, jsPlan)
  emitOperand(right, jsPlan)
  jsPlan.push({ opcode: 'binary', operator })
  jsPlan.push({ opcode: 'return' })
  program.push(encodeInstruction(operator === '+' ? Opcode.Add : Opcode.Mul))
  program.push(encodeInstruction(Opcode.Ret))
  return {
    id: 0,
    source,
    mode: FormulaMode.WasmFastPath,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: program.length,
    constNumberOffset: 0,
    constNumberLength: constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    program: Uint32Array.from(program),
    constants: Float64Array.from(constants),
    symbolicRefs: refs,
    parsedSymbolicRefs: parsedRefs,
    symbolicRanges: [],
    symbolicStrings: [],
    ast,
    optimizedAst: ast,
    astMatchesSource: true,
    deps,
    parsedDeps,
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan,
    maxStackDepth: computeMaxStackDepth(jsPlan),
  }
}

function buildDirectAggregateCandidate(node: FormulaNode, symbolicRanges: readonly string[]): DirectAggregateCandidate | undefined {
  if (node.kind !== 'CallExpr' || node.args.length !== 1) {
    return undefined
  }
  const callee = node.callee.trim().toUpperCase()
  const aggregateKind =
    callee === 'SUM'
      ? 'sum'
      : callee === 'COUNT'
        ? 'count'
        : callee === 'MIN'
          ? 'min'
          : callee === 'MAX'
            ? 'max'
            : callee === 'AVERAGE' || callee === 'AVG'
              ? 'average'
              : undefined
  if (!aggregateKind) {
    return undefined
  }
  const rangeNode = node.args[0]
  if (!rangeNode || rangeNode.kind !== 'RangeRef' || rangeNode.refKind !== 'cells') {
    return undefined
  }
  const qualifiedRange = formatRangeAddress(
    parseRangeAddress(
      rangeNode.sheetName ? `${rangeNode.sheetName}!${rangeNode.start}:${rangeNode.end}` : `${rangeNode.start}:${rangeNode.end}`,
    ),
  )
  const symbolicRangeIndex = symbolicRanges.indexOf(qualifiedRange)
  if (symbolicRangeIndex === -1) {
    return undefined
  }
  return {
    callee: node.callee,
    aggregateKind,
    symbolicRangeIndex,
  }
}

export function compileFormulaAst(source: string, ast: FormulaNode, options: CompileFormulaAstOptions = {}): CompiledFormula {
  const optimizedAst = optimizeFormula(ast)
  const bound = bindFormula(optimizedAst)
  const state: CompilerState = { program: [], constants: [], refs: [], ranges: [], strings: [] }
  const jsPlan = lowerToPlan(optimizedAst)
  const volatileMetadata = analyzeVolatileMetadata(options.originalAst ?? ast)
  const spillResult = producesSpillResult(optimizedAst)
  const directAggregateCandidate = buildDirectAggregateCandidate(optimizedAst, state.ranges)

  if (bound.mode === FormulaMode.WasmFastPath) {
    emitNode(optimizedAst, state)
  }

  state.program.push(encodeInstruction(Opcode.Ret))
  return {
    id: 0,
    source,
    mode: bound.mode,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: state.program.length,
    constNumberOffset: 0,
    constNumberLength: state.constants.length,
    rangeListOffset: 0,
    rangeListLength: state.ranges.length,
    program: Uint32Array.from(state.program),
    constants: Float64Array.from(state.constants),
    symbolicRefs: state.refs,
    parsedSymbolicRefs: state.refs.map((reference) => buildParsedCellReferenceInfo(reference)),
    symbolicRanges: state.ranges,
    parsedSymbolicRanges: state.ranges.map((reference) => buildParsedRangeReferenceInfo(reference)),
    symbolicStrings: state.strings,
    ast: options.originalAst ?? ast,
    optimizedAst,
    astMatchesSource: true,
    ...(directAggregateCandidate ? { directAggregateCandidate } : {}),
    deps: bound.deps,
    parsedDeps: bound.deps.map(parseDependencyReference),
    symbolicNames: options.symbolicNames ?? bound.symbolicNames,
    symbolicTables: options.symbolicTables ?? bound.symbolicTables,
    symbolicSpills: options.symbolicSpills ?? bound.symbolicSpills,
    volatile: volatileMetadata.volatile,
    randCallCount: volatileMetadata.randCallCount,
    producesSpill: spillResult,
    jsPlan,
    maxStackDepth: computeMaxStackDepth(jsPlan),
  }
}

export function compileFormula(source: string): CompiledFormula {
  return buildSimpleCompiledFormula(source) ?? compileFormulaAst(source, parseFormula(source))
}
