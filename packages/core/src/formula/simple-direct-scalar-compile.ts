import { FormulaMode, Opcode, type FormulaRecord } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode, JsPlanInstruction, ParsedCellReferenceInfo, ParsedDependencyReference } from '@bilig/formula'

const SIMPLE_DIRECT_BINARY_RE = /^([A-Za-z]+)([1-9][0-9]*)([+\-*/])(?:([A-Za-z]+)([1-9][0-9]*)|(\d+(?:\.\d+)?))$/

type SimpleDirectBinaryOperator = '+' | '-' | '*' | '/'

function encodeInstruction(opcode: Opcode, operand = 0): number {
  return (opcode << 24) | (operand & 0x00ff_ffff)
}

function columnToIndex(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    const code = column.charCodeAt(index)
    value = value * 26 + (code - 64)
  }
  return value - 1
}

function parseOperator(operator: string): SimpleDirectBinaryOperator | undefined {
  switch (operator) {
    case '+':
    case '-':
    case '*':
    case '/':
      return operator
    default:
      return undefined
  }
}

function operatorOpcode(operator: SimpleDirectBinaryOperator): Opcode {
  switch (operator) {
    case '+':
      return Opcode.Add
    case '-':
      return Opcode.Sub
    case '*':
      return Opcode.Mul
    case '/':
      return Opcode.Div
  }
}

function parsedCellRef(column: string, rowText: string): ParsedCellReferenceInfo {
  const normalizedColumn = column.toUpperCase()
  return {
    address: `${normalizedColumn}${rowText}`,
    row: Number.parseInt(rowText, 10) - 1,
    col: columnToIndex(normalizedColumn),
    rowAbsolute: false,
    colAbsolute: false,
  }
}

function cellNode(ref: ParsedCellReferenceInfo): FormulaNode {
  return {
    kind: 'CellRef',
    ref: ref.address,
  }
}

export function tryCompileSimpleDirectScalarFormula(source: string): CompiledFormula | undefined {
  const trimmed = source.trim()
  const match = SIMPLE_DIRECT_BINARY_RE.exec(trimmed)
  if (!match) {
    return undefined
  }

  const operator = parseOperator(match[3]!)
  if (operator === undefined) {
    return undefined
  }
  const opcode = operatorOpcode(operator)

  const leftRef = parsedCellRef(match[1]!, match[2]!)
  const rightRef = match[4] !== undefined ? parsedCellRef(match[4], match[5]!) : undefined
  const rightNumber = rightRef === undefined ? Number.parseFloat(match[6]!) : undefined
  if (rightRef === undefined && !Number.isFinite(rightNumber)) {
    return undefined
  }

  const rightNode: FormulaNode = rightRef ? cellNode(rightRef) : { kind: 'NumberLiteral', value: rightNumber! }
  const ast: FormulaNode = {
    kind: 'BinaryExpr',
    operator,
    left: cellNode(leftRef),
    right: rightNode,
  }

  const symbolicRefs = rightRef ? [leftRef.address, rightRef.address] : [leftRef.address]
  const parsedSymbolicRefs = rightRef ? [leftRef, rightRef] : [leftRef]
  const constants = rightRef ? new Float64Array() : Float64Array.of(rightNumber!)
  const program = Uint32Array.of(
    encodeInstruction(Opcode.PushCell, 0),
    rightRef ? encodeInstruction(Opcode.PushCell, 1) : encodeInstruction(Opcode.PushNumber, 0),
    encodeInstruction(opcode),
    encodeInstruction(Opcode.Ret),
  )
  const jsPlan: JsPlanInstruction[] = [
    { opcode: 'push-cell', address: leftRef.address },
    rightRef ? { opcode: 'push-cell', address: rightRef.address } : { opcode: 'push-number', value: rightNumber! },
    { opcode: 'binary', operator },
    { opcode: 'return' },
  ]
  const baseRecord: FormulaRecord = {
    id: 0,
    source: trimmed,
    mode: FormulaMode.WasmFastPath,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: program.length,
    constNumberOffset: 0,
    constNumberLength: constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    maxStackDepth: 2,
  }

  return {
    ...baseRecord,
    ast,
    optimizedAst: ast,
    astMatchesSource: true,
    deps: symbolicRefs,
    parsedDeps: parsedSymbolicRefs.map((ref) => ({ kind: 'cell', ...ref }) satisfies ParsedDependencyReference),
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan,
    program,
    constants,
    symbolicRefs,
    parsedSymbolicRefs,
    symbolicRanges: [],
    parsedSymbolicRanges: [],
    symbolicStrings: [],
  }
}
