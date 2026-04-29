import { BuiltinId, FormulaMode, Opcode, type FormulaRecord } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode, JsPlanInstruction, ParsedCellReferenceInfo, ParsedDependencyReference } from '@bilig/formula'

const SIMPLE_DIRECT_BINARY_RE = /^([A-Za-z]+)([1-9][0-9]*)([+\-*/])(?:([A-Za-z]+)([1-9][0-9]*)|(\d+(?:\.\d+)?))$/
const SIMPLE_DIRECT_ABS_RE = /^ABS\s*\(\s*([A-Za-z]+)([1-9][0-9]*)\s*\)$/i
const EMPTY_STRINGS: string[] = []
const EMPTY_CONSTANTS = new Float64Array()

type SimpleDirectBinaryOperator = '+' | '-' | '*' | '/'
type SimpleDirectScalarAst = FormulaNode & {
  readonly kind: 'BinaryExpr'
  readonly operator: SimpleDirectBinaryOperator
  readonly left: FormulaNode & { readonly kind: 'CellRef' }
  readonly right: FormulaNode & ({ readonly kind: 'CellRef' } | { readonly kind: 'NumberLiteral' })
}
type SimpleDirectAbsAst = FormulaNode & {
  readonly kind: 'CallExpr'
  readonly callee: 'ABS'
  readonly args: readonly [FormulaNode & { readonly kind: 'CellRef' }]
}

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

function indexToColumn(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) {
    value -= 1
    label = String.fromCharCode(65 + (value % 26)) + label
    value = Math.floor(value / 26)
  }
  return label
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

function translateParsedLocalCellRef(
  ref: ParsedCellReferenceInfo,
  rowDelta: number,
  colDelta: number,
): ParsedCellReferenceInfo | undefined {
  if (ref.sheetName !== undefined || ref.explicitSheet === true) {
    return undefined
  }
  const baseRow = ref.row
  const baseCol = ref.col
  if (baseRow === undefined || baseCol === undefined) {
    return undefined
  }
  const row = ref.rowAbsolute ? baseRow : baseRow + rowDelta
  const col = ref.colAbsolute ? baseCol : baseCol + colDelta
  if (row < 0 || col < 0) {
    return undefined
  }
  return {
    address: `${indexToColumn(col)}${row + 1}`,
    row,
    col,
    ...(ref.rowAbsolute !== undefined ? { rowAbsolute: ref.rowAbsolute } : {}),
    ...(ref.colAbsolute !== undefined ? { colAbsolute: ref.colAbsolute } : {}),
  }
}

function isSimpleDirectScalarAst(node: FormulaNode): node is SimpleDirectScalarAst {
  return (
    node.kind === 'BinaryExpr' &&
    (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/') &&
    node.left.kind === 'CellRef' &&
    (node.right.kind === 'CellRef' || node.right.kind === 'NumberLiteral')
  )
}

function isSimpleDirectAbsAst(node: FormulaNode): node is SimpleDirectAbsAst {
  return node.kind === 'CallExpr' && node.callee === 'ABS' && node.args.length === 1 && node.args[0]?.kind === 'CellRef'
}

export function translateSimpleDirectScalarFormula(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  source: string,
): CompiledFormula | undefined {
  if (
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0 ||
    (!isSimpleDirectScalarAst(compiled.optimizedAst) && !isSimpleDirectAbsAst(compiled.optimizedAst))
  ) {
    return undefined
  }
  const expectedRefCount = compiled.optimizedAst.kind === 'CallExpr' || compiled.optimizedAst.right.kind !== 'CellRef' ? 1 : 2
  if (compiled.parsedSymbolicRefs === undefined || compiled.parsedSymbolicRefs.length !== expectedRefCount) {
    return undefined
  }
  const translatedRefs: ParsedCellReferenceInfo[] = []
  for (let index = 0; index < compiled.parsedSymbolicRefs.length; index += 1) {
    const translated = translateParsedLocalCellRef(compiled.parsedSymbolicRefs[index]!, rowDelta, colDelta)
    if (!translated) {
      return undefined
    }
    translatedRefs.push(translated)
  }
  const symbolicRefs = translatedRefs.map((ref) => ref.address)
  return {
    ...compiled,
    source,
    astMatchesSource: false,
    deps: symbolicRefs,
    parsedDeps: translatedRefs.map((ref) => ({ kind: 'cell', ...ref }) satisfies ParsedDependencyReference),
    symbolicRefs,
    parsedSymbolicRefs: translatedRefs,
  }
}

export function tryCompileSimpleDirectScalarFormula(source: string): CompiledFormula | undefined {
  const trimmedSource = source.trim()
  const trimmed = trimmedSource.startsWith('=') ? trimmedSource.slice(1).trim() : trimmedSource
  const absMatch = SIMPLE_DIRECT_ABS_RE.exec(trimmed)
  if (absMatch) {
    const ref = parsedCellRef(absMatch[1]!, absMatch[2]!)
    const ast: FormulaNode = {
      kind: 'CallExpr',
      callee: 'ABS',
      args: [cellNode(ref)],
    }
    const program = Uint32Array.of(
      encodeInstruction(Opcode.PushCell, 0),
      encodeInstruction(Opcode.CallBuiltin, (BuiltinId.Abs << 8) | 1),
      encodeInstruction(Opcode.Ret),
    )
    const jsPlan: JsPlanInstruction[] = [
      { opcode: 'push-cell', address: ref.address },
      { opcode: 'call', callee: 'ABS', argc: 1 },
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
      constNumberLength: 0,
      rangeListOffset: 0,
      rangeListLength: 0,
      maxStackDepth: 1,
    }

    return {
      ...baseRecord,
      ast,
      optimizedAst: ast,
      astMatchesSource: true,
      deps: [ref.address],
      parsedDeps: [{ kind: 'cell', ...ref } satisfies ParsedDependencyReference],
      symbolicNames: EMPTY_STRINGS,
      symbolicTables: EMPTY_STRINGS,
      symbolicSpills: EMPTY_STRINGS,
      volatile: false,
      randCallCount: 0,
      producesSpill: false,
      jsPlan,
      program,
      constants: EMPTY_CONSTANTS,
      symbolicRefs: [ref.address],
      parsedSymbolicRefs: [ref],
      symbolicRanges: EMPTY_STRINGS,
      parsedSymbolicRanges: [],
      symbolicStrings: EMPTY_STRINGS,
    }
  }
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
  const constants = rightRef ? EMPTY_CONSTANTS : Float64Array.of(rightNumber!)
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
    symbolicNames: EMPTY_STRINGS,
    symbolicTables: EMPTY_STRINGS,
    symbolicSpills: EMPTY_STRINGS,
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan,
    program,
    constants,
    symbolicRefs,
    parsedSymbolicRefs,
    symbolicRanges: EMPTY_STRINGS,
    parsedSymbolicRanges: [],
    symbolicStrings: EMPTY_STRINGS,
  }
}
