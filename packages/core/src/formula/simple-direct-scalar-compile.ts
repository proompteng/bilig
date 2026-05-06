import { BuiltinId, FormulaMode, Opcode, type FormulaRecord } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode, JsPlanInstruction, ParsedCellReferenceInfo, ParsedDependencyReference } from '@bilig/formula'

const SIMPLE_DIRECT_BINARY_RE = /^([A-Za-z]+)([1-9][0-9]*)([+\-*/])(?:([A-Za-z]+)([1-9][0-9]*)|(\d+(?:\.\d+)?))(?:\+(\d+(?:\.\d+)?))?$/
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
type SimpleDirectScalarOffsetAst = FormulaNode & {
  readonly kind: 'BinaryExpr'
  readonly operator: '+'
  readonly left: SimpleDirectScalarAst
  readonly right: FormulaNode & { readonly kind: 'NumberLiteral' }
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
  const translated: ParsedCellReferenceInfo = {
    address: `${indexToColumn(col)}${row + 1}`,
    row,
    col,
  }
  if (ref.rowAbsolute !== undefined) {
    translated.rowAbsolute = ref.rowAbsolute
  }
  if (ref.colAbsolute !== undefined) {
    translated.colAbsolute = ref.colAbsolute
  }
  return translated
}

function parsedCellDependency(ref: ParsedCellReferenceInfo): ParsedDependencyReference {
  const dependency: ParsedDependencyReference = {
    kind: 'cell',
    address: ref.address,
  }
  if (ref.sheetName !== undefined) {
    dependency.sheetName = ref.sheetName
  }
  if (ref.explicitSheet !== undefined) {
    dependency.explicitSheet = ref.explicitSheet
  }
  if (ref.row !== undefined) {
    dependency.row = ref.row
  }
  if (ref.col !== undefined) {
    dependency.col = ref.col
  }
  if (ref.rowAbsolute !== undefined) {
    dependency.rowAbsolute = ref.rowAbsolute
  }
  if (ref.colAbsolute !== undefined) {
    dependency.colAbsolute = ref.colAbsolute
  }
  return dependency
}

function parsedTranslatedSourceRef(
  column: string,
  rowText: string,
  baseRef: ParsedCellReferenceInfo,
  rowDelta: number,
  colDelta: number,
): ParsedCellReferenceInfo | undefined {
  if (baseRef.sheetName !== undefined || baseRef.explicitSheet === true) {
    return undefined
  }
  const baseRow = baseRef.row
  const baseCol = baseRef.col
  if (baseRow === undefined || baseCol === undefined) {
    return undefined
  }
  const row = baseRef.rowAbsolute ? baseRow : baseRow + rowDelta
  const col = baseRef.colAbsolute ? baseCol : baseCol + colDelta
  const parsed = parsedCellRef(column, rowText)
  if (parsed.row !== row || parsed.col !== col) {
    return undefined
  }
  if (baseRef.rowAbsolute !== undefined) {
    parsed.rowAbsolute = baseRef.rowAbsolute
  }
  if (baseRef.colAbsolute !== undefined) {
    parsed.colAbsolute = baseRef.colAbsolute
  }
  return parsed
}

function translatedRefsFromSource(
  compiled: CompiledFormula,
  scalarAst: SimpleDirectScalarAst | SimpleDirectAbsAst,
  rowDelta: number,
  colDelta: number,
  source: string,
): ParsedCellReferenceInfo[] | undefined {
  const parsedRefs = compiled.parsedSymbolicRefs
  if (parsedRefs === undefined) {
    return undefined
  }
  const trimmedSource = source.trim()
  const trimmed = trimmedSource.startsWith('=') ? trimmedSource.slice(1).trim() : trimmedSource
  if (isSimpleDirectAbsAst(scalarAst)) {
    const absMatch = SIMPLE_DIRECT_ABS_RE.exec(trimmed)
    if (!absMatch || parsedRefs.length !== 1) {
      return undefined
    }
    const operand = parsedTranslatedSourceRef(absMatch[1]!, absMatch[2]!, parsedRefs[0]!, rowDelta, colDelta)
    return operand ? [operand] : undefined
  }

  const binaryMatch = SIMPLE_DIRECT_BINARY_RE.exec(trimmed)
  if (!binaryMatch) {
    return undefined
  }
  const left = parsedTranslatedSourceRef(binaryMatch[1]!, binaryMatch[2]!, parsedRefs[0]!, rowDelta, colDelta)
  if (!left) {
    return undefined
  }
  if (scalarAst.right.kind !== 'CellRef') {
    return parsedRefs.length === 1 && binaryMatch[4] === undefined ? [left] : undefined
  }
  if (parsedRefs.length !== 2 || binaryMatch[4] === undefined) {
    return undefined
  }
  const right = parsedTranslatedSourceRef(binaryMatch[4], binaryMatch[5]!, parsedRefs[1]!, rowDelta, colDelta)
  return right ? [left, right] : undefined
}

function translatedCompiledFormula(
  compiled: CompiledFormula,
  source: string,
  deps: string[],
  parsedDeps: ParsedDependencyReference[],
  parsedSymbolicRefs: ParsedCellReferenceInfo[],
): CompiledFormula {
  const translated: CompiledFormula = {
    id: compiled.id,
    source,
    mode: compiled.mode,
    depsPtr: compiled.depsPtr,
    depsLen: compiled.depsLen,
    programOffset: compiled.programOffset,
    programLength: compiled.programLength,
    constNumberOffset: compiled.constNumberOffset,
    constNumberLength: compiled.constNumberLength,
    rangeListOffset: compiled.rangeListOffset,
    rangeListLength: compiled.rangeListLength,
    maxStackDepth: compiled.maxStackDepth,
    ast: compiled.ast,
    optimizedAst: compiled.optimizedAst,
    astMatchesSource: false,
    deps,
    parsedDeps,
    symbolicNames: compiled.symbolicNames,
    symbolicTables: compiled.symbolicTables,
    symbolicSpills: compiled.symbolicSpills,
    volatile: compiled.volatile,
    randCallCount: compiled.randCallCount,
    producesSpill: compiled.producesSpill,
    jsPlan: compiled.jsPlan,
    program: compiled.program,
    constants: compiled.constants,
    symbolicRefs: deps,
    parsedSymbolicRefs,
    symbolicRanges: compiled.symbolicRanges,
    symbolicStrings: compiled.symbolicStrings,
  }
  if (compiled.parsedSymbolicRanges !== undefined) {
    translated.parsedSymbolicRanges = compiled.parsedSymbolicRanges
  }
  if (compiled.directAggregateCandidate !== undefined) {
    translated.directAggregateCandidate = compiled.directAggregateCandidate
  }
  return translated
}

function isSimpleDirectScalarAst(node: FormulaNode): node is SimpleDirectScalarAst {
  return (
    node.kind === 'BinaryExpr' &&
    (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/') &&
    node.left.kind === 'CellRef' &&
    (node.right.kind === 'CellRef' || node.right.kind === 'NumberLiteral')
  )
}

function isSimpleDirectScalarOffsetAst(node: FormulaNode): node is SimpleDirectScalarOffsetAst {
  return node.kind === 'BinaryExpr' && node.operator === '+' && isSimpleDirectScalarAst(node.left) && node.right.kind === 'NumberLiteral'
}

function isSimpleDirectAbsAst(node: FormulaNode): node is SimpleDirectAbsAst {
  return node.kind === 'CallExpr' && node.callee === 'ABS' && node.args.length === 1 && node.args[0]?.kind === 'CellRef'
}

export function translateSimpleDirectScalarFormulaWithParsedRefs(
  compiled: CompiledFormula,
  source: string,
  parsedSymbolicRefs: ParsedCellReferenceInfo[],
): CompiledFormula | undefined {
  if (
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0 ||
    (!isSimpleDirectScalarAst(compiled.optimizedAst) &&
      !isSimpleDirectScalarOffsetAst(compiled.optimizedAst) &&
      !isSimpleDirectAbsAst(compiled.optimizedAst))
  ) {
    return undefined
  }
  const scalarAst = isSimpleDirectScalarOffsetAst(compiled.optimizedAst) ? compiled.optimizedAst.left : compiled.optimizedAst
  const expectedRefCount = isSimpleDirectScalarAst(scalarAst) && scalarAst.right.kind === 'CellRef' ? 2 : 1
  if (parsedSymbolicRefs.length !== expectedRefCount) {
    return undefined
  }
  const symbolicRefs = parsedSymbolicRefs.map((ref) => ref.address)
  const parsedDeps = parsedSymbolicRefs.map(parsedCellDependency)
  return translatedCompiledFormula(compiled, source, symbolicRefs, parsedDeps, parsedSymbolicRefs)
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
    (!isSimpleDirectScalarAst(compiled.optimizedAst) &&
      !isSimpleDirectScalarOffsetAst(compiled.optimizedAst) &&
      !isSimpleDirectAbsAst(compiled.optimizedAst))
  ) {
    return undefined
  }
  const scalarAst = isSimpleDirectScalarOffsetAst(compiled.optimizedAst) ? compiled.optimizedAst.left : compiled.optimizedAst
  const expectedRefCount = isSimpleDirectScalarAst(scalarAst) && scalarAst.right.kind === 'CellRef' ? 2 : 1
  if (compiled.parsedSymbolicRefs === undefined || compiled.parsedSymbolicRefs.length !== expectedRefCount) {
    return undefined
  }
  const sourceRefs = translatedRefsFromSource(compiled, scalarAst, rowDelta, colDelta, source)
  if (sourceRefs) {
    const sourceSymbolicRefs = sourceRefs.map((ref) => ref.address)
    const sourceParsedDeps = sourceRefs.map(parsedCellDependency)
    return translatedCompiledFormula(compiled, source, sourceSymbolicRefs, sourceParsedDeps, sourceRefs)
  }
  const translatedRefs: ParsedCellReferenceInfo[] = []
  const symbolicRefs: ParsedCellReferenceInfo['address'][] = []
  const parsedDeps: ParsedDependencyReference[] = []
  for (let index = 0; index < compiled.parsedSymbolicRefs.length; index += 1) {
    const translated = translateParsedLocalCellRef(compiled.parsedSymbolicRefs[index]!, rowDelta, colDelta)
    if (!translated) {
      return undefined
    }
    translatedRefs[index] = translated
    symbolicRefs[index] = translated.address
    parsedDeps[index] = parsedCellDependency(translated)
  }
  return translatedCompiledFormula(compiled, source, symbolicRefs, parsedDeps, translatedRefs)
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
  const resultOffset = match[7] === undefined ? undefined : Number.parseFloat(match[7])
  if (resultOffset !== undefined && !Number.isFinite(resultOffset)) {
    return undefined
  }

  const rightNode: FormulaNode = rightRef ? cellNode(rightRef) : { kind: 'NumberLiteral', value: rightNumber! }
  const baseAst: FormulaNode = {
    kind: 'BinaryExpr',
    operator,
    left: cellNode(leftRef),
    right: rightNode,
  }
  const ast: FormulaNode =
    resultOffset === undefined
      ? baseAst
      : {
          kind: 'BinaryExpr',
          operator: '+',
          left: baseAst,
          right: { kind: 'NumberLiteral', value: resultOffset },
        }

  const symbolicRefs = rightRef ? [leftRef.address, rightRef.address] : [leftRef.address]
  const parsedSymbolicRefs = rightRef ? [leftRef, rightRef] : [leftRef]
  const constants = rightRef
    ? resultOffset === undefined
      ? EMPTY_CONSTANTS
      : Float64Array.of(resultOffset)
    : resultOffset === undefined
      ? Float64Array.of(rightNumber!)
      : Float64Array.of(rightNumber!, resultOffset)
  const programInstructions = [
    encodeInstruction(Opcode.PushCell, 0),
    rightRef ? encodeInstruction(Opcode.PushCell, 1) : encodeInstruction(Opcode.PushNumber, 0),
    encodeInstruction(opcode),
  ]
  if (resultOffset !== undefined) {
    programInstructions.push(encodeInstruction(Opcode.PushNumber, rightRef ? 0 : 1), encodeInstruction(Opcode.Add))
  }
  programInstructions.push(encodeInstruction(Opcode.Ret))
  const program = Uint32Array.from(programInstructions)
  const jsPlan: JsPlanInstruction[] = [
    { opcode: 'push-cell', address: leftRef.address },
    rightRef ? { opcode: 'push-cell', address: rightRef.address } : { opcode: 'push-number', value: rightNumber! },
    { opcode: 'binary', operator },
  ]
  if (resultOffset !== undefined) {
    jsPlan.push({ opcode: 'push-number', value: resultOffset }, { opcode: 'binary', operator: '+' })
  }
  jsPlan.push({ opcode: 'return' })
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
