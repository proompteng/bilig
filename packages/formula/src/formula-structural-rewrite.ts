import { ErrorCode } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import type { RangeAddress } from './addressing.js'
import { formatRangeAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import {
  compileFormulaAst,
  type CompiledFormula,
  type ParsedCellReferenceInfo,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
} from './compiler.js'
import type { JsPlanInstruction, ReferenceOperand } from './js-evaluator.js'
import { parseFormula } from './parser.js'
import { serializeFormula } from './formula-serializer.js'
import {
  formatAxisReference,
  formatCellReference,
  mapInterval,
  mapPointIndex,
  parseAxisReferenceParts,
  parseCellReferenceParts,
  quoteSheetNameIfNeeded,
  targetsSheet,
  type StructuralAxisTransform,
} from './translation-reference-utils.js'

export interface StructuralCompiledFormulaRewriteResult {
  source: string
  compiled: CompiledFormula
  reusedProgram: boolean
}

export function rewriteFormulaForStructuralTransform(
  source: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string {
  const ast = parseFormula(source)
  return serializeFormula(rewriteNodeForStructuralTransform(ast, ownerSheetName, targetSheetName, transform))
}

export function rewriteCompiledFormulaForStructuralTransform(
  compiled: CompiledFormula,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): StructuralCompiledFormulaRewriteResult {
  const currentAst = compiled.astMatchesSource === false ? parseFormula(compiled.source) : compiled.ast
  const currentOptimizedAst =
    compiled.astMatchesSource === false ? currentAst : compiled.optimizedAst === compiled.ast ? currentAst : compiled.optimizedAst
  const rewrittenAst = rewriteNodeForStructuralTransform(currentAst, ownerSheetName, targetSheetName, transform)
  const rewrittenOptimizedAst =
    currentOptimizedAst === currentAst
      ? rewrittenAst
      : rewriteNodeForStructuralTransform(currentOptimizedAst, ownerSheetName, targetSheetName, transform)
  const source = serializeFormula(rewrittenAst)

  if (!nodeStructuralShapeEqual(currentOptimizedAst, rewrittenOptimizedAst)) {
    return {
      source,
      compiled: compileFormulaAst(source, rewrittenOptimizedAst, {
        originalAst: rewrittenAst,
        symbolicNames: compiled.symbolicNames,
        symbolicTables: compiled.symbolicTables,
        symbolicSpills: compiled.symbolicSpills,
      }),
      reusedProgram: false,
    }
  }

  return {
    source,
    compiled: {
      ...compiled,
      source,
      ast: rewrittenAst,
      optimizedAst: rewrittenOptimizedAst,
      astMatchesSource: true,
      deps: compiled.deps.map((dependency) => rewriteQualifiedDependencyReference(dependency, ownerSheetName, targetSheetName, transform)),
      symbolicRefs: compiled.symbolicRefs.map((ref) => rewriteQualifiedCellReference(ref, ownerSheetName, targetSheetName, transform)),
      symbolicRanges: compiled.symbolicRanges.map((range) =>
        rewriteQualifiedRangeReference(range, ownerSheetName, targetSheetName, transform),
      ),
      jsPlan: compiled.jsPlan.map((instruction) => rewriteJsPlanInstruction(instruction, ownerSheetName, targetSheetName, transform)),
      ...(compiled.parsedDeps
        ? {
            parsedDeps: compiled.parsedDeps.map((dependency) =>
              rewriteParsedDependencyReference(dependency, ownerSheetName, targetSheetName, transform),
            ),
          }
        : {}),
      ...(compiled.parsedSymbolicRefs
        ? {
            parsedSymbolicRefs: compiled.parsedSymbolicRefs.map((ref) =>
              rewriteParsedCellReference(ref, ownerSheetName, targetSheetName, transform),
            ),
          }
        : {}),
      ...(compiled.parsedSymbolicRanges
        ? {
            parsedSymbolicRanges: compiled.parsedSymbolicRanges.map((range) =>
              rewriteParsedRangeReference(range, ownerSheetName, targetSheetName, transform),
            ),
          }
        : {}),
    },
    reusedProgram: true,
  }
}

export function rewriteAddressForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  const parsed = parseCellReferenceParts(address)
  if (!parsed) {
    throw new Error(`Invalid cell reference '${address}'`)
  }
  const nextRow = transform.axis === 'row' ? mapPointIndex(parsed.row, transform) : parsed.row
  const nextCol = transform.axis === 'column' ? mapPointIndex(parsed.col, transform) : parsed.col
  if (nextRow === undefined || nextCol === undefined) {
    return undefined
  }
  return formatCellReference(parsed, nextRow, nextCol)
}

export function rewriteRangeForStructuralTransform(
  startAddress: string,
  endAddress: string,
  transform: StructuralAxisTransform,
): { startAddress: string; endAddress: string } | undefined {
  const start = parseCellReferenceParts(startAddress)
  const end = parseCellReferenceParts(endAddress)
  if (!start || !end) {
    throw new Error(`Invalid range reference '${startAddress}:${endAddress}'`)
  }
  const nextRows =
    transform.axis === 'row'
      ? mapInterval(Math.min(start.row, end.row), Math.max(start.row, end.row), transform)
      : { start: Math.min(start.row, end.row), end: Math.max(start.row, end.row) }
  const nextCols =
    transform.axis === 'column'
      ? mapInterval(Math.min(start.col, end.col), Math.max(start.col, end.col), transform)
      : { start: Math.min(start.col, end.col), end: Math.max(start.col, end.col) }
  if (!nextRows || !nextCols) {
    return undefined
  }
  return {
    startAddress: formatCellReference(start, nextRows.start, nextCols.start),
    endAddress: formatCellReference(end, nextRows.end, nextCols.end),
  }
}

function rewriteNodeForStructuralTransform(
  node: FormulaNode,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
      return node
    case 'ArrayConstant':
      return {
        ...node,
        rows: node.rows.map((row) =>
          row.map((entry) => rewriteNodeForStructuralTransform(entry, ownerSheetName, targetSheetName, transform)),
        ),
      }
    case 'CellRef':
      return rewriteCellLikeNode(node, ownerSheetName, targetSheetName, transform)
    case 'SpillRef':
      return rewriteCellLikeNode(node, ownerSheetName, targetSheetName, transform)
    case 'ColumnRef':
      if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName) || transform.axis !== 'column') {
        return node
      }
      return rewriteAxisNode(node, transform)
    case 'RowRef':
      if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName) || transform.axis !== 'row') {
        return node
      }
      return rewriteAxisNode(node, transform)
    case 'RangeRef':
      return rewriteRangeNode(node, ownerSheetName, targetSheetName, transform)
    case 'UnaryExpr':
      return {
        ...node,
        argument: rewriteNodeForStructuralTransform(node.argument, ownerSheetName, targetSheetName, transform),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: rewriteNodeForStructuralTransform(node.left, ownerSheetName, targetSheetName, transform),
        right: rewriteNodeForStructuralTransform(node.right, ownerSheetName, targetSheetName, transform),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => rewriteNodeForStructuralTransform(arg, ownerSheetName, targetSheetName, transform)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: rewriteNodeForStructuralTransform(node.callee, ownerSheetName, targetSheetName, transform),
        args: node.args.map((arg) => rewriteNodeForStructuralTransform(arg, ownerSheetName, targetSheetName, transform)),
      }
  }
}

function nodeStructuralShapeEqual(left: FormulaNode, right: FormulaNode): boolean {
  if (left.kind !== right.kind) {
    return false
  }
  switch (left.kind) {
    case 'NumberLiteral':
      return right.kind === 'NumberLiteral' && left.value === right.value
    case 'BooleanLiteral':
      return right.kind === 'BooleanLiteral' && left.value === right.value
    case 'StringLiteral':
      return right.kind === 'StringLiteral' && left.value === right.value
    case 'ErrorLiteral':
      return right.kind === 'ErrorLiteral' && left.code === right.code
    case 'OmittedArgument':
      return true
    case 'ArrayConstant':
      return (
        right.kind === 'ArrayConstant' &&
        left.rows.length === right.rows.length &&
        left.rows.every(
          (row, rowIndex) =>
            row.length === right.rows[rowIndex]!.length &&
            row.every((entry, colIndex) => nodeStructuralShapeEqual(entry, right.rows[rowIndex]![colIndex]!)),
        )
      )
    case 'NameRef':
      return right.kind === 'NameRef' && left.name === right.name
    case 'StructuredRef':
      return right.kind === 'StructuredRef' && left.tableName === right.tableName && left.columnName === right.columnName
    case 'CellRef':
    case 'SpillRef':
    case 'ColumnRef':
    case 'RowRef':
      return true
    case 'RangeRef':
      return right.kind === 'RangeRef' && left.refKind === right.refKind
    case 'UnaryExpr':
      return right.kind === 'UnaryExpr' && left.operator === right.operator && nodeStructuralShapeEqual(left.argument, right.argument)
    case 'BinaryExpr':
      return (
        right.kind === 'BinaryExpr' &&
        left.operator === right.operator &&
        nodeStructuralShapeEqual(left.left, right.left) &&
        nodeStructuralShapeEqual(left.right, right.right)
      )
    case 'CallExpr':
      return (
        right.kind === 'CallExpr' &&
        left.callee === right.callee &&
        left.args.length === right.args.length &&
        left.args.every((arg, index) => nodeStructuralShapeEqual(arg, right.args[index]!))
      )
    case 'InvokeExpr':
      return (
        right.kind === 'InvokeExpr' &&
        left.args.length === right.args.length &&
        nodeStructuralShapeEqual(left.callee, right.callee) &&
        left.args.every((arg, index) => nodeStructuralShapeEqual(arg, right.args[index]!))
      )
  }
}

function rewriteCellLikeNode<T extends Extract<FormulaNode, { kind: 'CellRef' | 'SpillRef' }>>(
  node: T,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): FormulaNode {
  if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName)) {
    return node
  }
  const parsed = parseCellReferenceParts(node.ref)
  if (!parsed) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  const nextRow = transform.axis === 'row' ? mapPointIndex(parsed.row, transform) : parsed.row
  const nextCol = transform.axis === 'column' ? mapPointIndex(parsed.col, transform) : parsed.col
  if (nextRow === undefined || nextCol === undefined) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  return {
    ...node,
    ref: formatCellReference(parsed, nextRow, nextCol),
  }
}

function rewriteAxisNode<T extends Extract<FormulaNode, { kind: 'RowRef' | 'ColumnRef' }>>(
  node: T,
  transform: StructuralAxisTransform,
): FormulaNode {
  const parsed = parseAxisReferenceParts(node.ref, node.kind === 'RowRef' ? 'row' : 'column')
  if (!parsed) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  const nextIndex = mapPointIndex(parsed.index, transform)
  if (nextIndex === undefined) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  return {
    ...node,
    ref: formatAxisReference(parsed.absolute, nextIndex, node.kind === 'RowRef' ? 'row' : 'column'),
  }
}

function rewriteRangeNode(
  node: Extract<FormulaNode, { kind: 'RangeRef' }>,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): FormulaNode {
  if (!targetsSheet(node.sheetName, ownerSheetName, targetSheetName)) {
    return node
  }
  if ((node.refKind === 'rows' && transform.axis === 'column') || (node.refKind === 'cols' && transform.axis === 'row')) {
    return node
  }
  if (node.refKind === 'cells') {
    const start = parseCellReferenceParts(node.start)
    const end = parseCellReferenceParts(node.end)
    if (!start || !end) {
      return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
    }
    const nextRows =
      transform.axis === 'row'
        ? mapInterval(Math.min(start.row, end.row), Math.max(start.row, end.row), transform)
        : { start: Math.min(start.row, end.row), end: Math.max(start.row, end.row) }
    const nextCols =
      transform.axis === 'column'
        ? mapInterval(Math.min(start.col, end.col), Math.max(start.col, end.col), transform)
        : { start: Math.min(start.col, end.col), end: Math.max(start.col, end.col) }
    if (!nextRows || !nextCols) {
      return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
    }
    return {
      ...node,
      start: formatCellReference(start, nextRows.start, nextCols.start),
      end: formatCellReference(end, nextRows.end, nextCols.end),
    }
  }
  const start = parseAxisReferenceParts(node.start, node.refKind === 'rows' ? 'row' : 'column')
  const end = parseAxisReferenceParts(node.end, node.refKind === 'rows' ? 'row' : 'column')
  if (!start || !end) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  const nextInterval = mapInterval(Math.min(start.index, end.index), Math.max(start.index, end.index), transform)
  if (!nextInterval) {
    return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
  }
  return {
    ...node,
    start: formatAxisReference(start.absolute, nextInterval.start, node.refKind === 'rows' ? 'row' : 'column'),
    end: formatAxisReference(end.absolute, nextInterval.end, node.refKind === 'rows' ? 'row' : 'column'),
  }
}

function rewriteParsedCellReference<Reference extends ParsedCellReferenceInfo>(
  reference: Reference,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): Reference {
  const sheetName = reference.sheetName ?? ownerSheetName
  if (sheetName !== targetSheetName) {
    return reference
  }
  const nextAddress = rewriteAddressForStructuralTransform(reference.address, transform)
  if (!nextAddress) {
    return reference
  }
  const parsed = parseCellAddress(nextAddress, sheetName)
  return {
    ...reference,
    address: parsed.text,
    ...(reference.sheetName !== undefined ? { sheetName: parsed.sheetName } : {}),
    ...(reference.row !== undefined ? { row: parsed.row } : {}),
    ...(reference.col !== undefined ? { col: parsed.col } : {}),
  }
}

function rewriteParsedRangeReference(
  reference: ParsedRangeReferenceInfo,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): ParsedRangeReferenceInfo {
  const explicitSheetName = reference.sheetName
  if (!targetsSheet(explicitSheetName, ownerSheetName, targetSheetName)) {
    return reference
  }
  const nextRange = rewriteRangeAddressForStructuralTransform(
    parseRangeAddress(formatQualifiedRangeReference(explicitSheetName, reference.startAddress, reference.endAddress)),
    transform,
  )
  if (!nextRange) {
    return reference
  }
  const bounds =
    nextRange.kind === 'cells'
      ? {
          startRow: nextRange.start.row,
          endRow: nextRange.end.row,
          startCol: nextRange.start.col,
          endCol: nextRange.end.col,
        }
      : nextRange.kind === 'rows'
        ? {
            startRow: nextRange.start.row,
            endRow: nextRange.end.row,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: nextRange.start.col,
            endCol: nextRange.end.col,
          }
  return {
    ...reference,
    address: formatQualifiedRangeReference(explicitSheetName, nextRange.start.text, nextRange.end.text),
    refKind: nextRange.kind,
    startAddress: nextRange.start.text,
    endAddress: nextRange.end.text,
    ...bounds,
  }
}

function rewriteParsedDependencyReference(
  reference: ParsedDependencyReference,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): ParsedDependencyReference {
  return reference.kind === 'cell'
    ? rewriteParsedCellReference(reference, ownerSheetName, targetSheetName, transform)
    : rewriteParsedRangeReference(reference, ownerSheetName, targetSheetName, transform)
}

function rewriteQualifiedCellReference(
  raw: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string {
  const explicitlyQualified = raw.includes('!')
  const parsed = parseCellAddress(raw, ownerSheetName)
  if (parsed.sheetName !== targetSheetName) {
    return raw
  }
  const nextAddress = rewriteAddressForStructuralTransform(parsed.text, transform)
  if (!nextAddress) {
    return raw
  }
  return explicitlyQualified ? formatQualifiedCellReference(parsed.sheetName, nextAddress) : nextAddress
}

function rewriteQualifiedDependencyReference(
  raw: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string {
  if (!raw.includes(':')) {
    return rewriteQualifiedCellReference(raw, ownerSheetName, targetSheetName, transform)
  }
  return rewriteQualifiedRangeReference(raw, ownerSheetName, targetSheetName, transform)
}

function rewriteQualifiedRangeReference(
  raw: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string {
  const explicitlyQualified = raw.includes('!')
  const parsed = parseRangeAddress(raw, ownerSheetName)
  const sheetName = parsed.sheetName ?? ownerSheetName
  if (sheetName !== targetSheetName) {
    return raw
  }
  const nextRange = rewriteRangeAddressForStructuralTransform(parsed, transform)
  if (!nextRange) {
    return raw
  }
  if (explicitlyQualified) {
    return formatRangeAddress(nextRange)
  }
  return `${nextRange.start.text}:${nextRange.end.text}`
}

function rewriteRangeAddressForStructuralTransform(range: RangeAddress, transform: StructuralAxisTransform): RangeAddress | undefined {
  switch (range.kind) {
    case 'cells': {
      const nextRange = rewriteRangeForStructuralTransform(range.start.text, range.end.text, transform)
      if (!nextRange) {
        return undefined
      }
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, nextRange.startAddress, nextRange.endAddress))
    }
    case 'rows':
      if (transform.axis !== 'row') {
        return range
      }
      return rewriteAxisRangeAddress(range, transform)
    case 'cols':
      if (transform.axis !== 'column') {
        return range
      }
      return rewriteAxisRangeAddress(range, transform)
  }
}

function rewriteAxisRangeAddress(
  range: Extract<RangeAddress, { kind: 'rows' | 'cols' }>,
  transform: StructuralAxisTransform,
): RangeAddress | undefined {
  const startIndex = range.kind === 'rows' ? range.start.row : range.start.col
  const endIndex = range.kind === 'rows' ? range.end.row : range.end.col
  const nextInterval = mapInterval(startIndex, endIndex, transform)
  if (!nextInterval) {
    return undefined
  }
  const prefix = range.sheetName ? `${quoteSheetNameIfNeeded(range.sheetName)}!` : ''
  const startText =
    range.kind === 'rows' ? formatAxisReference(false, nextInterval.start, 'row') : formatAxisReference(false, nextInterval.start, 'column')
  const endText =
    range.kind === 'rows' ? formatAxisReference(false, nextInterval.end, 'row') : formatAxisReference(false, nextInterval.end, 'column')
  return parseRangeAddress(`${prefix}${startText}:${endText}`)
}

function rewriteJsPlanInstruction(
  instruction: JsPlanInstruction,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): JsPlanInstruction {
  switch (instruction.opcode) {
    case 'push-cell':
      return {
        ...instruction,
        address: rewriteReferenceOperandAddress(instruction.sheetName, instruction.address, ownerSheetName, targetSheetName, transform),
      }
    case 'push-range': {
      const nextRange = rewritePlanRangeInstruction(
        instruction.sheetName,
        instruction.start,
        instruction.end,
        instruction.refKind,
        ownerSheetName,
        targetSheetName,
        transform,
      )
      return nextRange ? { ...instruction, ...nextRange } : instruction
    }
    case 'lookup-exact-match':
    case 'lookup-approximate-match': {
      const nextRange = rewritePlanRangeInstruction(
        instruction.sheetName,
        instruction.start,
        instruction.end,
        instruction.refKind,
        ownerSheetName,
        targetSheetName,
        transform,
      )
      if (!nextRange) {
        return instruction
      }
      const parsed = parseRangeAddress(formatQualifiedRangeReference(instruction.sheetName, nextRange.start, nextRange.end))
      if (parsed.kind !== 'cells') {
        return instruction
      }
      return {
        ...instruction,
        ...nextRange,
        startRow: parsed.start.row,
        endRow: parsed.end.row,
        startCol: parsed.start.col,
        endCol: parsed.end.col,
      }
    }
    case 'call':
      return instruction.argRefs
        ? {
            ...instruction,
            argRefs: instruction.argRefs.map((argRef) =>
              argRef ? rewriteReferenceOperand(argRef, ownerSheetName, targetSheetName, transform) : argRef,
            ),
          }
        : instruction
    case 'push-lambda':
      return {
        ...instruction,
        body: instruction.body.map((step) => rewriteJsPlanInstruction(step, ownerSheetName, targetSheetName, transform)),
      }
    case 'push-number':
    case 'push-boolean':
    case 'push-string':
    case 'push-error':
    case 'push-name':
    case 'push-omitted':
    case 'make-array':
    case 'unary':
    case 'binary':
    case 'invoke':
    case 'begin-scope':
    case 'bind-name':
    case 'end-scope':
    case 'jump-if-false':
    case 'jump':
    case 'return':
      return instruction
  }
}

function rewriteReferenceOperand(
  operand: ReferenceOperand,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): ReferenceOperand {
  switch (operand.kind) {
    case 'cell':
      return operand.address
        ? {
            ...operand,
            address: rewriteReferenceOperandAddress(operand.sheetName, operand.address, ownerSheetName, targetSheetName, transform),
          }
        : operand
    case 'range': {
      if (!operand.start || !operand.end || !operand.refKind) {
        return operand
      }
      const nextRange = rewritePlanRangeInstruction(
        operand.sheetName,
        operand.start,
        operand.end,
        operand.refKind,
        ownerSheetName,
        targetSheetName,
        transform,
      )
      return nextRange ? { ...operand, ...nextRange } : operand
    }
    case 'row':
    case 'col':
      return operand
  }
}

function rewriteReferenceOperandAddress(
  explicitSheetName: string | undefined,
  address: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string {
  if ((explicitSheetName ?? ownerSheetName) !== targetSheetName) {
    return address
  }
  return rewriteAddressForStructuralTransform(address, transform) ?? address
}

function rewritePlanRangeInstruction(
  explicitSheetName: string | undefined,
  start: string,
  end: string,
  refKind: 'cells' | 'rows' | 'cols',
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): { start: string; end: string } | undefined {
  if ((explicitSheetName ?? ownerSheetName) !== targetSheetName) {
    return undefined
  }
  if (refKind === 'cells') {
    const nextRange = rewriteRangeForStructuralTransform(start, end, transform)
    return nextRange
      ? {
          start: nextRange.startAddress,
          end: nextRange.endAddress,
        }
      : undefined
  }
  if ((refKind === 'rows' && transform.axis !== 'row') || (refKind === 'cols' && transform.axis !== 'column')) {
    return undefined
  }
  const parsed = parseRangeAddress(formatQualifiedRangeReference(explicitSheetName, start, end))
  const nextRange = rewriteRangeAddressForStructuralTransform(parsed, transform)
  return nextRange
    ? {
        start: nextRange.start.text,
        end: nextRange.end.text,
      }
    : undefined
}

function formatQualifiedCellReference(sheetName: string | undefined, address: string): string {
  if (!sheetName) {
    return address
  }
  const parsed = parseCellAddress(address, sheetName)
  return `${quoteSheetNameIfNeeded(sheetName)}!${parsed.text}`
}

function formatQualifiedRangeReference(sheetName: string | undefined, start: string, end: string): string {
  const prefix = sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!` : ''
  return `${prefix}${start}:${end}`
}
