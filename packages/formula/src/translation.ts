import { ErrorCode, FormulaMode } from '@bilig/protocol'
import type { BinaryExprNode, FormulaNode } from './ast.js'
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

const CELL_REF_RE = /^(\$?)([A-Z]+)(\$?)([1-9][0-9]*)$/
const COLUMN_REF_RE = /^(\$?)([A-Z]+)$/
const ROW_REF_RE = /^(\$?)([1-9][0-9]*)$/

const BINARY_PRECEDENCE: Record<BinaryExprNode['operator'], number> = {
  '=': 1,
  '<>': 1,
  '>': 1,
  '>=': 1,
  '<': 1,
  '<=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
}

export type StructuralAxisKind = 'row' | 'column'

export type StructuralAxisTransform =
  | { kind: 'insert'; axis: StructuralAxisKind; start: number; count: number }
  | { kind: 'delete'; axis: StructuralAxisKind; start: number; count: number }
  | { kind: 'move'; axis: StructuralAxisKind; start: number; count: number; target: number }

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`)
}

const ERROR_LITERAL_TEXT: Record<number, string> = {
  [ErrorCode.Ref]: '#REF!',
  [ErrorCode.Name]: '#NAME?',
  [ErrorCode.Div0]: '#DIV/0!',
  [ErrorCode.NA]: '#N/A',
  [ErrorCode.Value]: '#VALUE!',
  [ErrorCode.Cycle]: '#CYCLE!',
  [ErrorCode.Spill]: '#SPILL!',
  [ErrorCode.Blocked]: '#BLOCKED!',
}

export function translateFormulaReferences(source: string, rowDelta: number, colDelta: number): string {
  const ast = parseFormula(source)
  return serializeFormula(translateNode(ast, rowDelta, colDelta))
}

export function buildRelativeFormulaTemplateKey(source: string, ownerRow: number, ownerCol: number): string {
  return buildRelativeFormulaTemplateKeyFromAst(parseFormula(source), ownerRow, ownerCol)
}

export function buildRelativeFormulaTemplateKeyFromAst(node: FormulaNode, ownerRow: number, ownerCol: number): string {
  return buildRelativeFormulaTemplateKeyInternal(node, ownerRow, ownerCol)
}

export interface CompiledFormulaTranslationResult {
  source: string
  compiled: CompiledFormula
}

export function canTranslateCompiledFormulaWithoutAst(compiled: CompiledFormula): boolean {
  return (
    (compiled.symbolicRanges.length === 0 || compiled.directAggregateCandidate !== undefined) &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    !compiled.jsPlan.some((instruction) => instruction.opcode === 'lookup-exact-match' || instruction.opcode === 'lookup-approximate-match')
  )
}

export function translateCompiledFormulaWithoutAst(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  sourceOverride?: string,
): CompiledFormulaTranslationResult {
  const translatedParsedDeps = compiled.parsedDeps?.map((dependency) => translateParsedDependencyReference(dependency, rowDelta, colDelta))
  const translatedParsedSymbolicRefs = compiled.parsedSymbolicRefs?.map((reference) =>
    translateParsedCellReference(reference, rowDelta, colDelta),
  )
  const translatedParsedSymbolicRanges = compiled.parsedSymbolicRanges?.map((range) =>
    translateParsedRangeReference(range, rowDelta, colDelta),
  )
  const source = sourceOverride ?? compiled.source
  const translatedCellMap = buildTranslatedCellReferenceMap(compiled.parsedSymbolicRefs, translatedParsedSymbolicRefs)
  const translatedRangeMap = buildTranslatedRangeReferenceMap(compiled.parsedSymbolicRanges, translatedParsedSymbolicRanges)

  return {
    source,
    compiled: {
      ...compiled,
      source,
      astMatchesSource: false,
      deps:
        translatedParsedDeps?.map((dependency) => formatParsedDependencyReference(dependency)) ??
        compiled.deps.map((dependency) => translateQualifiedDependencyReference(dependency, rowDelta, colDelta)),
      symbolicRefs:
        translatedParsedSymbolicRefs?.map((reference) => formatParsedCellReference(reference)) ??
        compiled.symbolicRefs.map((ref) => translateQualifiedCellReference(ref, rowDelta, colDelta)),
      symbolicRanges:
        translatedParsedSymbolicRanges?.map((range) => formatParsedRangeReference(range)) ??
        compiled.symbolicRanges.map((range) => translateQualifiedRangeReference(range, rowDelta, colDelta)),
      jsPlan:
        compiled.symbolicRanges.length === 0 && compiled.mode === FormulaMode.WasmFastPath
          ? compiled.jsPlan
          : compiled.jsPlan.map((instruction) =>
              translateJsPlanInstructionWithoutAst(instruction, translatedCellMap, translatedRangeMap, rowDelta, colDelta),
            ),
      ...(translatedParsedDeps ? { parsedDeps: translatedParsedDeps } : {}),
      ...(translatedParsedSymbolicRefs ? { parsedSymbolicRefs: translatedParsedSymbolicRefs } : {}),
      ...(translatedParsedSymbolicRanges ? { parsedSymbolicRanges: translatedParsedSymbolicRanges } : {}),
    },
  }
}

export function translateCompiledFormula(
  compiled: CompiledFormula,
  rowDelta: number,
  colDelta: number,
  sourceOverride?: string,
): CompiledFormulaTranslationResult {
  const translatedAst = translateNode(compiled.ast, rowDelta, colDelta)
  const translatedOptimizedAst =
    compiled.optimizedAst === compiled.ast ? translatedAst : translateNode(compiled.optimizedAst, rowDelta, colDelta)
  const translatedParsedDeps = compiled.parsedDeps?.map((dependency) => translateParsedDependencyReference(dependency, rowDelta, colDelta))
  const translatedParsedSymbolicRefs = compiled.parsedSymbolicRefs?.map((reference) =>
    translateParsedCellReference(reference, rowDelta, colDelta),
  )
  const translatedParsedSymbolicRanges = compiled.parsedSymbolicRanges?.map((range) =>
    translateParsedRangeReference(range, rowDelta, colDelta),
  )
  const source = sourceOverride ?? serializeFormula(translatedAst)

  return {
    source,
    compiled: {
      ...compiled,
      source,
      ast: translatedAst,
      optimizedAst: translatedOptimizedAst,
      astMatchesSource: true,
      deps:
        translatedParsedDeps?.map((dependency) => formatParsedDependencyReference(dependency)) ??
        compiled.deps.map((dependency) => translateQualifiedDependencyReference(dependency, rowDelta, colDelta)),
      symbolicRefs:
        translatedParsedSymbolicRefs?.map((reference) => formatParsedCellReference(reference)) ??
        compiled.symbolicRefs.map((ref) => translateQualifiedCellReference(ref, rowDelta, colDelta)),
      symbolicRanges:
        translatedParsedSymbolicRanges?.map((range) => formatParsedRangeReference(range)) ??
        compiled.symbolicRanges.map((range) => translateQualifiedRangeReference(range, rowDelta, colDelta)),
      jsPlan: compiled.jsPlan.map((instruction) => translateJsPlanInstruction(instruction, rowDelta, colDelta)),
      ...(translatedParsedDeps ? { parsedDeps: translatedParsedDeps } : {}),
      ...(translatedParsedSymbolicRefs ? { parsedSymbolicRefs: translatedParsedSymbolicRefs } : {}),
      ...(translatedParsedSymbolicRanges ? { parsedSymbolicRanges: translatedParsedSymbolicRanges } : {}),
    },
  }
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

export interface StructuralCompiledFormulaRewriteResult {
  source: string
  compiled: CompiledFormula
  reusedProgram: boolean
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

export function renameFormulaSheetReferences(source: string, oldSheetName: string, newSheetName: string): string {
  const ast = parseFormula(source)
  return serializeFormula(renameNodeSheetReferences(ast, oldSheetName, newSheetName))
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

function translateNode(node: FormulaNode, rowDelta: number, colDelta: number): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'StructuredRef':
      return node
    case 'CellRef':
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta),
      }
    case 'SpillRef':
      return {
        ...node,
        ref: translateCellReference(node.ref, rowDelta, colDelta),
      }
    case 'ColumnRef':
      return {
        ...node,
        ref: translateColumnReference(node.ref, colDelta),
      }
    case 'RowRef':
      return {
        ...node,
        ref: translateRowReference(node.ref, rowDelta),
      }
    case 'RangeRef':
      return {
        ...node,
        start:
          node.refKind === 'cells'
            ? translateCellReference(node.start, rowDelta, colDelta)
            : node.refKind === 'cols'
              ? translateColumnReference(node.start, colDelta)
              : translateRowReference(node.start, rowDelta),
        end:
          node.refKind === 'cells'
            ? translateCellReference(node.end, rowDelta, colDelta)
            : node.refKind === 'cols'
              ? translateColumnReference(node.end, colDelta)
              : translateRowReference(node.end, rowDelta),
      }
    case 'UnaryExpr':
      return {
        ...node,
        argument: translateNode(node.argument, rowDelta, colDelta),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: translateNode(node.left, rowDelta, colDelta),
        right: translateNode(node.right, rowDelta, colDelta),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => translateNode(arg, rowDelta, colDelta)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: translateNode(node.callee, rowDelta, colDelta),
        args: node.args.map((arg) => translateNode(arg, rowDelta, colDelta)),
      }
  }
}

function buildRelativeFormulaTemplateKeyInternal(node: FormulaNode, ownerRow: number, ownerCol: number): string {
  switch (node.kind) {
    case 'NumberLiteral':
      return `n:${node.value}`
    case 'BooleanLiteral':
      return node.value ? 'b:1' : 'b:0'
    case 'StringLiteral':
      return `s:${JSON.stringify(node.value)}`
    case 'ErrorLiteral':
      return `e:${node.code}`
    case 'NameRef':
      return `name:${node.name}`
    case 'StructuredRef':
      return `table:${node.tableName}[${node.columnName}]`
    case 'CellRef':
      return `cell:${templateSheetKey(node.sheetName)}:${buildRelativeCellReferenceKey(node.ref, ownerRow, ownerCol)}`
    case 'SpillRef':
      return `spill:${templateSheetKey(node.sheetName)}:${buildRelativeCellReferenceKey(node.ref, ownerRow, ownerCol)}`
    case 'ColumnRef':
      return `col:${templateSheetKey(node.sheetName)}:${buildRelativeAxisReferenceKey(node.ref, ownerCol, 'column')}`
    case 'RowRef':
      return `row:${templateSheetKey(node.sheetName)}:${buildRelativeAxisReferenceKey(node.ref, ownerRow, 'row')}`
    case 'RangeRef':
      return `range:${node.refKind}:${templateSheetKey(node.sheetName)}:${buildRelativeRangeReferenceKey(node, ownerRow, ownerCol)}`
    case 'UnaryExpr':
      return `unary:${node.operator}:${buildRelativeFormulaTemplateKeyInternal(node.argument, ownerRow, ownerCol)}`
    case 'BinaryExpr':
      return `binary:${node.operator}:${buildRelativeFormulaTemplateKeyInternal(node.left, ownerRow, ownerCol)}:${buildRelativeFormulaTemplateKeyInternal(node.right, ownerRow, ownerCol)}`
    case 'CallExpr':
      return `call:${node.callee}:${node.args.map((arg) => buildRelativeFormulaTemplateKeyInternal(arg, ownerRow, ownerCol)).join('|')}`
    case 'InvokeExpr':
      return `invoke:${buildRelativeFormulaTemplateKeyInternal(node.callee, ownerRow, ownerCol)}:${node.args.map((arg) => buildRelativeFormulaTemplateKeyInternal(arg, ownerRow, ownerCol)).join('|')}`
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
    case 'NameRef':
    case 'StructuredRef':
      return node
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

function templateSheetKey(sheetName: string | undefined): string {
  return sheetName === undefined ? '.' : JSON.stringify(sheetName)
}

function buildRelativeCellReferenceKey(ref: string, ownerRow: number, ownerCol: number): string {
  const parsed = parseCellReferenceParts(ref)
  if (!parsed) {
    return `invalid:${ref}`
  }
  const colKey = parsed.colAbsolute ? `ac${parsed.col}` : `rc${parsed.col - ownerCol}`
  const rowKey = parsed.rowAbsolute ? `ar${parsed.row}` : `rr${parsed.row - ownerRow}`
  return `${colKey}:${rowKey}`
}

function buildRelativeAxisReferenceKey(ref: string, ownerIndex: number, kind: 'row' | 'column'): string {
  const parsed = parseAxisReferenceParts(ref, kind)
  if (!parsed) {
    return `invalid:${ref}`
  }
  return parsed.absolute ? `a${parsed.index}` : `r${parsed.index - ownerIndex}`
}

function buildRelativeRangeReferenceKey(node: Extract<FormulaNode, { kind: 'RangeRef' }>, ownerRow: number, ownerCol: number): string {
  switch (node.refKind) {
    case 'cells':
      return `${buildRelativeCellReferenceKey(node.start, ownerRow, ownerCol)}:${buildRelativeCellReferenceKey(node.end, ownerRow, ownerCol)}`
    case 'rows':
      return `${buildRelativeAxisReferenceKey(node.start, ownerRow, 'row')}:${buildRelativeAxisReferenceKey(node.end, ownerRow, 'row')}`
    case 'cols':
      return `${buildRelativeAxisReferenceKey(node.start, ownerCol, 'column')}:${buildRelativeAxisReferenceKey(node.end, ownerCol, 'column')}`
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

function renameNodeSheetReferences(node: FormulaNode, oldSheetName: string, newSheetName: string): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'StructuredRef':
      return node
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return {
        ...node,
        ...(node.sheetName === oldSheetName ? { sheetName: newSheetName } : {}),
      }
    case 'UnaryExpr':
      return {
        ...node,
        argument: renameNodeSheetReferences(node.argument, oldSheetName, newSheetName),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: renameNodeSheetReferences(node.left, oldSheetName, newSheetName),
        right: renameNodeSheetReferences(node.right, oldSheetName, newSheetName),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => renameNodeSheetReferences(arg, oldSheetName, newSheetName)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: renameNodeSheetReferences(node.callee, oldSheetName, newSheetName),
        args: node.args.map((arg) => renameNodeSheetReferences(arg, oldSheetName, newSheetName)),
      }
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

function translateParsedCellReference<Reference extends ParsedCellReferenceInfo>(
  reference: Reference,
  rowDelta: number,
  colDelta: number,
): Reference {
  const parts =
    reference.rowAbsolute !== undefined && reference.colAbsolute !== undefined && reference.row !== undefined && reference.col !== undefined
      ? {
          row: reference.row,
          col: reference.col,
          rowAbsolute: reference.rowAbsolute,
          colAbsolute: reference.colAbsolute,
        }
      : parseCellReferenceParts(reference.address)
  if (!parts) {
    return reference
  }
  const nextRow = parts.rowAbsolute ? parts.row : parts.row + rowDelta
  const nextCol = parts.colAbsolute ? parts.col : parts.col + colDelta
  const nextLocalAddress = formatCellReference(parts, nextRow, nextCol)
  const nextAddress =
    reference.explicitSheet || reference.sheetName !== undefined
      ? formatQualifiedCellReference(reference.sheetName, nextLocalAddress)
      : nextLocalAddress
  return {
    ...reference,
    address: nextAddress,
    ...(reference.sheetName !== undefined ? { sheetName: reference.sheetName } : {}),
    ...(reference.explicitSheet !== undefined ? { explicitSheet: reference.explicitSheet } : {}),
    ...(reference.row !== undefined ? { row: nextRow } : {}),
    ...(reference.col !== undefined ? { col: nextCol } : {}),
    ...(reference.rowAbsolute !== undefined ? { rowAbsolute: parts.rowAbsolute } : {}),
    ...(reference.colAbsolute !== undefined ? { colAbsolute: parts.colAbsolute } : {}),
  }
}

function translateParsedRangeReference(reference: ParsedRangeReferenceInfo, rowDelta: number, colDelta: number): ParsedRangeReferenceInfo {
  const nextRange = translateParsedRangeReferenceInfo(reference, rowDelta, colDelta)
  const bounds =
    nextRange.refKind === 'cells'
      ? {
          startRow: nextRange.startRow,
          endRow: nextRange.endRow,
          startCol: nextRange.startCol,
          endCol: nextRange.endCol,
        }
      : nextRange.refKind === 'rows'
        ? {
            startRow: nextRange.startRow,
            endRow: nextRange.endRow,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: nextRange.startCol,
            endCol: nextRange.endCol,
          }
  return {
    ...reference,
    address: formatParsedRangeReference(nextRange),
    refKind: nextRange.refKind,
    startAddress: nextRange.startAddress,
    endAddress: nextRange.endAddress,
    ...bounds,
    ...(reference.explicitSheet !== undefined ? { explicitSheet: reference.explicitSheet } : {}),
  }
}

function translateParsedDependencyReference(
  reference: ParsedDependencyReference,
  rowDelta: number,
  colDelta: number,
): ParsedDependencyReference {
  return reference.kind === 'cell'
    ? translateParsedCellReference(reference, rowDelta, colDelta)
    : translateParsedRangeReference(reference, rowDelta, colDelta)
}

function translateQualifiedCellReference(raw: string, rowDelta: number, colDelta: number): string {
  const explicitlyQualified = raw.includes('!')
  const parsed = parseCellAddress(raw)
  const nextAddress = translateCellReference(parsed.text, rowDelta, colDelta)
  return explicitlyQualified ? formatQualifiedCellReference(parsed.sheetName, nextAddress) : nextAddress
}

function formatParsedCellReference(reference: ParsedCellReferenceInfo): string {
  const localAddress = formatParsedLocalCellReference(reference)
  return reference.explicitSheet || reference.sheetName !== undefined
    ? formatQualifiedCellReference(reference.sheetName, localAddress)
    : localAddress
}

function formatParsedLocalCellReference(reference: ParsedCellReferenceInfo): string {
  const parts =
    reference.row !== undefined && reference.col !== undefined && reference.rowAbsolute !== undefined && reference.colAbsolute !== undefined
      ? {
          row: reference.row,
          col: reference.col,
          rowAbsolute: reference.rowAbsolute,
          colAbsolute: reference.colAbsolute,
        }
      : parseCellReferenceParts(reference.address)
  if (!parts) {
    return stripSheetQualifier(reference.address)
  }
  return formatCellReference(parts, parts.row, parts.col)
}

function formatParsedRangeReference(reference: ParsedRangeReferenceInfo): string {
  return formatQualifiedRangeReference(
    reference.explicitSheet ? reference.sheetName : undefined,
    reference.startAddress,
    reference.endAddress,
  )
}

function stripSheetQualifier(reference: string): string {
  const bang = reference.lastIndexOf('!')
  return bang === -1 ? reference : reference.slice(bang + 1)
}

function translatedCellInstructionKey(sheetName: string | undefined, address: string): string {
  return `${sheetName ?? ''}\t${address}`
}

function translatedRangeInstructionKey(
  sheetName: string | undefined,
  refKind: 'cells' | 'rows' | 'cols',
  start: string,
  end: string,
): string {
  return `${sheetName ?? ''}\t${refKind}\t${start}\t${end}`
}

function buildTranslatedCellReferenceMap(
  original: readonly ParsedCellReferenceInfo[] | undefined,
  translated: readonly ParsedCellReferenceInfo[] | undefined,
): Map<string, ParsedCellReferenceInfo> {
  const output = new Map<string, ParsedCellReferenceInfo>()
  if (!original || !translated || original.length !== translated.length) {
    return output
  }
  for (let index = 0; index < original.length; index += 1) {
    const source = original[index]
    const target = translated[index]
    if (!source || !target) {
      continue
    }
    output.set(translatedCellInstructionKey(source.sheetName, formatParsedLocalCellReference(source)), target)
  }
  return output
}

function buildTranslatedRangeReferenceMap(
  original: readonly ParsedRangeReferenceInfo[] | undefined,
  translated: readonly ParsedRangeReferenceInfo[] | undefined,
): Map<string, ParsedRangeReferenceInfo> {
  const output = new Map<string, ParsedRangeReferenceInfo>()
  if (!original || !translated || original.length !== translated.length) {
    return output
  }
  for (let index = 0; index < original.length; index += 1) {
    const source = original[index]
    const target = translated[index]
    if (!source || !target) {
      continue
    }
    output.set(translatedRangeInstructionKey(source.sheetName, source.refKind, source.startAddress, source.endAddress), target)
  }
  return output
}

function formatParsedDependencyReference(reference: ParsedDependencyReference): string {
  return reference.kind === 'cell' ? formatParsedCellReference(reference) : formatParsedRangeReference(reference)
}

function translateQualifiedDependencyReference(raw: string, rowDelta: number, colDelta: number): string {
  if (!raw.includes(':')) {
    return translateQualifiedCellReference(raw, rowDelta, colDelta)
  }
  return translateQualifiedRangeReference(raw, rowDelta, colDelta)
}

function translateQualifiedRangeReference(raw: string, rowDelta: number, colDelta: number): string {
  const explicitlyQualified = raw.includes('!')
  const parsed = parseRangeAddress(raw)
  const nextRange = translateRangeAddress(parsed, rowDelta, colDelta)
  if (explicitlyQualified) {
    return formatRangeAddress(nextRange)
  }
  return `${nextRange.start.text}:${nextRange.end.text}`
}

function translateRangeAddress(range: RangeAddress, rowDelta: number, colDelta: number): RangeAddress {
  switch (range.kind) {
    case 'cells': {
      const startAddress = translateCellReference(range.start.text, rowDelta, colDelta)
      const endAddress = translateCellReference(range.end.text, rowDelta, colDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, startAddress, endAddress))
    }
    case 'rows': {
      const start = translateRowReference(range.start.text, rowDelta)
      const end = translateRowReference(range.end.text, rowDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, start, end))
    }
    case 'cols': {
      const start = translateColumnReference(range.start.text, colDelta)
      const end = translateColumnReference(range.end.text, colDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, start, end))
    }
  }
}

function translateParsedRangeReferenceInfo(
  reference: ParsedRangeReferenceInfo,
  rowDelta: number,
  colDelta: number,
): ParsedRangeReferenceInfo {
  if (reference.refKind === 'cells') {
    const startRow = (reference.startRowAbsolute ?? false) ? reference.startRow : reference.startRow + rowDelta
    const endRow = (reference.endRowAbsolute ?? false) ? reference.endRow : reference.endRow + rowDelta
    const startCol = (reference.startColAbsolute ?? false) ? reference.startCol : reference.startCol + colDelta
    const endCol = (reference.endColAbsolute ?? false) ? reference.endCol : reference.endCol + colDelta
    const startAddress = formatCellReference(
      {
        row: reference.startRow,
        col: reference.startCol,
        rowAbsolute: reference.startRowAbsolute ?? false,
        colAbsolute: reference.startColAbsolute ?? false,
      },
      startRow,
      startCol,
    )
    const endAddress = formatCellReference(
      {
        row: reference.endRow,
        col: reference.endCol,
        rowAbsolute: reference.endRowAbsolute ?? false,
        colAbsolute: reference.endColAbsolute ?? false,
      },
      endRow,
      endCol,
    )
    return {
      ...reference,
      startAddress,
      endAddress,
      startRow,
      endRow,
      startCol,
      endCol,
    }
  }
  if (reference.refKind === 'rows') {
    const startRow = (reference.startRowAbsolute ?? false) ? reference.startRow : reference.startRow + rowDelta
    const endRow = (reference.endRowAbsolute ?? false) ? reference.endRow : reference.endRow + rowDelta
    return {
      ...reference,
      startAddress: formatAxisReference(reference.startRowAbsolute ?? false, startRow, 'row'),
      endAddress: formatAxisReference(reference.endRowAbsolute ?? false, endRow, 'row'),
      startRow,
      endRow,
      startCol: 0,
      endCol: 0,
    }
  }
  const startCol = (reference.startColAbsolute ?? false) ? reference.startCol : reference.startCol + colDelta
  const endCol = (reference.endColAbsolute ?? false) ? reference.endCol : reference.endCol + colDelta
  return {
    ...reference,
    startAddress: formatAxisReference(reference.startColAbsolute ?? false, startCol, 'column'),
    endAddress: formatAxisReference(reference.endColAbsolute ?? false, endCol, 'column'),
    startRow: 0,
    endRow: 0,
    startCol,
    endCol,
  }
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

function translateJsPlanInstruction(instruction: JsPlanInstruction, rowDelta: number, colDelta: number): JsPlanInstruction {
  switch (instruction.opcode) {
    case 'push-cell':
      return {
        ...instruction,
        address: translateCellReference(instruction.address, rowDelta, colDelta),
      }
    case 'push-range': {
      const nextRange = translatePlanRangeInstruction(instruction.sheetName, instruction.start, instruction.end, rowDelta, colDelta)
      return { ...instruction, ...nextRange }
    }
    case 'lookup-exact-match':
    case 'lookup-approximate-match': {
      const nextRange = translatePlanRangeInstruction(instruction.sheetName, instruction.start, instruction.end, rowDelta, colDelta)
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
            argRefs: instruction.argRefs.map((argRef) => (argRef ? translateReferenceOperand(argRef, rowDelta, colDelta) : argRef)),
          }
        : instruction
    case 'push-lambda':
      return {
        ...instruction,
        body: instruction.body.map((step) => translateJsPlanInstruction(step, rowDelta, colDelta)),
      }
    case 'push-number':
    case 'push-boolean':
    case 'push-string':
    case 'push-error':
    case 'push-name':
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

function translateJsPlanInstructionWithoutAst(
  instruction: JsPlanInstruction,
  translatedCellMap: ReadonlyMap<string, ParsedCellReferenceInfo>,
  translatedRangeMap: ReadonlyMap<string, ParsedRangeReferenceInfo>,
  rowDelta: number,
  colDelta: number,
): JsPlanInstruction {
  switch (instruction.opcode) {
    case 'push-cell': {
      const translated = translatedCellMap.get(translatedCellInstructionKey(instruction.sheetName, instruction.address))
      return translated
        ? {
            ...instruction,
            address: formatParsedLocalCellReference(translated),
          }
        : translateJsPlanInstruction(instruction, rowDelta, colDelta)
    }
    case 'push-range': {
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(instruction.sheetName, instruction.refKind, instruction.start, instruction.end),
      )
      return translated
        ? {
            ...instruction,
            start: translated.startAddress,
            end: translated.endAddress,
          }
        : translateJsPlanInstruction(instruction, rowDelta, colDelta)
    }
    case 'lookup-exact-match':
    case 'lookup-approximate-match': {
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(instruction.sheetName, instruction.refKind, instruction.start, instruction.end),
      )
      if (!translated || translated.refKind !== 'cells') {
        return translateJsPlanInstruction(instruction, rowDelta, colDelta)
      }
      return {
        ...instruction,
        start: translated.startAddress,
        end: translated.endAddress,
        startRow: translated.startRow,
        endRow: translated.endRow,
        startCol: translated.startCol,
        endCol: translated.endCol,
      }
    }
    case 'call':
      return instruction.argRefs
        ? {
            ...instruction,
            argRefs: instruction.argRefs.map((argRef) =>
              argRef ? translateReferenceOperandWithoutAst(argRef, translatedCellMap, translatedRangeMap, rowDelta, colDelta) : argRef,
            ),
          }
        : instruction
    case 'push-lambda':
      return {
        ...instruction,
        body: instruction.body.map((step) =>
          translateJsPlanInstructionWithoutAst(step, translatedCellMap, translatedRangeMap, rowDelta, colDelta),
        ),
      }
    case 'push-number':
    case 'push-boolean':
    case 'push-string':
    case 'push-error':
    case 'push-name':
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

function translateReferenceOperand(operand: ReferenceOperand, rowDelta: number, colDelta: number): ReferenceOperand {
  switch (operand.kind) {
    case 'cell':
      return operand.address
        ? {
            ...operand,
            address: translateCellReference(operand.address, rowDelta, colDelta),
          }
        : operand
    case 'range':
      if (!operand.start || !operand.end || !operand.refKind) {
        return operand
      }
      return {
        ...operand,
        ...translatePlanRangeInstruction(operand.sheetName, operand.start, operand.end, rowDelta, colDelta),
      }
    case 'row':
      return operand.address
        ? {
            ...operand,
            address: translateRowReference(operand.address, rowDelta),
          }
        : operand
    case 'col':
      return operand.address
        ? {
            ...operand,
            address: translateColumnReference(operand.address, colDelta),
          }
        : operand
  }
}

function translateReferenceOperandWithoutAst(
  operand: ReferenceOperand,
  translatedCellMap: ReadonlyMap<string, ParsedCellReferenceInfo>,
  translatedRangeMap: ReadonlyMap<string, ParsedRangeReferenceInfo>,
  rowDelta: number,
  colDelta: number,
): ReferenceOperand {
  switch (operand.kind) {
    case 'cell': {
      if (!operand.address) {
        return operand
      }
      const translated = translatedCellMap.get(translatedCellInstructionKey(operand.sheetName, operand.address))
      return translated
        ? {
            ...operand,
            address: formatParsedLocalCellReference(translated),
          }
        : translateReferenceOperand(operand, rowDelta, colDelta)
    }
    case 'range': {
      if (!operand.start || !operand.end || !operand.refKind) {
        return operand
      }
      const translated = translatedRangeMap.get(
        translatedRangeInstructionKey(operand.sheetName, operand.refKind, operand.start, operand.end),
      )
      return translated
        ? {
            ...operand,
            start: translated.startAddress,
            end: translated.endAddress,
            refKind: translated.refKind,
          }
        : translateReferenceOperand(operand, rowDelta, colDelta)
    }
    case 'row':
    case 'col':
      return translateReferenceOperand(operand, rowDelta, colDelta)
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

function translatePlanRangeInstruction(
  explicitSheetName: string | undefined,
  start: string,
  end: string,
  rowDelta: number,
  colDelta: number,
): { start: string; end: string } {
  const parsed = parseRangeAddress(formatQualifiedRangeReference(explicitSheetName, start, end))
  const nextRange = translateRangeAddress(parsed, rowDelta, colDelta)
  return {
    start: nextRange.start.text,
    end: nextRange.end.text,
  }
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

function translateCellReference(ref: string, rowDelta: number, colDelta: number): string {
  const parsed = parseCellReferenceParts(ref)
  if (!parsed) {
    throw new Error(`Invalid cell reference '${ref}'`)
  }
  const nextCol = parsed.colAbsolute ? parsed.col : parsed.col + colDelta
  const nextRow = parsed.rowAbsolute ? parsed.row : parsed.row + rowDelta
  if (nextCol < 0 || nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatCellReference(parsed, nextRow, nextCol)
}

function translateColumnReference(ref: string, colDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, 'column')
  if (!parsed) {
    throw new Error(`Invalid column reference '${ref}'`)
  }
  const nextCol = parsed.absolute ? parsed.index : parsed.index + colDelta
  if (nextCol < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatAxisReference(parsed.absolute, nextCol, 'column')
}

function translateRowReference(ref: string, rowDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, 'row')
  if (!parsed) {
    throw new Error(`Invalid row reference '${ref}'`)
  }
  const nextRow = parsed.absolute ? parsed.index : parsed.index + rowDelta
  if (nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatAxisReference(parsed.absolute, nextRow, 'row')
}

export function serializeFormula(node: FormulaNode, parentPrecedence = 0, parentAssociativity: 'left' | 'right' | null = null): string {
  switch (node.kind) {
    case 'NumberLiteral':
      return String(node.value)
    case 'BooleanLiteral':
      return node.value ? 'TRUE' : 'FALSE'
    case 'StringLiteral':
      return `"${node.value.replaceAll('"', '""')}"`
    case 'ErrorLiteral':
      return ERROR_LITERAL_TEXT[node.code] ?? '#ERROR!'
    case 'NameRef':
      return node.name
    case 'StructuredRef':
      return `${node.tableName}[${node.columnName}]`
    case 'CellRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'SpillRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}#`
    case 'ColumnRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'RowRef':
      return `${formatSheetPrefix(node.sheetName)}${node.ref}`
    case 'RangeRef':
      return `${formatSheetPrefix(node.sheetName)}${node.start}:${node.end}`
    case 'UnaryExpr':
      return `${node.operator}${serializeFormula(node.argument, 6)}`
    case 'CallExpr':
      return `${node.callee}(${node.args.map((arg) => serializeFormula(arg)).join(',')})`
    case 'InvokeExpr': {
      const callee =
        node.callee.kind === 'CallExpr' || node.callee.kind === 'InvokeExpr'
          ? serializeFormula(node.callee)
          : `(${serializeFormula(node.callee)})`
      return `${callee}(${node.args.map((arg) => serializeFormula(arg)).join(',')})`
    }
    case 'BinaryExpr': {
      const precedence = BINARY_PRECEDENCE[node.operator]
      const isRightAssociative = node.operator === '^'
      const left = serializeFormula(node.left, precedence, 'left')
      const right = serializeFormula(node.right, precedence, 'right')
      const output = `${left}${node.operator}${right}`
      const needsParens =
        precedence < parentPrecedence ||
        (precedence === parentPrecedence &&
          ((parentAssociativity === 'left' && isRightAssociative) || (parentAssociativity === 'right' && !isRightAssociative)))
      return needsParens ? `(${output})` : output
    }
  }
}

function formatSheetPrefix(sheetName?: string): string {
  if (!sheetName) {
    return ''
  }
  return `${quoteSheetNameIfNeeded(sheetName)}!`
}

function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function columnToIndex(column: string): number {
  let value = 0
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64)
  }
  return value - 1
}

function indexToColumn(index: number): string {
  let current = index + 1
  let output = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    current = Math.floor((current - 1) / 26)
  }
  return output
}

function targetsSheet(explicitSheetName: string | undefined, ownerSheetName: string, targetSheetName: string): boolean {
  return (explicitSheetName ?? ownerSheetName) === targetSheetName
}

interface ParsedCellReference {
  colAbsolute: boolean
  rowAbsolute: boolean
  col: number
  row: number
}

function parseCellReferenceParts(ref: string): ParsedCellReference | undefined {
  const match = CELL_REF_RE.exec(ref.toUpperCase())
  if (!match) {
    return undefined
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match
  return {
    colAbsolute: colAbsolute === '$',
    rowAbsolute: rowAbsolute === '$',
    col: columnToIndex(columnText!),
    row: Number.parseInt(rowText!, 10) - 1,
  }
}

function formatCellReference(parts: ParsedCellReference, row: number, col: number): string {
  return `${parts.colAbsolute ? '$' : ''}${indexToColumn(col)}${parts.rowAbsolute ? '$' : ''}${row + 1}`
}

interface ParsedAxisReference {
  absolute: boolean
  index: number
}

function parseAxisReferenceParts(ref: string, kind: StructuralAxisKind): ParsedAxisReference | undefined {
  const match = (kind === 'row' ? ROW_REF_RE : COLUMN_REF_RE).exec(ref.toUpperCase())
  if (!match) {
    return undefined
  }
  return kind === 'row'
    ? {
        absolute: match[1] === '$',
        index: Number.parseInt(match[2]!, 10) - 1,
      }
    : {
        absolute: match[1] === '$',
        index: columnToIndex(match[2]!),
      }
}

function formatAxisReference(absolute: boolean, index: number, kind: StructuralAxisKind): string {
  const prefix = absolute ? '$' : ''
  return kind === 'row' ? `${prefix}${index + 1}` : `${prefix}${indexToColumn(index)}`
}

function mapPointIndex(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case 'insert':
      return index >= transform.start ? index + transform.count : index
    case 'delete':
      if (index < transform.start) {
        return index
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count
      }
      return undefined
    case 'move':
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.start) {
          return index + transform.count
        }
      } else if (transform.target > transform.start) {
        if (index >= transform.start + transform.count && index < transform.target + transform.count) {
          return index - transform.count
        }
      }
      if (index >= transform.start && index < transform.start + transform.count) {
        return transform.target + (index - transform.start)
      }
      return index
    default:
      return assertNever(transform)
  }
}

function mapInterval(start: number, end: number, transform: StructuralAxisTransform): { start: number; end: number } | undefined {
  switch (transform.kind) {
    case 'insert': {
      if (transform.start <= start) {
        return { start: start + transform.count, end: end + transform.count }
      }
      if (transform.start <= end) {
        return { start, end: end + transform.count }
      }
      return { start, end }
    }
    case 'delete': {
      const deleteEnd = transform.start + transform.count - 1
      if (deleteEnd < start) {
        return { start: start - transform.count, end: end - transform.count }
      }
      if (transform.start > end) {
        return { start, end }
      }
      const survivingStart = start < transform.start ? start : deleteEnd + 1
      const survivingEnd = end > deleteEnd ? end : transform.start - 1
      if (survivingStart > survivingEnd) {
        return undefined
      }
      const nextStart = mapPointIndex(survivingStart, transform)
      const nextEnd = mapPointIndex(survivingEnd, transform)
      return nextStart === undefined || nextEnd === undefined ? undefined : { start: nextStart, end: nextEnd }
    }
    case 'move': {
      const segments =
        transform.target < transform.start
          ? [
              { start: 0, end: transform.target - 1, delta: 0 },
              { start: transform.target, end: transform.start - 1, delta: transform.count },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              { start: transform.start + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
          : [
              { start: 0, end: transform.start - 1, delta: 0 },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              {
                start: transform.start + transform.count,
                end: transform.target + transform.count - 1,
                delta: -transform.count,
              },
              { start: transform.target + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
      let nextStart: number | undefined
      let nextEnd: number | undefined
      segments.forEach((segment) => {
        const overlapStart = Math.max(start, segment.start)
        const overlapEnd = Math.min(end, segment.end)
        if (overlapStart > overlapEnd) {
          return
        }
        const mappedStart = overlapStart + segment.delta
        const mappedEnd = overlapEnd + segment.delta
        nextStart = nextStart === undefined ? mappedStart : Math.min(nextStart, mappedStart)
        nextEnd = nextEnd === undefined ? mappedEnd : Math.max(nextEnd, mappedEnd)
      })
      if (nextStart === undefined || nextEnd === undefined) {
        return undefined
      }
      return { start: nextStart, end: nextEnd }
    }
    default:
      return assertNever(transform)
  }
}
