import {
  type CompiledFormula,
  type DirectAggregateCandidate,
  type FormulaNode,
  type ParsedCellReferenceInfo,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
  type StructuralAxisTransform,
  parseFormula,
  parseRangeAddress,
} from '@bilig/formula'
import { MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import { mapStructuralAxisInterval } from '../../engine-structural-utils.js'
import { resolveRuntimeDirectLookupBinding } from '../direct-vector-lookup.js'
import type {
  EngineRuntimeState,
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectCriteriaResultTransform,
  RuntimeDirectLookupDescriptor,
  RuntimeDirectScalarOperand,
} from '../runtime-state.js'
import type { RegionGraph } from '../../deps/region-graph.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'

export type ParsedCompiledFormula = CompiledFormula & {
  parsedDeps?: ParsedDependencyReference[]
  parsedSymbolicRefs?: ParsedCellReferenceInfo[]
  parsedSymbolicRanges?: ParsedRangeReferenceInfo[]
  directAggregateCandidate?: DirectAggregateCandidate
}

function internDirectAggregateRegions(args: {
  readonly regionGraph: RegionGraph
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}): { readonly regionId: number; readonly regionIds?: readonly number[] } {
  const regionIds: number[] = []
  for (let col = args.colStart; col <= args.colEnd; col += 1) {
    regionIds.push(
      args.regionGraph.internSingleColumnRegion({
        sheetName: args.sheetName,
        rowStart: args.rowStart,
        rowEnd: args.rowEnd,
        col,
      }),
    )
  }
  return regionIds.length === 1 ? { regionId: regionIds[0]! } : { regionId: regionIds[0]!, regionIds }
}

export function renameDirectAggregateDescriptorSheet(args: {
  readonly descriptor: RuntimeDirectAggregateDescriptor
  readonly oldSheetName: string
  readonly newSheetName: string
  readonly regionGraph: RegionGraph
}): RuntimeDirectAggregateDescriptor {
  if (args.descriptor.sheetName !== args.oldSheetName) {
    return args.descriptor
  }
  const descriptorColEnd = args.descriptor.colEnd ?? args.descriptor.col
  const nextRegions = internDirectAggregateRegions({
    regionGraph: args.regionGraph,
    sheetName: args.newSheetName,
    rowStart: args.descriptor.rowStart,
    rowEnd: args.descriptor.rowEnd,
    colStart: args.descriptor.col,
    colEnd: descriptorColEnd,
  })
  return {
    ...args.descriptor,
    ...nextRegions,
    colEnd: descriptorColEnd,
    sheetName: args.newSheetName,
  }
}

export function rewriteDirectAggregateDescriptorForStructuralTransform(args: {
  readonly descriptor: RuntimeDirectAggregateDescriptor
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
  readonly regionGraph: RegionGraph
}): RuntimeDirectAggregateDescriptor | undefined {
  if (args.descriptor.sheetName !== args.targetSheetName) {
    return undefined
  }
  const nextRows =
    args.transform.axis === 'row'
      ? mapStructuralAxisInterval(args.descriptor.rowStart, args.descriptor.rowEnd, args.transform)
      : { start: args.descriptor.rowStart, end: args.descriptor.rowEnd }
  const descriptorColEnd = args.descriptor.colEnd ?? args.descriptor.col
  const nextCols =
    args.transform.axis === 'column'
      ? mapStructuralAxisInterval(args.descriptor.col, descriptorColEnd, args.transform)
      : { start: args.descriptor.col, end: descriptorColEnd }
  if (!nextRows || !nextCols) {
    return undefined
  }
  const regions = internDirectAggregateRegions({
    regionGraph: args.regionGraph,
    sheetName: args.descriptor.sheetName,
    rowStart: nextRows.start,
    rowEnd: nextRows.end,
    colStart: nextCols.start,
    colEnd: nextCols.end,
  })
  return {
    ...args.descriptor,
    ...regions,
    rowStart: nextRows.start,
    rowEnd: nextRows.end,
    col: nextCols.start,
    colEnd: nextCols.end,
    length: (nextRows.end - nextRows.start + 1) * (nextCols.end - nextCols.start + 1),
  }
}

function staticCellValue(node: FormulaNode | undefined): CellValue | undefined {
  if (!node) {
    return undefined
  }
  switch (node.kind) {
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'ErrorLiteral':
      return { tag: ValueTag.Error, code: node.code }
    case 'NumberLiteral':
      return { tag: ValueTag.Number, value: node.value }
    case 'StringLiteral':
      return { tag: ValueTag.String, value: node.value, stringId: 0 }
    case 'UnaryExpr':
      if (node.operator === '-' && node.argument.kind === 'NumberLiteral') {
        return { tag: ValueTag.Number, value: -node.argument.value }
      }
      return undefined
    case 'ArrayConstant':
      return undefined
    case 'CellRef':
    case 'CallExpr':
    case 'BinaryExpr':
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return undefined
  }
}

function flattenCriteriaProduct(node: FormulaNode): FormulaNode[] {
  if (node.kind !== 'BinaryExpr' || node.operator !== '*') {
    return [node]
  }
  return [...flattenCriteriaProduct(node.left), ...flattenCriteriaProduct(node.right)]
}

function resolveDirectCriteriaRange(
  node: FormulaNode | undefined,
  ownerSheetName: string,
):
  | {
      sheetName: string
      rowStart: number
      rowEnd: number
      col: number
      length: number
    }
  | undefined {
  if (!node || node.kind !== 'RangeRef' || node.sheetEndName !== undefined) {
    return undefined
  }
  const parsed = parseRangeAddress(`${node.start}:${node.end}`, node.sheetName ?? ownerSheetName)
  const sheetName = parsed.sheetName ?? node.sheetName ?? ownerSheetName
  if (parsed.kind === 'cols') {
    if (parsed.start.col !== parsed.end.col) {
      return undefined
    }
    return {
      sheetName,
      rowStart: 0,
      rowEnd: MAX_ROWS - 1,
      col: parsed.start.col,
      length: MAX_ROWS,
    }
  }
  if (parsed.kind !== 'cells' || parsed.start.col !== parsed.end.col) {
    return undefined
  }
  return {
    sheetName,
    rowStart: parsed.start.row,
    rowEnd: parsed.end.row,
    col: parsed.start.col,
    length: parsed.end.row - parsed.start.row + 1,
  }
}

type DirectCriteriaResolvedRange = NonNullable<ReturnType<typeof resolveDirectCriteriaRange>>
type DirectCriteriaCallNode = Extract<FormulaNode, { readonly kind: 'CallExpr' }>
type DirectCriteriaCellRefNode = Extract<FormulaNode, { readonly kind: 'CellRef' }>
const DIRECT_CRITERIA_PLAN_CALLEES = new Set([
  'COUNTIF',
  'COUNTIFS',
  'SUMIF',
  'SUMIFS',
  'AVERAGEIF',
  'AVERAGEIFS',
  'MINIFS',
  'MAXIFS',
  'INDEX',
])

function callName(node: FormulaNode | undefined): string | undefined {
  return node?.kind === 'CallExpr' ? node.callee.trim().toUpperCase() : undefined
}

function mayContainDirectCriteriaPlanCall(compiled: ParsedCompiledFormula): boolean {
  return compiled.jsPlan.some((instruction) => {
    const opcode = Reflect.get(instruction, 'opcode')
    const callee = Reflect.get(instruction, 'callee')
    return opcode === 'call' && typeof callee === 'string' && DIRECT_CRITERIA_PLAN_CALLEES.has(callee.trim().toUpperCase())
  })
}

function appendDirectCriteriaResultTransform(
  descriptor: RuntimeDirectCriteriaDescriptor,
  transform: RuntimeDirectCriteriaResultTransform,
): RuntimeDirectCriteriaDescriptor {
  return {
    ...descriptor,
    resultTransforms: [...(descriptor.resultTransforms ?? []), transform],
  }
}

export function buildDirectCriteriaDescriptor(args: {
  readonly compiled: ParsedCompiledFormula
  readonly source: string
  readonly ownerSheetName: string
  readonly workbook: Pick<EngineRuntimeState, 'workbook'>['workbook']
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly regionGraph: RegionGraph
}): RuntimeDirectCriteriaDescriptor | undefined {
  const sameCellRef = (left: DirectCriteriaCellRefNode, right: DirectCriteriaCellRefNode): boolean => {
    return (
      left.ref.trim().toUpperCase() === right.ref.trim().toUpperCase() &&
      (left.sheetName ?? args.ownerSheetName).trim().toUpperCase() === (right.sheetName ?? args.ownerSheetName).trim().toUpperCase()
    )
  }

  const resolveDatePartCellRef = (
    node: FormulaNode | undefined,
    expectedCallee: 'YEAR' | 'MONTH',
  ): DirectCriteriaCellRefNode | undefined => {
    if (callName(node) !== expectedCallee || node?.kind !== 'CallExpr' || node.args.length !== 1) {
      return undefined
    }
    const cellRef = node.args[0]
    return cellRef?.kind === 'CellRef' ? cellRef : undefined
  }

  const resolveMonthStartCellRef = (node: FormulaNode | undefined): DirectCriteriaCellRefNode | undefined => {
    if (callName(node) !== 'DATE' || node?.kind !== 'CallExpr' || node.args.length !== 3) {
      return undefined
    }
    const yearCell = resolveDatePartCellRef(node.args[0], 'YEAR')
    const monthCell = resolveDatePartCellRef(node.args[1], 'MONTH')
    const day = staticCellValue(node.args[2])
    if (!yearCell || !monthCell || day?.tag !== ValueTag.Number || !Object.is(day.value, 1) || !sameCellRef(yearCell, monthCell)) {
      return undefined
    }
    return yearCell
  }

  const resolveMonthBoundaryCellRef = (
    node: FormulaNode | undefined,
  ): { readonly cellRef: DirectCriteriaCellRefNode; readonly offsetMonths: number } | undefined => {
    const monthStartCell = resolveMonthStartCellRef(node)
    if (monthStartCell) {
      return { cellRef: monthStartCell, offsetMonths: 0 }
    }
    if (callName(node) !== 'EDATE' || node?.kind !== 'CallExpr' || node.args.length !== 2) {
      return undefined
    }
    const startCell = resolveMonthStartCellRef(node.args[0])
    const offset = staticCellValue(node.args[1])
    if (!startCell || offset?.tag !== ValueTag.Number || !Number.isFinite(offset.value)) {
      return undefined
    }
    return { cellRef: startCell, offsetMonths: offset.value }
  }

  const buildMonthBoundaryCriterion = (
    boundary: { readonly cellRef: DirectCriteriaCellRefNode; readonly offsetMonths: number },
    prefix: string,
    suffix: string,
  ): RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined => {
    const sheetName = boundary.cellRef.sheetName ?? args.ownerSheetName
    if (!args.workbook.getSheet(sheetName)) {
      return undefined
    }
    return {
      kind: 'cell-month-boundary-string-concat',
      cellIndex: args.ensureCellTracked(sheetName, boundary.cellRef.ref),
      prefix,
      suffix,
      offsetMonths: boundary.offsetMonths,
    }
  }

  const resolveEmptyStringCellGuard = (
    node: FormulaNode | undefined,
  ): { readonly cellRef: DirectCriteriaCellRefNode; readonly operator: '=' | '<>' } | undefined => {
    if (!node || node.kind !== 'BinaryExpr' || (node.operator !== '=' && node.operator !== '<>')) {
      return undefined
    }
    const leftLiteral = staticCellValue(node.left)
    const rightLiteral = staticCellValue(node.right)
    const leftCell = node.left.kind === 'CellRef' ? node.left : undefined
    const rightCell = node.right.kind === 'CellRef' ? node.right : undefined
    if (leftCell && rightLiteral?.tag === ValueTag.String && rightLiteral.value === '') {
      return { cellRef: leftCell, operator: node.operator }
    }
    if (rightCell && leftLiteral?.tag === ValueTag.String && leftLiteral.value === '') {
      return { cellRef: rightCell, operator: node.operator }
    }
    return undefined
  }

  const appendIfEmptyCellTransform = (
    descriptor: RuntimeDirectCriteriaDescriptor | undefined,
    cellRef: DirectCriteriaCellRefNode,
    fallback: CellValue | undefined,
  ): RuntimeDirectCriteriaDescriptor | undefined => {
    if (!descriptor || !fallback) {
      return undefined
    }
    const sheetName = cellRef.sheetName ?? args.ownerSheetName
    if (!args.workbook.getSheet(sheetName)) {
      return undefined
    }
    return appendDirectCriteriaResultTransform(descriptor, {
      kind: 'if-empty-cell',
      cellIndex: args.ensureCellTracked(sheetName, cellRef.ref),
      fallback,
    })
  }

  const resolveCriterionOperand = (
    criterionNode: FormulaNode | undefined,
  ): RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined => {
    if (!criterionNode) {
      return undefined
    }
    if (criterionNode.kind === 'CellRef') {
      const sheetName = criterionNode.sheetName ?? args.ownerSheetName
      if (!args.workbook.getSheet(sheetName)) {
        return undefined
      }
      return {
        kind: 'cell',
        cellIndex: args.ensureCellTracked(sheetName, criterionNode.ref),
      }
    }
    if (criterionNode.kind === 'BinaryExpr' && criterionNode.operator === '&') {
      const leftLiteral = staticCellValue(criterionNode.left)
      const rightLiteral = staticCellValue(criterionNode.right)
      const leftCell = criterionNode.left.kind === 'CellRef' ? criterionNode.left : undefined
      const rightCell = criterionNode.right.kind === 'CellRef' ? criterionNode.right : undefined
      const leftMonthBoundary = resolveMonthBoundaryCellRef(criterionNode.left)
      const rightMonthBoundary = resolveMonthBoundaryCellRef(criterionNode.right)
      if (leftLiteral?.tag === ValueTag.String && rightMonthBoundary) {
        return buildMonthBoundaryCriterion(rightMonthBoundary, leftLiteral.value, '')
      }
      if (rightLiteral?.tag === ValueTag.String && leftMonthBoundary) {
        return buildMonthBoundaryCriterion(leftMonthBoundary, '', rightLiteral.value)
      }
      if (leftLiteral?.tag === ValueTag.String && rightCell) {
        const sheetName = rightCell.sheetName ?? args.ownerSheetName
        if (!args.workbook.getSheet(sheetName)) {
          return undefined
        }
        return {
          kind: 'cell-string-concat',
          cellIndex: args.ensureCellTracked(sheetName, rightCell.ref),
          prefix: leftLiteral.value,
          suffix: '',
        }
      }
      if (rightLiteral?.tag === ValueTag.String && leftCell) {
        const sheetName = leftCell.sheetName ?? args.ownerSheetName
        if (!args.workbook.getSheet(sheetName)) {
          return undefined
        }
        return {
          kind: 'cell-string-concat',
          cellIndex: args.ensureCellTracked(sheetName, leftCell.ref),
          prefix: '',
          suffix: rightLiteral.value,
        }
      }
    }
    const literal = staticCellValue(criterionNode)
    return literal ? { kind: 'literal', value: literal } : undefined
  }

  const withRangeRegion = (range: DirectCriteriaResolvedRange): DirectCriteriaResolvedRange & { readonly regionId: number } => ({
    ...range,
    regionId: args.regionGraph.internSingleColumnRegion({
      sheetName: range.sheetName,
      rowStart: range.rowStart,
      rowEnd: range.rowEnd,
      col: range.col,
    }),
  })

  const resolveIndexOffsetOperand = (node: FormulaNode | undefined): RuntimeDirectScalarOperand | undefined => {
    if (!node) {
      return undefined
    }
    const literal = staticCellValue(node)
    if (literal?.tag === ValueTag.Number) {
      return { kind: 'literal-number', value: literal.value }
    }
    if (node.kind !== 'CellRef') {
      return undefined
    }
    const sheetName = node.sheetName ?? args.ownerSheetName
    if (!args.workbook.getSheet(sheetName)) {
      return undefined
    }
    return {
      kind: 'cell',
      cellIndex: args.ensureCellTracked(sheetName, node.ref),
    }
  }

  const pair = (
    rangeNode: FormulaNode | undefined,
    criterionNode: FormulaNode | undefined,
  ): RuntimeDirectCriteriaDescriptor['criteriaPairs'][number] | undefined => {
    const range = resolveDirectCriteriaRange(rangeNode, args.ownerSheetName)
    const criterion = resolveCriterionOperand(criterionNode)
    if (!range || !criterion) {
      return undefined
    }
    return {
      range: withRangeRegion(range),
      criterion,
    }
  }

  const directCriteriaPairFromEquality = (node: FormulaNode): RuntimeDirectCriteriaDescriptor['criteriaPairs'][number] | undefined => {
    if (node.kind !== 'BinaryExpr' || node.operator !== '=') {
      return undefined
    }
    return pair(node.left, node.right) ?? pair(node.right, node.left)
  }

  const buildCriteriaPairsFromProduct = (
    node: FormulaNode,
  ): Array<RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]> | undefined => {
    const pairs: Array<RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]> = []
    for (const factor of flattenCriteriaProduct(node)) {
      const criteriaPair = directCriteriaPairFromEquality(factor)
      if (!criteriaPair) {
        return undefined
      }
      pairs.push(criteriaPair)
    }
    if (pairs.length === 0) {
      return undefined
    }
    const expectedLength = pairs[0]!.range.length
    return pairs.every((current) => current.range.length === expectedLength) ? pairs : undefined
  }

  const isStaticNumber = (node: FormulaNode | undefined, expected: number): boolean => {
    const value = staticCellValue(node)
    return value?.tag === ValueTag.Number && Object.is(value.value, expected)
  }

  const buildIndexMatchCriteriaDescriptor = (node: DirectCriteriaCallNode): RuntimeDirectCriteriaDescriptor | undefined => {
    const aggregateRange = resolveDirectCriteriaRange(node.args[0], args.ownerSheetName)
    const rowLookup = node.args[1]
    const columnLookup = node.args[2]
    if (
      !aggregateRange ||
      node.args.length < 2 ||
      node.args.length > 3 ||
      (columnLookup !== undefined && !isStaticNumber(columnLookup, 1)) ||
      rowLookup?.kind !== 'CallExpr' ||
      rowLookup.callee.trim().toUpperCase() !== 'MATCH' ||
      rowLookup.args.length !== 3 ||
      !isStaticNumber(rowLookup.args[0], 1) ||
      !isStaticNumber(rowLookup.args[2], 0)
    ) {
      return undefined
    }
    const criteriaPairs = buildCriteriaPairsFromProduct(rowLookup.args[1]!)
    if (!criteriaPairs || criteriaPairs.some((criteriaPair) => criteriaPair.range.length !== aggregateRange.length)) {
      return undefined
    }
    return {
      aggregateKind: 'first',
      aggregateRange: withRangeRegion(aggregateRange),
      criteriaPairs,
    }
  }

  const buildSimpleIndexMatchCriteriaDescriptor = (node: DirectCriteriaCallNode): RuntimeDirectCriteriaDescriptor | undefined => {
    const aggregateRange = resolveDirectCriteriaRange(node.args[0], args.ownerSheetName)
    const rowLookup = node.args[1]
    const columnLookup = node.args[2]
    if (
      !aggregateRange ||
      node.args.length < 2 ||
      node.args.length > 3 ||
      (columnLookup !== undefined && !isStaticNumber(columnLookup, 1)) ||
      rowLookup?.kind !== 'CallExpr' ||
      rowLookup.callee.trim().toUpperCase() !== 'MATCH' ||
      rowLookup.args.length !== 3 ||
      !isStaticNumber(rowLookup.args[2], 0)
    ) {
      return undefined
    }
    const criteriaPair = pair(rowLookup.args[1], rowLookup.args[0])
    if (!criteriaPair || criteriaPair.range.length !== aggregateRange.length) {
      return undefined
    }
    return {
      aggregateKind: 'first',
      aggregateRange: withRangeRegion(aggregateRange),
      firstMatchMode: 'exact-lookup',
      criteriaPairs: [criteriaPair],
    }
  }

  const buildIndexReferenceCriteriaDescriptor = (node: DirectCriteriaCallNode): RuntimeDirectCriteriaDescriptor | undefined => {
    if (node.args.length < 2 || node.args.length > 3) {
      return undefined
    }
    const rangeNode = node.args[0]
    if (!rangeNode || rangeNode.kind !== 'RangeRef' || rangeNode.refKind !== 'cells' || rangeNode.sheetEndName !== undefined) {
      return undefined
    }
    const parsedRange = parseRangeAddress(`${rangeNode.start}:${rangeNode.end}`, rangeNode.sheetName ?? args.ownerSheetName)
    if (parsedRange.kind !== 'cells') {
      return undefined
    }
    const sheetName = parsedRange.sheetName ?? rangeNode.sheetName ?? args.ownerSheetName
    if (!args.workbook.getSheet(sheetName)) {
      return undefined
    }
    const rowStart = parsedRange.start.row
    const rowEnd = parsedRange.end.row
    const colStart = parsedRange.start.col
    const colEnd = parsedRange.end.col
    const colCount = colEnd - colStart + 1
    const columnLookup = node.args[2]
    const columnValue = staticCellValue(columnLookup)
    const columnOffset =
      columnLookup === undefined
        ? colCount === 1
          ? 1
          : undefined
        : columnValue?.tag === ValueTag.Number
          ? Math.trunc(columnValue.value)
          : undefined
    if (columnOffset === undefined || columnOffset < 1 || columnOffset > colCount) {
      return undefined
    }
    const offsetOperand = resolveIndexOffsetOperand(node.args[1])
    if (!offsetOperand) {
      return undefined
    }
    return {
      aggregateKind: 'first',
      aggregateRange: withRangeRegion({
        sheetName,
        rowStart,
        rowEnd,
        col: colStart + columnOffset - 1,
        length: rowEnd - rowStart + 1,
      }),
      offsetOperand,
      criteriaPairs: [],
    }
  }

  const buildDescriptorForNode = (node: FormulaNode): RuntimeDirectCriteriaDescriptor | undefined => {
    if (node.kind !== 'CallExpr') {
      return undefined
    }
    const callee = node.callee.trim().toUpperCase()

    if (callee === 'IF') {
      const guard = resolveEmptyStringCellGuard(node.args[0])
      if (!guard || node.args.length < 3) {
        return undefined
      }
      const trueValue = staticCellValue(node.args[1])
      const falseValue = staticCellValue(node.args[2])
      if (guard.operator === '=' && trueValue) {
        return appendIfEmptyCellTransform(buildDescriptorForNode(node.args[2]!), guard.cellRef, trueValue)
      }
      if (guard.operator === '<>' && falseValue) {
        return appendIfEmptyCellTransform(buildDescriptorForNode(node.args[1]!), guard.cellRef, falseValue)
      }
      return undefined
    }

    if (callee === 'ROUND') {
      const digits = staticCellValue(node.args[1])
      if (node.args.length !== 2 || !digits) {
        return undefined
      }
      const inner = buildDescriptorForNode(node.args[0]!)
      return inner ? appendDirectCriteriaResultTransform(inner, { kind: 'round', digits }) : undefined
    }

    if (callee === 'IFERROR') {
      const fallback = staticCellValue(node.args[1])
      if (node.args.length !== 2 || !fallback) {
        return undefined
      }
      const inner = buildDescriptorForNode(node.args[0]!)
      return inner ? appendDirectCriteriaResultTransform(inner, { kind: 'if-error', fallback }) : undefined
    }

    if (callee === 'INDEX') {
      return (
        buildIndexReferenceCriteriaDescriptor(node) ??
        buildSimpleIndexMatchCriteriaDescriptor(node) ??
        buildIndexMatchCriteriaDescriptor(node)
      )
    }

    if (callee === 'COUNTIF') {
      const criteriaPair = pair(node.args[0], node.args[1])
      if (!criteriaPair) {
        return undefined
      }
      return {
        aggregateKind: 'count',
        aggregateRange: undefined,
        criteriaPairs: [criteriaPair],
      }
    }

    if (callee === 'COUNTIFS') {
      if (node.args.length === 0 || node.args.length % 2 !== 0) {
        return undefined
      }
      const criteriaPairs: Array<RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]> = []
      for (let index = 0; index < node.args.length; index += 2) {
        const criteriaPair = pair(node.args[index], node.args[index + 1])
        if (!criteriaPair) {
          return undefined
        }
        criteriaPairs.push(criteriaPair)
      }
      const expectedLength = criteriaPairs[0]!.range.length
      if (criteriaPairs.some((current) => current.range.length !== expectedLength)) {
        return undefined
      }
      return {
        aggregateKind: 'count',
        aggregateRange: undefined,
        criteriaPairs,
      }
    }

    if (callee === 'SUMIF' || callee === 'AVERAGEIF') {
      const criteriaPair = pair(node.args[0], node.args[1])
      if (!criteriaPair) {
        return undefined
      }
      const aggregateRange = resolveDirectCriteriaRange(node.args[2] ?? node.args[0], args.ownerSheetName)
      if (!aggregateRange || aggregateRange.length !== criteriaPair.range.length) {
        return undefined
      }
      return {
        aggregateKind: callee === 'SUMIF' ? 'sum' : 'average',
        aggregateRange: withRangeRegion(aggregateRange),
        criteriaPairs: [criteriaPair],
      }
    }

    if (callee !== 'SUMIFS' && callee !== 'AVERAGEIFS' && callee !== 'MINIFS' && callee !== 'MAXIFS') {
      return undefined
    }
    const aggregateRange = resolveDirectCriteriaRange(node.args[0], args.ownerSheetName)
    if (!aggregateRange || node.args.length < 3 || node.args.length % 2 === 0) {
      return undefined
    }
    const criteriaPairs: Array<RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]> = []
    for (let index = 1; index < node.args.length; index += 2) {
      const criteriaPair = pair(node.args[index], node.args[index + 1])
      if (!criteriaPair || criteriaPair.range.length !== aggregateRange.length) {
        return undefined
      }
      criteriaPairs.push(criteriaPair)
    }
    return {
      aggregateKind: callee === 'SUMIFS' ? 'sum' : callee === 'AVERAGEIFS' ? 'average' : callee === 'MINIFS' ? 'min' : 'max',
      aggregateRange: withRangeRegion(aggregateRange),
      criteriaPairs,
    }
  }

  if (args.compiled.astMatchesSource === false && mayContainDirectCriteriaPlanCall(args.compiled)) {
    return buildDescriptorForNode(parseFormula(args.source))
  }
  return buildDescriptorForNode(args.compiled.optimizedAst)
}

export function buildDirectAggregateDescriptor(args: {
  readonly compiled: ParsedCompiledFormula
  readonly ownerSheetName: string
  readonly regionGraph: RegionGraph
}): RuntimeDirectAggregateDescriptor | undefined {
  const buildDescriptor = (descriptor: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
    readonly aggregateKind: RuntimeDirectAggregateDescriptor['aggregateKind']
    readonly resultOffset?: number
  }): RuntimeDirectAggregateDescriptor => {
    const regions = internDirectAggregateRegions({
      regionGraph: args.regionGraph,
      sheetName: descriptor.sheetName,
      rowStart: descriptor.rowStart,
      rowEnd: descriptor.rowEnd,
      colStart: descriptor.colStart,
      colEnd: descriptor.colEnd,
    })
    return {
      ...regions,
      aggregateKind: descriptor.aggregateKind,
      sheetName: descriptor.sheetName,
      rowStart: descriptor.rowStart,
      rowEnd: descriptor.rowEnd,
      col: descriptor.colStart,
      colEnd: descriptor.colEnd,
      length: (descriptor.rowEnd - descriptor.rowStart + 1) * (descriptor.colEnd - descriptor.colStart + 1),
      ...(descriptor.resultOffset !== undefined ? { resultOffset: descriptor.resultOffset } : {}),
    }
  }
  const directAggregateCandidate = args.compiled.directAggregateCandidate
  if (directAggregateCandidate) {
    const rangeInfo = args.compiled.parsedSymbolicRanges?.[directAggregateCandidate.symbolicRangeIndex]
    if (rangeInfo && rangeInfo.refKind === 'cells' && (rangeInfo.sheetName === undefined || rangeInfo.sheetName === args.ownerSheetName)) {
      const sheetName = rangeInfo.sheetName ?? args.ownerSheetName
      return buildDescriptor({
        sheetName,
        rowStart: rangeInfo.startRow,
        rowEnd: rangeInfo.endRow,
        colStart: rangeInfo.startCol,
        colEnd: rangeInfo.endCol,
        aggregateKind: directAggregateCandidate.aggregateKind,
        ...(directAggregateCandidate.resultOffset !== undefined ? { resultOffset: directAggregateCandidate.resultOffset } : {}),
      })
    }
  }
  const node = args.compiled.optimizedAst
  if (node.kind !== 'CallExpr' || node.args.length !== 1) {
    return undefined
  }
  if (args.compiled.symbolicNames.length > 0 || args.compiled.symbolicTables.length > 0 || args.compiled.symbolicSpills.length > 0) {
    return undefined
  }
  const callee = node.callee.trim().toUpperCase()
  if (callee !== 'SUM' && callee !== 'AVERAGE' && callee !== 'AVG' && callee !== 'COUNT' && callee !== 'MIN' && callee !== 'MAX') {
    return undefined
  }
  const rangeNode = node.args[0]
  if (!rangeNode || rangeNode.kind !== 'RangeRef' || rangeNode.refKind !== 'cells' || rangeNode.sheetEndName !== undefined) {
    return undefined
  }
  if (rangeNode.sheetName !== undefined && rangeNode.sheetName !== args.ownerSheetName) {
    return undefined
  }
  const parsedRange = parseRangeAddress(`${rangeNode.start}:${rangeNode.end}`, rangeNode.sheetName ?? args.ownerSheetName)
  if (parsedRange.kind !== 'cells') {
    return undefined
  }
  const sheetName = rangeNode.sheetName ?? args.ownerSheetName
  return buildDescriptor({
    sheetName,
    rowStart: parsedRange.start.row,
    rowEnd: parsedRange.end.row,
    colStart: parsedRange.start.col,
    colEnd: parsedRange.end.col,
    aggregateKind:
      callee === 'SUM' ? 'sum' : callee === 'COUNT' ? 'count' : callee === 'MIN' ? 'min' : callee === 'MAX' ? 'max' : 'average',
  })
}

export function buildDirectLookupDescriptor(args: {
  readonly compiled: ParsedCompiledFormula
  readonly ownerSheetName: string
  readonly workbook: Pick<EngineRuntimeState, 'workbook'>['workbook']
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly exactLookup: Pick<ExactColumnIndexService, 'prepareVectorLookup'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'prepareVectorLookup'>
}): RuntimeDirectLookupDescriptor | undefined {
  const binding = resolveRuntimeDirectLookupBinding(args.compiled.jsPlan, args.ownerSheetName)
  if (!binding) {
    return undefined
  }
  const lookupSheet = args.workbook.getSheet(binding.lookupSheetName)
  if (!lookupSheet || !args.workbook.getSheet(binding.operandSheetName)) {
    return undefined
  }
  const operandCellIndex = args.ensureCellTracked(binding.operandSheetName, binding.operandAddress)
  if (binding.kind === 'exact') {
    const prepared = args.exactLookup.prepareVectorLookup({
      sheetName: binding.lookupSheetName,
      rowStart: binding.rowStart,
      rowEnd: binding.rowEnd,
      col: binding.col,
    })
    if (prepared.comparableKind === 'numeric' && prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
      return {
        kind: 'exact-uniform-numeric',
        operandCellIndex,
        sheetName: binding.lookupSheetName,
        sheetId: lookupSheet.id,
        rowStart: binding.rowStart,
        rowEnd: binding.rowEnd,
        col: binding.col,
        length: prepared.length,
        columnVersion: prepared.columnVersion,
        structureVersion: prepared.structureVersion,
        sheetColumnVersions: prepared.sheetColumnVersions,
        start: prepared.uniformStart,
        step: prepared.uniformStep,
        searchMode: binding.searchMode,
      }
    }
    return {
      kind: 'exact',
      operandCellIndex,
      prepared,
      searchMode: binding.searchMode,
    }
  }
  const prepared = args.sortedLookup.prepareVectorLookup({
    sheetName: binding.lookupSheetName,
    rowStart: binding.rowStart,
    rowEnd: binding.rowEnd,
    col: binding.col,
  })
  if (prepared.comparableKind === 'numeric' && prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
    return {
      kind: 'approximate-uniform-numeric',
      operandCellIndex,
      sheetName: binding.lookupSheetName,
      sheetId: lookupSheet.id,
      rowStart: binding.rowStart,
      rowEnd: binding.rowEnd,
      col: binding.col,
      length: prepared.length,
      columnVersion: prepared.columnVersion,
      structureVersion: prepared.structureVersion,
      sheetColumnVersions: prepared.sheetColumnVersions,
      start: prepared.uniformStart,
      step: prepared.uniformStep,
      matchMode: binding.matchMode,
    }
  }
  if (
    prepared.comparableKind === 'numeric' &&
    prepared.repeatedUniformStart !== undefined &&
    prepared.repeatedUniformStep !== undefined &&
    prepared.repeatedUniformRunLength !== undefined
  ) {
    return {
      kind: 'approximate-uniform-numeric',
      operandCellIndex,
      sheetName: binding.lookupSheetName,
      sheetId: lookupSheet.id,
      rowStart: binding.rowStart,
      rowEnd: binding.rowEnd,
      col: binding.col,
      length: prepared.length,
      columnVersion: prepared.columnVersion,
      structureVersion: prepared.structureVersion,
      sheetColumnVersions: prepared.sheetColumnVersions,
      start: prepared.repeatedUniformStart,
      step: prepared.repeatedUniformStep,
      repeatedRunLength: prepared.repeatedUniformRunLength,
      matchMode: binding.matchMode,
    }
  }
  return {
    kind: 'approximate',
    operandCellIndex,
    prepared,
    matchMode: binding.matchMode,
  }
}

export function hasLookupPlanInstruction(plan: readonly { opcode: string }[]): boolean {
  for (let index = 0; index < plan.length; index += 1) {
    const opcode = plan[index]?.opcode
    if (opcode === 'lookup-exact-match' || opcode === 'lookup-approximate-match') {
      return true
    }
  }
  return false
}
