import {
  type CompiledFormula,
  type DirectAggregateCandidate,
  type FormulaNode,
  type ParsedCellReferenceInfo,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
  type StructuralAxisTransform,
  parseRangeAddress,
} from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { mapStructuralAxisInterval } from '../../engine-structural-utils.js'
import { resolveRuntimeDirectLookupBinding } from '../direct-vector-lookup.js'
import type {
  EngineRuntimeState,
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectCriteriaResultTransform,
  RuntimeDirectLookupDescriptor,
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
    case 'CellRef':
    case 'CallExpr':
    case 'BinaryExpr':
    case 'ColumnRef':
    case 'InvokeExpr':
    case 'NameRef':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
      return undefined
  }
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
  if (!node || node.kind !== 'RangeRef' || node.refKind !== 'cells') {
    return undefined
  }
  const parsed = parseRangeAddress(`${node.start}:${node.end}`, node.sheetName ?? ownerSheetName)
  if (parsed.kind !== 'cells' || parsed.start.col !== parsed.end.col) {
    return undefined
  }
  return {
    sheetName: node.sheetName ?? ownerSheetName,
    rowStart: parsed.start.row,
    rowEnd: parsed.end.row,
    col: parsed.start.col,
    length: parsed.end.row - parsed.start.row + 1,
  }
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
  readonly ownerSheetName: string
  readonly workbook: Pick<EngineRuntimeState, 'workbook'>['workbook']
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly regionGraph: RegionGraph
}): RuntimeDirectCriteriaDescriptor | undefined {
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
      range: {
        ...range,
        regionId: args.regionGraph.internSingleColumnRegion({
          sheetName: range.sheetName,
          rowStart: range.rowStart,
          rowEnd: range.rowEnd,
          col: range.col,
        }),
      },
      criterion,
    }
  }

  const buildDescriptorForNode = (node: FormulaNode): RuntimeDirectCriteriaDescriptor | undefined => {
    if (node.kind !== 'CallExpr') {
      return undefined
    }
    const callee = node.callee.trim().toUpperCase()

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
        aggregateRange: {
          ...aggregateRange,
          regionId: args.regionGraph.internSingleColumnRegion({
            sheetName: aggregateRange.sheetName,
            rowStart: aggregateRange.rowStart,
            rowEnd: aggregateRange.rowEnd,
            col: aggregateRange.col,
          }),
        },
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
      aggregateRange: {
        ...aggregateRange,
        regionId: args.regionGraph.internSingleColumnRegion({
          sheetName: aggregateRange.sheetName,
          rowStart: aggregateRange.rowStart,
          rowEnd: aggregateRange.rowEnd,
          col: aggregateRange.col,
        }),
      },
      criteriaPairs,
    }
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
  if (!rangeNode || rangeNode.kind !== 'RangeRef' || rangeNode.refKind !== 'cells') {
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
