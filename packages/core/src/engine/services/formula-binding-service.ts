import { Effect } from 'effect'
import {
  compileFormulaAst,
  type CompiledFormula,
  type DirectAggregateCandidate,
  type FormulaNode,
  type Float64Arena,
  type ParsedCellReferenceInfo,
  type ParsedDependencyReference,
  type ParsedRangeReferenceInfo,
  parseCellAddress,
  parseRangeAddress,
  renameFormulaSheetReferences,
  serializeFormula,
  type Uint32Arena,
} from '@bilig/formula'
import { FormulaMode, ErrorCode, MAX_COLS, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import { resolveRuntimeDirectLookupBinding } from '../direct-vector-lookup.js'
import {
  entityPayload,
  isExactLookupColumnEntity,
  isRangeEntity,
  isSortedLookupColumnEntity,
  makeCellEntity,
  makeExactLookupColumnEntity,
  makeRangeEntity,
  makeSortedLookupColumnEntity,
} from '../../entity-ids.js'
import { growUint32 } from '../../engine-buffer-utils.js'
import { resolveMetadataReferencesInAst, spillDependencyKeyFromRef, tableDependencyKey } from '../../engine-metadata-utils.js'
import { errorValue } from '../../engine-value-utils.js'
import type { FormulaInstanceTable } from '../../formula/formula-instance-table.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import { retargetFormulaInstance } from '../../formula/structural-retargeting.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import {
  type EngineRuntimeState,
  type MaterializedDependencies,
  type RuntimeDirectAggregateDescriptor,
  type RuntimeDirectLookupDescriptor,
  type RuntimeDirectCriteriaDescriptor,
  type RuntimeFormula,
  UNRESOLVED_WASM_OPERAND,
  type U32,
} from '../runtime-state.js'
import { EngineFormulaBindingError } from '../errors.js'
import type { EngineCompiledPlanService } from './compiled-plan-service.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'

export interface EngineFormulaBindingService {
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => Effect.Effect<boolean, EngineFormulaBindingError>
  readonly clearFormula: (cellIndex: number) => Effect.Effect<boolean, EngineFormulaBindingError>
  readonly invalidateFormula: (cellIndex: number) => Effect.Effect<void, EngineFormulaBindingError>
  readonly rewriteCellFormulasForSheetRename: (
    oldSheetName: string,
    newSheetName: string,
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>
  readonly rebuildAllFormulaBindings: () => Effect.Effect<number[], EngineFormulaBindingError>
  readonly rebindFormulaCells: (
    candidates: readonly number[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>
  readonly rebindDefinedNameDependents: (
    names: readonly string[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>
  readonly rebindTableDependents: (
    tableNames: readonly string[],
    formulaChangedCount: number,
  ) => Effect.Effect<number, EngineFormulaBindingError>
  readonly rebindFormulasForSheet: (
    sheetName: string,
    formulaChangedCount: number,
    candidates?: readonly number[] | U32,
  ) => Effect.Effect<number, EngineFormulaBindingError>
  readonly bindFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly bindPreparedFormulaNow: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
  ) => boolean
  readonly rewriteFormulaCompiledPreservingBindingNow: (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
  ) => boolean
  readonly bindInitialFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => void
  readonly clearFormulaNow: (cellIndex: number) => boolean
  readonly invalidateFormulaNow: (cellIndex: number) => void
  readonly refreshRangeDependenciesNow: (rangeIndices: readonly number[]) => void
  readonly retargetRangeDependenciesNow: (transaction: StructuralTransaction, rangeIndices: readonly number[]) => void
  readonly rebindFormulaCellsNow: (candidates: readonly number[], formulaChangedCount: number) => number
  readonly rebindDefinedNameDependentsNow: (names: readonly string[], formulaChangedCount: number) => number
  readonly rebindTableDependentsNow: (tableNames: readonly string[], formulaChangedCount: number) => number
  readonly rebindFormulasForSheetNow: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
}

function formulaBindingErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function appendTrackedReverseEdge<Key extends string | number>(
  registry: Map<Key, Set<number>>,
  key: Key,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key)
  if (existing) {
    existing.add(dependentCellIndex)
    return
  }
  registry.set(key, new Set([dependentCellIndex]))
}

function removeTrackedReverseEdge<Key extends string | number>(
  registry: Map<Key, Set<number>>,
  key: Key,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key)
  if (!existing) {
    return
  }
  existing.delete(dependentCellIndex)
  if (existing.size === 0) {
    registry.delete(key)
  }
}

function collectTrackedDependents<Key extends string | number>(registry: Map<Key, Set<number>>, keys: readonly Key[]): number[] {
  const candidates = new Set<number>()
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex)
    })
  })
  return [...candidates]
}

function aggregateColumnDependencyKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}

function formulaColumnCountKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}

function directLookupColumnInfo(directLookup: RuntimeDirectLookupDescriptor): {
  sheetName: string
  col: number
  isExact: boolean
} {
  switch (directLookup.kind) {
    case 'exact':
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: true,
      }
    case 'exact-uniform-numeric':
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: true,
      }
    case 'approximate':
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: false,
      }
    case 'approximate-uniform-numeric':
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: false,
      }
  }
}

function directLookupStructureEqual(
  left: RuntimeDirectLookupDescriptor | undefined,
  right: RuntimeDirectLookupDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right || left.kind !== right.kind) {
    return false
  }
  switch (left.kind) {
    case 'exact':
      return (
        right.kind === 'exact' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.prepared.sheetName === right.prepared.sheetName &&
        left.prepared.rowStart === right.prepared.rowStart &&
        left.prepared.rowEnd === right.prepared.rowEnd &&
        left.prepared.col === right.prepared.col &&
        left.searchMode === right.searchMode
      )
    case 'exact-uniform-numeric':
      return (
        right.kind === 'exact-uniform-numeric' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.sheetName === right.sheetName &&
        left.rowStart === right.rowStart &&
        left.rowEnd === right.rowEnd &&
        left.col === right.col &&
        left.searchMode === right.searchMode
      )
    case 'approximate':
      return (
        right.kind === 'approximate' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.prepared.sheetName === right.prepared.sheetName &&
        left.prepared.rowStart === right.prepared.rowStart &&
        left.prepared.rowEnd === right.prepared.rowEnd &&
        left.prepared.col === right.prepared.col &&
        left.matchMode === right.matchMode
      )
    case 'approximate-uniform-numeric':
      return (
        right.kind === 'approximate-uniform-numeric' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.sheetName === right.sheetName &&
        left.rowStart === right.rowStart &&
        left.rowEnd === right.rowEnd &&
        left.col === right.col &&
        left.matchMode === right.matchMode
      )
  }
}

function directCriteriaOperandEqual(
  left: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined,
  right: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right || left.kind !== right.kind) {
    return false
  }
  if (left.kind === 'cell') {
    return right.kind === 'cell' && left.cellIndex === right.cellIndex
  }
  return right.kind === 'literal' && JSON.stringify(left.value) === JSON.stringify(right.value)
}

function directCriteriaStructureEqual(
  left: RuntimeDirectCriteriaDescriptor | undefined,
  right: RuntimeDirectCriteriaDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  if (left.aggregateKind !== right.aggregateKind) {
    return false
  }
  const leftRange = left.aggregateRange
  const rightRange = right.aggregateRange
  if (
    leftRange?.sheetName !== rightRange?.sheetName ||
    leftRange?.rowStart !== rightRange?.rowStart ||
    leftRange?.rowEnd !== rightRange?.rowEnd ||
    leftRange?.col !== rightRange?.col ||
    leftRange?.length !== rightRange?.length
  ) {
    return false
  }
  if (left.criteriaPairs.length !== right.criteriaPairs.length) {
    return false
  }
  for (let index = 0; index < left.criteriaPairs.length; index += 1) {
    const leftPair = left.criteriaPairs[index]!
    const rightPair = right.criteriaPairs[index]!
    if (
      leftPair.range.sheetName !== rightPair.range.sheetName ||
      leftPair.range.rowStart !== rightPair.range.rowStart ||
      leftPair.range.rowEnd !== rightPair.range.rowEnd ||
      leftPair.range.col !== rightPair.range.col ||
      leftPair.range.length !== rightPair.range.length ||
      !directCriteriaOperandEqual(leftPair.criterion, rightPair.criterion)
    ) {
      return false
    }
  }
  return true
}

function directAggregateStructureEqual(
  left: RuntimeDirectAggregateDescriptor | undefined,
  right: RuntimeDirectAggregateDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  return (
    left.aggregateKind === right.aggregateKind &&
    left.sheetName === right.sheetName &&
    left.rowStart === right.rowStart &&
    left.rowEnd === right.rowEnd &&
    left.col === right.col &&
    left.length === right.length
  )
}

function staticIntegerValue(node: FormulaNode | undefined): number | undefined {
  if (!node) {
    return undefined
  }
  if (node.kind === 'NumberLiteral') {
    return Number.isInteger(node.value) ? node.value : undefined
  }
  if (
    node.kind === 'UnaryExpr' &&
    node.operator === '-' &&
    node.argument.kind === 'NumberLiteral' &&
    Number.isInteger(node.argument.value)
  ) {
    return -node.argument.value
  }
  return undefined
}

function hasIndexedExactLookupCandidate(node: FormulaNode): boolean {
  return collectIndexedExactLookupCandidates(node).length > 0
}

function hasDirectApproximateLookupCandidate(node: FormulaNode): boolean {
  return collectDirectApproximateLookupCandidates(node).length > 0
}

interface IndexedExactLookupCandidate {
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

interface PreparedFormulaBindingShape {
  readonly dependencies: {
    readonly rangeDependencies: Uint32Array
  }
  readonly compiled: CompiledFormula
  readonly directLookup: RuntimeDirectLookupDescriptor | undefined
  readonly directAggregate: RuntimeDirectAggregateDescriptor | undefined
  readonly directCriteria: RuntimeDirectCriteriaDescriptor | undefined
}

function floatArrayEqual(left: ArrayLike<number>, right: ArrayLike<number>): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function hasStableSymbolicRangeLayout(
  current: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  next: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  rangeDependencyCount: number,
): boolean {
  if (current.symbolicRanges.length !== next.symbolicRanges.length) {
    return false
  }
  if (rangeDependencyCount !== current.symbolicRanges.length) {
    return false
  }
  if (current.parsedSymbolicRanges === undefined || next.parsedSymbolicRanges === undefined) {
    return current.parsedSymbolicRanges === next.parsedSymbolicRanges
  }
  if (current.parsedSymbolicRanges.length !== next.parsedSymbolicRanges.length) {
    return false
  }
  for (let index = 0; index < current.parsedSymbolicRanges.length; index += 1) {
    if (current.parsedSymbolicRanges[index]?.refKind !== next.parsedSymbolicRanges[index]?.refKind) {
      return false
    }
  }
  return true
}

function hasInPlaceDependencyRebindShape(existing: RuntimeFormula, prepared: PreparedFormulaBindingShape): boolean {
  return (
    uint32ArrayEqual(existing.rangeDependencies, prepared.dependencies.rangeDependencies) &&
    hasStableSymbolicRangeLayout(existing.compiled, prepared.compiled, existing.rangeDependencies.length) &&
    existing.compiled.symbolicNames.length === 0 &&
    prepared.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    prepared.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    prepared.compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    prepared.directLookup === undefined &&
    existing.directAggregate === undefined &&
    prepared.directAggregate === undefined &&
    existing.directCriteria === undefined &&
    prepared.directCriteria === undefined
  )
}

function canRewriteCompiledPreservingBindings(existing: RuntimeFormula, compiled: CompiledFormula): boolean {
  return (
    hasStableSymbolicRangeLayout(existing.compiled, compiled, existing.rangeDependencies.length) &&
    existing.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    existing.directAggregate === undefined &&
    existing.directCriteria === undefined &&
    uint32ArrayEqual(existing.compiled.program, compiled.program) &&
    floatArrayEqual(existing.compiled.constants, compiled.constants) &&
    existing.compiled.mode === compiled.mode
  )
}

function canRewriteCompiledPreservingDirectAggregate(existing: RuntimeFormula, compiled: CompiledFormula): boolean {
  return (
    existing.rangeDependencies.length === 0 &&
    existing.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    existing.directAggregate !== undefined &&
    existing.directCriteria === undefined &&
    existing.programLength === compiled.program.length &&
    existing.constNumberLength === compiled.constants.length &&
    existing.compiled.mode === compiled.mode
  )
}

function collectIndexedExactLookupCandidates(node: FormulaNode): IndexedExactLookupCandidate[] {
  switch (node.kind) {
    case 'CallExpr': {
      const callee = node.callee.trim().toUpperCase()
      const lookupRange = node.args[1]
      if (lookupRange?.kind === 'RangeRef' && lookupRange.refKind === 'cells' && lookupRange.start !== lookupRange.end) {
        const isIndexedLookupCall =
          (callee === 'MATCH' && node.args.length === 3 && staticIntegerValue(node.args[2]) === 0) ||
          (callee === 'XMATCH' &&
            node.args.length >= 2 &&
            node.args.length <= 4 &&
            (node.args.length === 2 || staticIntegerValue(node.args[2]) === 0) &&
            (node.args.length < 4 || staticIntegerValue(node.args[3]) === 1 || staticIntegerValue(node.args[3]) === -1))
        if (isIndexedLookupCall) {
          const parsedRange = parseRangeAddress(`${lookupRange.start}:${lookupRange.end}`, lookupRange.sheetName)
          if (parsedRange.kind === 'cells') {
            return [
              {
                ...(lookupRange.sheetName === undefined ? {} : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectIndexedExactLookupCandidates),
            ]
          }
        }
      }
      return node.args.flatMap(collectIndexedExactLookupCandidates)
    }
    case 'UnaryExpr':
      return collectIndexedExactLookupCandidates(node.argument)
    case 'BinaryExpr':
      return [...collectIndexedExactLookupCandidates(node.left), ...collectIndexedExactLookupCandidates(node.right)]
    case 'InvokeExpr':
      return [...collectIndexedExactLookupCandidates(node.callee), ...node.args.flatMap(collectIndexedExactLookupCandidates)]
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return []
  }
}

interface DirectApproximateLookupCandidate {
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

type ParsedCompiledFormula = CompiledFormula & {
  parsedDeps?: ParsedDependencyReference[]
  parsedSymbolicRefs?: ParsedCellReferenceInfo[]
  parsedSymbolicRanges?: ParsedRangeReferenceInfo[]
  directAggregateCandidate?: DirectAggregateCandidate
}

function uint32ArrayEqual(left: Uint32Array | readonly number[], right: Uint32Array | readonly number[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function stringArrayEqual(left: readonly string[], right: readonly string[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function collectDirectApproximateLookupCandidates(node: FormulaNode): DirectApproximateLookupCandidate[] {
  switch (node.kind) {
    case 'CallExpr': {
      const callee = node.callee.trim().toUpperCase()
      const lookupRange = node.args[1]
      if (lookupRange?.kind === 'RangeRef' && lookupRange.refKind === 'cells' && lookupRange.start !== lookupRange.end) {
        const matchMode = staticIntegerValue(node.args[2])
        const searchMode = node.args.length >= 4 ? staticIntegerValue(node.args[3]) : 1
        const isDirectApproximateLookupCall =
          (callee === 'MATCH' && node.args.length === 3 && (matchMode === 1 || matchMode === -1)) ||
          (callee === 'XMATCH' &&
            node.args.length >= 3 &&
            node.args.length <= 4 &&
            (matchMode === 1 || matchMode === -1) &&
            searchMode === 1)
        if (isDirectApproximateLookupCall) {
          const parsedRange = parseRangeAddress(`${lookupRange.start}:${lookupRange.end}`, lookupRange.sheetName)
          if (parsedRange.kind === 'cells') {
            return [
              {
                ...(lookupRange.sheetName === undefined ? {} : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectDirectApproximateLookupCandidates),
            ]
          }
        }
      }
      return node.args.flatMap(collectDirectApproximateLookupCandidates)
    }
    case 'UnaryExpr':
      return collectDirectApproximateLookupCandidates(node.argument)
    case 'BinaryExpr':
      return [...collectDirectApproximateLookupCandidates(node.left), ...collectDirectApproximateLookupCandidates(node.right)]
    case 'InvokeExpr':
      return [...collectDirectApproximateLookupCandidates(node.callee), ...node.args.flatMap(collectDirectApproximateLookupCandidates)]
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return []
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

function buildDirectCriteriaDescriptor(args: {
  readonly compiled: ParsedCompiledFormula
  readonly ownerSheetName: string
  readonly workbook: Pick<EngineRuntimeState, 'workbook'>['workbook']
  readonly ensureCellTracked: (sheetName: string, address: string) => number
}): RuntimeDirectCriteriaDescriptor | undefined {
  const node = args.compiled.optimizedAst
  if (node.kind !== 'CallExpr') {
    return undefined
  }
  const callee = node.callee.trim().toUpperCase()

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
    return { range, criterion }
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
      aggregateRange,
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
    aggregateRange,
    criteriaPairs,
  }
}

function buildDirectAggregateDescriptor(args: {
  readonly compiled: ParsedCompiledFormula
  readonly ownerSheetName: string
}): RuntimeDirectAggregateDescriptor | undefined {
  const directAggregateCandidate = args.compiled.directAggregateCandidate
  if (directAggregateCandidate) {
    const rangeInfo = args.compiled.parsedSymbolicRanges?.[directAggregateCandidate.symbolicRangeIndex]
    if (
      rangeInfo &&
      rangeInfo.refKind === 'cells' &&
      (rangeInfo.sheetName === undefined || rangeInfo.sheetName === args.ownerSheetName) &&
      rangeInfo.startCol === rangeInfo.endCol
    ) {
      return {
        aggregateKind: directAggregateCandidate.aggregateKind,
        sheetName: rangeInfo.sheetName ?? args.ownerSheetName,
        rowStart: rangeInfo.startRow,
        rowEnd: rangeInfo.endRow,
        col: rangeInfo.startCol,
        length: rangeInfo.endRow - rangeInfo.startRow + 1,
      }
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
  const range = resolveDirectCriteriaRange(rangeNode, args.ownerSheetName)
  if (!range) {
    return undefined
  }
  return {
    aggregateKind:
      callee === 'SUM' ? 'sum' : callee === 'COUNT' ? 'count' : callee === 'MIN' ? 'min' : callee === 'MAX' ? 'max' : 'average',
    sheetName: range.sheetName,
    rowStart: range.rowStart,
    rowEnd: range.rowEnd,
    col: range.col,
    length: range.length,
  }
}

function buildDirectLookupDescriptor(args: {
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
  if (!args.workbook.getSheet(binding.lookupSheetName) || !args.workbook.getSheet(binding.operandSheetName)) {
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

function hasLookupPlanInstruction(plan: readonly { opcode: string }[]): boolean {
  for (let index = 0; index < plan.length; index += 1) {
    const opcode = plan[index]?.opcode
    if (opcode === 'lookup-exact-match' || opcode === 'lookup-approximate-match') {
      return true
    }
  }
  return false
}

const PUSH_CELL_OPCODE = Number(Opcode.PushCell)
const PUSH_RANGE_OPCODE = Number(Opcode.PushRange)
const PUSH_STRING_OPCODE = Number(Opcode.PushString)

export function createEngineFormulaBindingService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'ranges' | 'getUseColumnIndex'> & {
    counters?: EngineCounters
  }
  readonly compiledPlans: EngineCompiledPlanService
  readonly formulaInstances: FormulaInstanceTable
  readonly resolveTemplateForCell: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly exactLookup: Pick<ExactColumnIndexService, 'primeColumnIndex' | 'prepareVectorLookup'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'primeColumnIndex' | 'prepareVectorLookup'>
  readonly edgeArena: EdgeArena
  readonly programArena: Uint32Arena
  readonly constantArena: Float64Arena
  readonly rangeListArena: Uint32Arena
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>
    reverseRangeEdges: Array<EdgeSlice | undefined>
    reverseDefinedNameEdges: Map<string, Set<number>>
    reverseTableEdges: Map<string, Set<number>>
    reverseSpillEdges: Map<string, Set<number>>
    reverseAggregateColumnEdges: Map<number, Set<number>>
    reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly forEachSheetCell: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => void
  readonly scheduleWasmProgramSync: () => void
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly resolveStructuredReference: (tableName: string, columnName: string) => FormulaNode | undefined
  readonly resolveSpillReference: (currentSheetName: string, sheetName: string | undefined, address: string) => FormulaNode | undefined
  readonly getDependencyBuildEpoch: () => number
  readonly setDependencyBuildEpoch: (next: number) => void
  readonly getDependencyBuildSeen: () => U32
  readonly setDependencyBuildSeen: (next: U32) => void
  readonly getDependencyBuildCells: () => U32
  readonly setDependencyBuildCells: (next: U32) => void
  readonly getDependencyBuildEntities: () => U32
  readonly setDependencyBuildEntities: (next: U32) => void
  readonly getDependencyBuildRanges: () => U32
  readonly setDependencyBuildRanges: (next: U32) => void
  readonly getDependencyBuildNewRanges: () => U32
  readonly setDependencyBuildNewRanges: (next: U32) => void
  readonly getSymbolicRefBindings: () => U32
  readonly setSymbolicRefBindings: (next: U32) => void
  readonly getSymbolicRangeBindings: () => U32
  readonly setSymbolicRangeBindings: (next: U32) => void
}): EngineFormulaBindingService {
  const resolvedCompiledCache = new Map<string, ParsedCompiledFormula>()
  const formulaColumnCounts = new Map<number, number>()

  const normalizeLookupCompileMode = (compiled: ParsedCompiledFormula): ParsedCompiledFormula => {
    if (compiled.mode !== FormulaMode.WasmFastPath) {
      return compiled
    }
    if (!hasIndexedExactLookupCandidate(compiled.optimizedAst) && !hasDirectApproximateLookupCandidate(compiled.optimizedAst)) {
      return compiled
    }
    return {
      ...compiled,
      mode: FormulaMode.JsOnly,
    }
  }

  const ensureDependencyBuildCapacity = (
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity = 0,
    symbolicRangeCapacity = 0,
  ): void => {
    if (cellCapacity > args.getDependencyBuildSeen().length) {
      args.setDependencyBuildSeen(growUint32(args.getDependencyBuildSeen(), cellCapacity))
    }
    if (cellCapacity > args.getDependencyBuildCells().length) {
      args.setDependencyBuildCells(growUint32(args.getDependencyBuildCells(), cellCapacity))
    }
    if (dependencyCapacity > args.getDependencyBuildEntities().length) {
      args.setDependencyBuildEntities(growUint32(args.getDependencyBuildEntities(), dependencyCapacity))
    }
    if (dependencyCapacity > args.getDependencyBuildRanges().length) {
      args.setDependencyBuildRanges(growUint32(args.getDependencyBuildRanges(), dependencyCapacity))
    }
    if (dependencyCapacity > args.getDependencyBuildNewRanges().length) {
      args.setDependencyBuildNewRanges(growUint32(args.getDependencyBuildNewRanges(), dependencyCapacity))
    }
    if (symbolicRefCapacity > args.getSymbolicRefBindings().length) {
      args.setSymbolicRefBindings(growUint32(args.getSymbolicRefBindings(), symbolicRefCapacity))
    }
    if (symbolicRangeCapacity > args.getSymbolicRangeBindings().length) {
      args.setSymbolicRangeBindings(growUint32(args.getSymbolicRangeBindings(), symbolicRangeCapacity))
    }
  }

  const setReverseEdgeSlice = (entityId: number, slice: EdgeSlice): void => {
    const empty = slice.ptr < 0 || slice.len === 0
    if (isRangeEntity(entityId)) {
      args.reverseState.reverseRangeEdges[entityPayload(entityId)] = empty ? undefined : slice
      return
    }
    if (isExactLookupColumnEntity(entityId)) {
      if (empty) {
        args.reverseState.reverseExactLookupColumnEdges.delete(entityPayload(entityId))
      } else {
        args.reverseState.reverseExactLookupColumnEdges.set(entityPayload(entityId), slice)
      }
      return
    }
    if (isSortedLookupColumnEntity(entityId)) {
      if (empty) {
        args.reverseState.reverseSortedLookupColumnEdges.delete(entityPayload(entityId))
      } else {
        args.reverseState.reverseSortedLookupColumnEdges.set(entityPayload(entityId), slice)
      }
      return
    }
    args.reverseState.reverseCellEdges[entityPayload(entityId)] = empty ? undefined : slice
  }

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)]
    }
    if (isExactLookupColumnEntity(entityId)) {
      return args.reverseState.reverseExactLookupColumnEdges.get(entityPayload(entityId))
    }
    if (isSortedLookupColumnEntity(entityId)) {
      return args.reverseState.reverseSortedLookupColumnEdges.get(entityPayload(entityId))
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)]
  }

  const appendReverseEdge = (entityId: number, dependentEntityId: number): void => {
    const slice = getReverseEdgeSlice(entityId) ?? args.edgeArena.empty()
    setReverseEdgeSlice(entityId, args.edgeArena.appendUnique(slice, dependentEntityId))
  }

  const removeReverseEdge = (entityId: number, dependentEntityId: number): void => {
    const slice = getReverseEdgeSlice(entityId)
    if (!slice) {
      return
    }
    setReverseEdgeSlice(entityId, args.edgeArena.removeValue(slice, dependentEntityId))
  }

  const refreshRangeDependenciesNow = (rangeIndices: readonly number[]): void => {
    const refreshed = new Set<number>()
    const materializer = {
      ensureCell: (sheetId: number, row: number, col: number) => args.ensureCellTrackedByCoords(sheetId, row, col),
      forEachSheetCell: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => args.forEachSheetCell(sheetId, fn),
      isFormulaCell: (cellIndex: number) => (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0,
    }
    rangeIndices.forEach((rangeIndex) => {
      if (refreshed.has(rangeIndex)) {
        return
      }
      refreshed.add(rangeIndex)
      syncRangeDependencyEdges(rangeIndex, args.state.ranges.refresh(rangeIndex, materializer))
    })
  }

  const retargetRangeDependenciesNow = (transaction: StructuralTransaction, rangeIndices: readonly number[]): void => {
    const materializer = {
      ensureCell: (sheetId: number, row: number, col: number) => args.ensureCellTrackedByCoords(sheetId, row, col),
      forEachSheetCell: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => args.forEachSheetCell(sheetId, fn),
      isFormulaCell: (cellIndex: number) => (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0,
    }
    const touched = args.state.ranges.applyStructuralTransaction(transaction, rangeIndices, materializer)
    touched.forEach(({ rangeIndex, oldDependencySources, newDependencySources }) => {
      syncRangeDependencyEdges(rangeIndex, { oldDependencySources, newDependencySources })
    })
  }

  const syncRangeDependencyEdges = (
    rangeIndex: number,
    deps: { oldDependencySources: Uint32Array; newDependencySources: Uint32Array },
  ): void => {
    const rangeEntity = makeRangeEntity(rangeIndex)
    const nextSources = new Set<number>(deps.newDependencySources)
    deps.oldDependencySources.forEach((dependencyEntity) => {
      if (!nextSources.has(dependencyEntity)) {
        removeReverseEdge(dependencyEntity, rangeEntity)
      }
    })
    const priorSources = new Set<number>(deps.oldDependencySources)
    deps.newDependencySources.forEach((dependencyEntity) => {
      if (!priorSources.has(dependencyEntity)) {
        appendReverseEdge(dependencyEntity, rangeEntity)
      }
    })
  }

  const rangeDependenciesHaveNoFormulaMembers = (rangeDependencies: Uint32Array): boolean =>
    rangeDependencies.every((rangeIndex) => args.state.ranges.getFormulaMembersView(rangeIndex).length === 0)

  const appendDefinedNameReverseEdge = (name: string, dependentCellIndex: number): void => {
    appendTrackedReverseEdge(args.reverseState.reverseDefinedNameEdges, normalizeDefinedName(name), dependentCellIndex)
  }

  const removeDefinedNameReverseEdge = (name: string, dependentCellIndex: number): void => {
    removeTrackedReverseEdge(args.reverseState.reverseDefinedNameEdges, normalizeDefinedName(name), dependentCellIndex)
  }

  const pruneTrackedDependencyCell = (cellIndex: number, ownerCellIndex: number): void => {
    if (cellIndex === ownerCellIndex) {
      return
    }
    if (getReverseEdgeSlice(makeCellEntity(cellIndex))) {
      return
    }
    if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0) {
      return
    }
    args.state.workbook.pruneCellIfEmpty(cellIndex)
  }

  const pruneOrphanedDependencyCells = (cellIndices: readonly number[]): void => {
    cellIndices.forEach((cellIndex) => {
      if (getReverseEdgeSlice(makeCellEntity(cellIndex))) {
        return
      }
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0) {
        return
      }
      args.state.workbook.pruneCellIfEmpty(cellIndex)
    })
  }

  const updateFormulaDependenciesInPlaceNow = (
    cellIndex: number,
    existing: RuntimeFormula,
    prepared: ReturnType<typeof prepareFormulaBindingFromCompiledNow>,
    ownerSheetName: string,
    source: string,
  ): void => {
    const formulaEntity = makeCellEntity(cellIndex)
    const previousDependencies = args.edgeArena.read(existing.dependencyEntities)
    const nextDependencies = prepared.dependencies.dependencyEntities
    const nextDependencySet = new Set<number>(nextDependencies)
    const previousDependencySet = new Set<number>(previousDependencies)

    previousDependencies.forEach((dependencyEntity) => {
      if (nextDependencySet.has(dependencyEntity)) {
        return
      }
      removeReverseEdge(dependencyEntity, formulaEntity)
      if (!isRangeEntity(dependencyEntity)) {
        pruneTrackedDependencyCell(entityPayload(dependencyEntity), cellIndex)
      }
    })
    nextDependencies.forEach((dependencyEntity) => {
      if (previousDependencySet.has(dependencyEntity)) {
        return
      }
      appendReverseEdge(dependencyEntity, formulaEntity)
    })

    const plan = args.compiledPlans.replace(existing.planId, source, prepared.plan.compiled, prepared.templateId)
    args.compiledPlans.release(prepared.plan.id)
    existing.source = source
    existing.planId = plan.id
    existing.templateId = prepared.templateId
    existing.compiled = plan.compiled
    existing.plan = plan
    existing.dependencyIndices = prepared.dependencies.dependencyIndices
    existing.dependencyEntities = args.edgeArena.replace(existing.dependencyEntities, nextDependencies)
    existing.runtimeProgram = prepared.runtimeProgram
    existing.constants = prepared.compiled.constants
    existing.programLength = prepared.runtimeProgram.length
    existing.constNumberLength = prepared.compiled.constants.length
    existing.directLookup = undefined
    existing.directAggregate = undefined
    existing.directCriteria = undefined
    args.state.workbook.cellStore.flags[cellIndex] =
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
    if (existing.compiled.mode === FormulaMode.JsOnly) {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
    } else {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
    }
    args.scheduleWasmProgramSync()
    recordFormulaInstanceNow(cellIndex, source, prepared.templateId)

    primeLookupCandidatesNow(ownerSheetName, undefined, prepared.indexedExactLookupCandidates, prepared.directApproximateLookupCandidates)
  }

  const rewriteFormulaCompiledPreservingBindingNow = (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (!existing) {
      return false
    }
    const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    if (!ownerSheetName) {
      return false
    }
    let nextDirectAggregate: RuntimeDirectAggregateDescriptor | undefined
    if (canRewriteCompiledPreservingBindings(existing, compiled) && rangeDependenciesHaveNoFormulaMembers(existing.rangeDependencies)) {
      nextDirectAggregate = undefined
    } else if (canRewriteCompiledPreservingDirectAggregate(existing, compiled)) {
      nextDirectAggregate = buildDirectAggregateDescriptor({
        compiled: compiled as ParsedCompiledFormula,
        ownerSheetName,
      })
      if (!nextDirectAggregate) {
        return false
      }
    } else {
      return false
    }
    const nextTemplateId = templateId ?? existing.templateId
    const plan = args.compiledPlans.replace(existing.planId, source, compiled, nextTemplateId)
    const previousDirectAggregate = existing.directAggregate
    existing.source = source
    existing.planId = plan.id
    existing.templateId = nextTemplateId
    existing.compiled = plan.compiled
    existing.plan = plan
    existing.constants = compiled.constants
    existing.programLength = compiled.program.length
    existing.constNumberLength = compiled.constants.length
    existing.directAggregate = nextDirectAggregate
    if (previousDirectAggregate || nextDirectAggregate) {
      const previousSheet = previousDirectAggregate ? args.state.workbook.getSheet(previousDirectAggregate.sheetName) : undefined
      if (previousDirectAggregate && previousSheet) {
        removeTrackedReverseEdge(
          args.reverseState.reverseAggregateColumnEdges,
          aggregateColumnDependencyKey(previousSheet.id, previousDirectAggregate.col),
          cellIndex,
        )
      }
      const nextSheet = nextDirectAggregate ? args.state.workbook.getSheet(nextDirectAggregate.sheetName) : undefined
      if (nextDirectAggregate && nextSheet) {
        appendTrackedReverseEdge(
          args.reverseState.reverseAggregateColumnEdges,
          aggregateColumnDependencyKey(nextSheet.id, nextDirectAggregate.col),
          cellIndex,
        )
      }
    }
    args.state.workbook.cellStore.flags[cellIndex] =
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
    if (compiled.mode === FormulaMode.JsOnly) {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
    } else {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
    }
    recordFormulaInstanceNow(cellIndex, source, nextTemplateId)
    return true
  }

  const isCellIndexMappedNow = (cellIndex: number): boolean => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const row = args.state.workbook.cellStore.rows[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || row === undefined || col === undefined) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    return sheet?.grid.get(row, col) === cellIndex
  }

  const compileFormulaForCell = (
    cellIndex: number,
    currentSheetName: string,
    source: string,
  ): { compiled: ParsedCompiledFormula; templateResolution: FormulaTemplateResolution } => {
    const row = args.state.workbook.cellStore.rows[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (row === undefined || col === undefined) {
      throw new Error(`Cannot resolve formula template without coordinates for cell ${cellIndex}`)
    }
    const templateResolution = args.resolveTemplateForCell(source, row, col)
    const compiled = normalizeLookupCompileMode(templateResolution.compiled as ParsedCompiledFormula)
    if (compiled.symbolicNames.length === 0 && compiled.symbolicTables.length === 0 && compiled.symbolicSpills.length === 0) {
      return {
        compiled,
        templateResolution,
      }
    }

    const resolved = resolveMetadataReferencesInAst(compiled.ast, {
      resolveName: (name) => args.state.workbook.getDefinedName(name)?.value,
      resolveStructuredReference: (tableName, columnName) => args.resolveStructuredReference(tableName, columnName),
      resolveSpillReference: (sheetName, address) => args.resolveSpillReference(currentSheetName, sheetName, address),
    })
    if (!resolved.substituted || !resolved.fullyResolved) {
      return {
        compiled,
        templateResolution,
      }
    }

    const resolvedCacheKey = `${currentSheetName}\u0000${source}\u0000${serializeFormula(resolved.node)}`
    let resolvedCompiled = resolvedCompiledCache.get(resolvedCacheKey)
    if (!resolvedCompiled) {
      resolvedCompiled = normalizeLookupCompileMode(
        compileFormulaAst(source, resolved.node, {
          originalAst: compiled.ast,
          symbolicNames: compiled.symbolicNames,
          symbolicTables: compiled.symbolicTables,
          symbolicSpills: compiled.symbolicSpills,
        }) as ParsedCompiledFormula,
      )
      resolvedCompiledCache.set(resolvedCacheKey, resolvedCompiled)
    }
    return {
      compiled: resolvedCompiled,
      templateResolution,
    }
  }

  const recordFormulaInstanceNow = (cellIndex: number, source: string, templateId: number | undefined): void => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const row = args.state.workbook.cellStore.rows[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || row === undefined || col === undefined) {
      args.formulaInstances.delete(cellIndex)
      return
    }
    const sheetName = args.state.workbook.getSheetNameById(sheetId)
    if (!sheetName) {
      args.formulaInstances.delete(cellIndex)
      return
    }
    const existing = args.formulaInstances.get(cellIndex)
    if (existing) {
      args.formulaInstances.upsert(
        retargetFormulaInstance(existing, {
          sheetName,
          row,
          col,
          source,
          ...(templateId !== undefined ? { templateId } : {}),
        }),
      )
      return
    }
    args.formulaInstances.upsert({
      cellIndex,
      sheetName,
      row,
      col,
      source,
      ...(templateId !== undefined ? { templateId } : {}),
    })
  }

  const materializeDependencies = (
    currentSheetName: string,
    compiled: ParsedCompiledFormula,
    directAggregate: RuntimeDirectAggregateDescriptor | undefined,
    directLookupBinding: ReturnType<typeof resolveRuntimeDirectLookupBinding> | undefined,
  ): MaterializedDependencies => {
    const currentSheetId = args.state.workbook.getSheet(currentSheetName)?.id
    const deps = compiled.deps
    const parsedCellDeps = compiled.parsedDeps
    if (
      compiled.symbolicRanges.length === 0 &&
      parsedCellDeps !== undefined &&
      parsedCellDeps.length === deps.length &&
      parsedCellDeps.length > 0 &&
      parsedCellDeps.length <= 2 &&
      parsedCellDeps.every((dependency) => dependency?.kind === 'cell')
    ) {
      ensureDependencyBuildCapacity(args.state.workbook.cellStore.size + 1, parsedCellDeps.length + 1, compiled.symbolicRefs.length + 1, 1)
      let dependencyIndexCount = 0
      let dependencyEntityCount = 0
      for (let depIndex = 0; depIndex < parsedCellDeps.length; depIndex += 1) {
        const parsedDep = parsedCellDeps[depIndex]!
        if (parsedDep.sheetName && !args.state.workbook.getSheet(parsedDep.sheetName)) {
          continue
        }
        const cellIndex =
          parsedDep.sheetName === undefined && parsedDep.row !== undefined && parsedDep.col !== undefined && currentSheetId !== undefined
            ? args.ensureCellTrackedByCoords(currentSheetId, parsedDep.row, parsedDep.col)
            : args.ensureCellTracked(parsedDep.sheetName ?? currentSheetName, parsedDep.address)
        let seen = false
        for (let existingIndex = 0; existingIndex < dependencyIndexCount; existingIndex += 1) {
          if (args.getDependencyBuildCells()[existingIndex] === cellIndex) {
            seen = true
            break
          }
        }
        if (!seen) {
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
          dependencyIndexCount += 1
        }
        args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
        dependencyEntityCount += 1
      }
      return {
        dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
        dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
        rangeDependencies: args.getDependencyBuildRanges().slice(0, 0),
        symbolicRangeIndices: args.getSymbolicRangeBindings(),
        symbolicRangeCount: 0,
        newRangeIndices: args.getDependencyBuildNewRanges(),
        newRangeCount: 0,
      }
    }

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    )
    let epoch = args.getDependencyBuildEpoch() + 1
    if (epoch === 0xffff_ffff) {
      epoch = 1
      args.getDependencyBuildSeen().fill(0)
    }
    args.setDependencyBuildEpoch(epoch)

    let dependencyIndexCount = 0
    let dependencyEntityCount = 0
    let rangeDependencyCount = 0
    let newRangeCount = 0
    const symbolicRangeIndexByAddress =
      compiled.symbolicRanges.length > 0 ? new Map(compiled.symbolicRanges.map((range, index) => [range, index])) : undefined
    args.getSymbolicRangeBindings().fill(UNRESOLVED_WASM_OPERAND, 0, compiled.symbolicRanges.length)

    for (let depIndex = 0; depIndex < deps.length; depIndex += 1) {
      const dep = deps[depIndex]!
      const parsedDep = compiled.parsedDeps?.[depIndex]
      if (parsedDep?.kind === 'cell') {
        if (parsedDep.sheetName && !args.state.workbook.getSheet(parsedDep.sheetName)) {
          continue
        }
        const cellIndex =
          parsedDep.sheetName === undefined && parsedDep.row !== undefined && parsedDep.col !== undefined && currentSheetId !== undefined
            ? args.ensureCellTrackedByCoords(currentSheetId, parsedDep.row, parsedDep.col)
            : args.ensureCellTracked(parsedDep.sheetName ?? currentSheetName, parsedDep.address)
        if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
          args.getDependencyBuildSeen()[cellIndex] = epoch
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
          dependencyIndexCount += 1
        }
        args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
        dependencyEntityCount += 1
        continue
      }
      if (dep.includes(':')) {
        const parsedRangeDep =
          parsedDep?.kind === 'range' ? parsedDep : compiled.parsedSymbolicRanges?.find((range) => range.address === dep)
        const range =
          parsedRangeDep === undefined
            ? parseRangeAddress(dep, currentSheetName)
            : parseRangeAddress(`${parsedRangeDep.startAddress}:${parsedRangeDep.endAddress}`, parsedRangeDep.sheetName ?? currentSheetName)
        const sheetName = range.sheetName ?? currentSheetName
        const isDirectLookupColumn =
          directLookupBinding !== undefined &&
          range.kind === 'cells' &&
          range.start.col === range.end.col &&
          sheetName === directLookupBinding.lookupSheetName &&
          range.start.col === directLookupBinding.col &&
          range.start.row === directLookupBinding.rowStart &&
          range.end.row === directLookupBinding.rowEnd
        if (isDirectLookupColumn) {
          const sheet = args.state.workbook.getSheet(sheetName)
          if (sheet) {
            for (let row = range.start.row; row <= range.end.row; row += 1) {
              const cellIndex = sheet.grid.get(row, range.start.col)
              if (cellIndex === -1) {
                continue
              }
              if ((args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) === 0) {
                continue
              }
              if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
                args.getDependencyBuildSeen()[cellIndex] = epoch
                args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
                dependencyIndexCount += 1
              }
              args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
              dependencyEntityCount += 1
            }
          }
          continue
        }
        const symbolicRangeIndex = symbolicRangeIndexByAddress?.get(dep) ?? -1
        if (range.sheetName && !args.state.workbook.getSheet(sheetName)) {
          continue
        }
        const sheet = args.state.workbook.getSheet(sheetName)
        if (!sheet) {
          continue
        }
        const compactDirectAggregateRange =
          directAggregate !== undefined &&
          range.kind === 'cells' &&
          range.start.col === directAggregate.col &&
          range.end.col === directAggregate.col &&
          range.start.row === directAggregate.rowStart &&
          range.end.row === directAggregate.rowEnd &&
          sheetName === directAggregate.sheetName
        if (compactDirectAggregateRange) {
          if ((formulaColumnCounts.get(formulaColumnCountKey(sheet.id, range.start.col)) ?? 0) === 0) {
            continue
          }
          for (let row = range.start.row; row <= range.end.row; row += 1) {
            const cellIndex = sheet.grid.get(row, range.start.col)
            if (cellIndex === -1) {
              continue
            }
            if ((args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) === 0) {
              continue
            }
            if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
              args.getDependencyBuildSeen()[cellIndex] = epoch
              args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
              dependencyIndexCount += 1
            }
            args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
            dependencyEntityCount += 1
          }
          continue
        }
        const registered = args.state.ranges.intern(sheet.id, range, {
          ensureCell: (sheetId, row, col) => args.ensureCellTrackedByCoords(sheetId, row, col),
          forEachSheetCell: (sheetId, fn) => args.forEachSheetCell(sheetId, fn),
          isFormulaCell: (cellIndex) => (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0,
        })
        if (symbolicRangeIndex !== -1) {
          args.getSymbolicRangeBindings()[symbolicRangeIndex] = registered.rangeIndex
        }
        const rangeEntity = makeRangeEntity(registered.rangeIndex)
        args.getDependencyBuildEntities()[dependencyEntityCount] = rangeEntity
        dependencyEntityCount += 1
        args.getDependencyBuildRanges()[rangeDependencyCount] = registered.rangeIndex
        rangeDependencyCount += 1
        const memberIndices = args.state.ranges.getFormulaMembersView(registered.rangeIndex)
        for (let memberIndex = 0; memberIndex < memberIndices.length; memberIndex += 1) {
          const cellIndex = memberIndices[memberIndex]!
          if (args.getDependencyBuildSeen()[cellIndex] === epoch) {
            continue
          }
          args.getDependencyBuildSeen()[cellIndex] = epoch
          args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
          dependencyIndexCount += 1
        }
        if (registered.materialized) {
          args.getDependencyBuildNewRanges()[newRangeCount] = registered.rangeIndex
          newRangeCount += 1
        }
        continue
      }
      const parsed = parseCellAddress(dep, currentSheetName)
      const sheetName = parsed.sheetName ?? currentSheetName
      if (parsed.sheetName && !args.state.workbook.getSheet(sheetName)) {
        continue
      }
      const cellIndex = args.ensureCellTracked(sheetName, parsed.text)
      if (args.getDependencyBuildSeen()[cellIndex] !== epoch) {
        args.getDependencyBuildSeen()[cellIndex] = epoch
        args.getDependencyBuildCells()[dependencyIndexCount] = cellIndex
        dependencyIndexCount += 1
      }
      args.getDependencyBuildEntities()[dependencyEntityCount] = makeCellEntity(cellIndex)
      dependencyEntityCount += 1
    }
    return {
      dependencyIndices: args.getDependencyBuildCells().slice(0, dependencyIndexCount),
      dependencyEntities: args.getDependencyBuildEntities().slice(0, dependencyEntityCount),
      rangeDependencies: args.getDependencyBuildRanges().slice(0, rangeDependencyCount),
      symbolicRangeIndices: args.getSymbolicRangeBindings(),
      symbolicRangeCount: compiled.symbolicRanges.length,
      newRangeIndices: args.getDependencyBuildNewRanges(),
      newRangeCount,
    }
  }

  const clearFormulaNow = (cellIndex: number): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (existing) {
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      const col = args.state.workbook.cellStore.cols[cellIndex]
      if (sheetId !== undefined && col !== undefined) {
        const columnKey = formulaColumnCountKey(sheetId, col)
        const nextCount = (formulaColumnCounts.get(columnKey) ?? 1) - 1
        if (nextCount <= 0) {
          formulaColumnCounts.delete(columnKey)
        } else {
          formulaColumnCounts.set(columnKey, nextCount)
        }
      }
      const dependencyEntities = args.edgeArena.readView(existing.dependencyEntities)
      const formulaEntity = makeCellEntity(cellIndex)
      for (let index = 0; index < dependencyEntities.length; index += 1) {
        const dependencyEntity = dependencyEntities[index]!
        removeReverseEdge(dependencyEntity, formulaEntity)
        if (!isRangeEntity(dependencyEntity)) {
          pruneTrackedDependencyCell(entityPayload(dependencyEntity), cellIndex)
        }
      }
      existing.compiled.symbolicNames.forEach((name) => {
        removeDefinedNameReverseEdge(name, cellIndex)
      })
      existing.compiled.symbolicTables.forEach((name) => {
        removeTrackedReverseEdge(args.reverseState.reverseTableEdges, tableDependencyKey(name), cellIndex)
      })
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      existing.compiled.symbolicSpills.forEach((key) => {
        removeTrackedReverseEdge(args.reverseState.reverseSpillEdges, spillDependencyKeyFromRef(key, ownerSheetName), cellIndex)
      })
      const existingDirectLookup = existing.directLookup
      if (existingDirectLookup) {
        const lookupInfo = directLookupColumnInfo(existingDirectLookup)
        const lookupSheet = args.state.workbook.getSheet(lookupInfo.sheetName)
        if (lookupSheet) {
          const lookupEntity = lookupInfo.isExact
            ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
            : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col)
          removeReverseEdge(lookupEntity, formulaEntity)
          if (existingDirectLookup.kind === 'approximate' || existingDirectLookup.kind === 'approximate-uniform-numeric') {
            const rowStart =
              existingDirectLookup.kind === 'approximate' ? existingDirectLookup.prepared.rowStart : existingDirectLookup.rowStart
            const rowEnd = existingDirectLookup.kind === 'approximate' ? existingDirectLookup.prepared.rowEnd : existingDirectLookup.rowEnd
            for (let row = rowStart; row <= rowEnd; row += 1) {
              const memberCellIndex = args.ensureCellTrackedByCoords(lookupSheet.id, row, lookupInfo.col)
              removeReverseEdge(makeCellEntity(memberCellIndex), lookupEntity)
            }
          }
        }
      }
      if (existing.directAggregate) {
        const aggregateSheet = args.state.workbook.getSheet(existing.directAggregate.sheetName)
        if (aggregateSheet) {
          removeTrackedReverseEdge(
            args.reverseState.reverseAggregateColumnEdges,
            aggregateColumnDependencyKey(aggregateSheet.id, existing.directAggregate.col),
            cellIndex,
          )
        }
      }
      for (let index = 0; index < existing.rangeDependencies.length; index += 1) {
        const rangeIndex = existing.rangeDependencies[index]!
        const dependencySources = args.state.ranges.getDependencySourceEntities(rangeIndex)
        const released = args.state.ranges.release(rangeIndex)
        if (!released.removed) {
          continue
        }
        const rangeEntity = makeRangeEntity(rangeIndex)
        for (let sourceIndex = 0; sourceIndex < dependencySources.length; sourceIndex += 1) {
          const dependencyEntity = dependencySources[sourceIndex]!
          removeReverseEdge(dependencyEntity, rangeEntity)
          if (!isRangeEntity(dependencyEntity)) {
            pruneTrackedDependencyCell(entityPayload(dependencyEntity), cellIndex)
          }
        }
        setReverseEdgeSlice(rangeEntity, args.edgeArena.empty())
      }
      args.edgeArena.free(existing.dependencyEntities)
      args.compiledPlans.release(existing.planId)
    }
    args.formulaInstances.delete(cellIndex)
    args.state.formulas.delete(cellIndex)
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
    args.scheduleWasmProgramSync()
    return existing !== undefined
  }

  const bindPreparedFormulaPreparedNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    prepared: ReturnType<typeof prepareFormulaBindingFromCompiledNow>,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    const topologyChanged =
      existing === undefined ||
      !uint32ArrayEqual(args.edgeArena.readView(existing.dependencyEntities), prepared.dependencies.dependencyEntities) ||
      !uint32ArrayEqual(existing.rangeDependencies, prepared.dependencies.rangeDependencies) ||
      !stringArrayEqual(existing.compiled.symbolicNames, prepared.compiled.symbolicNames) ||
      !stringArrayEqual(existing.compiled.symbolicTables, prepared.compiled.symbolicTables) ||
      !stringArrayEqual(existing.compiled.symbolicSpills, prepared.compiled.symbolicSpills) ||
      !directLookupStructureEqual(existing.directLookup, prepared.directLookup) ||
      !directAggregateStructureEqual(existing.directAggregate, prepared.directAggregate) ||
      !directCriteriaStructureEqual(existing.directCriteria, prepared.directCriteria)

    if (existing && !topologyChanged) {
      args.compiledPlans.release(existing.planId)
      existing.source = source
      existing.planId = prepared.plan.id
      existing.templateId = prepared.templateId
      existing.compiled = prepared.plan.compiled
      existing.plan = prepared.plan
      existing.dependencyIndices = prepared.dependencies.dependencyIndices
      existing.runtimeProgram = prepared.runtimeProgram
      existing.constants = prepared.compiled.constants
      existing.programLength = prepared.runtimeProgram.length
      existing.constNumberLength = prepared.compiled.constants.length
      existing.directLookup = prepared.directLookup
      existing.directAggregate = prepared.directAggregate
      existing.directCriteria = prepared.directCriteria
      args.state.workbook.cellStore.flags[cellIndex] =
        ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
      if (existing.compiled.mode === FormulaMode.JsOnly) {
        args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
      } else {
        args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
      }
      args.scheduleWasmProgramSync()
      recordFormulaInstanceNow(cellIndex, source, prepared.templateId)

      primeLookupCandidatesNow(
        ownerSheetName,
        prepared.directLookup,
        prepared.indexedExactLookupCandidates,
        prepared.directApproximateLookupCandidates,
      )
      return false
    }
    if (existing && hasInPlaceDependencyRebindShape(existing, prepared)) {
      updateFormulaDependenciesInPlaceNow(cellIndex, existing, prepared, ownerSheetName, source)
      return true
    }
    if (existing) {
      clearFormulaNow(cellIndex)
    }
    installFreshFormulaNow(cellIndex, ownerSheetName, source, prepared)
    return topologyChanged
  }

  const bindFormulaNow = (cellIndex: number, ownerSheetName: string, source: string): boolean => {
    if (args.state.counters) {
      addEngineCounter(args.state.counters, 'formulasBound')
    }
    return bindPreparedFormulaPreparedNow(cellIndex, ownerSheetName, source, prepareFormulaBindingNow(cellIndex, ownerSheetName, source))
  }

  const bindPreparedFormulaNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
  ): boolean =>
    bindPreparedFormulaPreparedNow(
      cellIndex,
      ownerSheetName,
      source,
      prepareFormulaBindingFromCompiledNow(cellIndex, ownerSheetName, source, compiled as ParsedCompiledFormula, templateId),
    )

  const bindInitialFormulaNow = (cellIndex: number, ownerSheetName: string, source: string): void => {
    if (args.state.counters) {
      addEngineCounter(args.state.counters, 'formulasBound')
    }
    installFreshFormulaNow(cellIndex, ownerSheetName, source, prepareFormulaBindingNow(cellIndex, ownerSheetName, source))
  }

  const installFreshFormulaNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    prepared: ReturnType<typeof prepareFormulaBindingNow>,
  ): void => {
    const dependencyEntities = args.edgeArena.replace(args.edgeArena.empty(), prepared.dependencies.dependencyEntities)
    const runtimeFormula: RuntimeFormula = {
      cellIndex,
      formulaSlotId: 0,
      planId: prepared.plan.id,
      templateId: prepared.templateId,
      source,
      compiled: prepared.plan.compiled,
      plan: prepared.plan,
      dependencyIndices: prepared.dependencies.dependencyIndices,
      dependencyEntities,
      rangeDependencies: prepared.dependencies.rangeDependencies,
      runtimeProgram: prepared.runtimeProgram,
      constants: prepared.compiled.constants,
      programOffset: 0,
      programLength: prepared.runtimeProgram.length,
      constNumberOffset: 0,
      constNumberLength: prepared.compiled.constants.length,
      rangeListOffset: 0,
      rangeListLength: prepared.dependencies.rangeDependencies.length,
      directLookup: prepared.directLookup,
      directAggregate: prepared.directAggregate,
      directCriteria: prepared.directCriteria,
    }
    const formulaSlotId = args.state.formulas.set(cellIndex, runtimeFormula)
    runtimeFormula.formulaSlotId = formulaSlotId
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId !== undefined && col !== undefined) {
      const columnKey = formulaColumnCountKey(sheetId, col)
      formulaColumnCounts.set(columnKey, (formulaColumnCounts.get(columnKey) ?? 0) + 1)
    }
    args.state.workbook.cellStore.flags[cellIndex] =
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
    if (runtimeFormula.compiled.mode === FormulaMode.JsOnly) {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
    } else {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
    }
    recordFormulaInstanceNow(cellIndex, source, prepared.templateId)

    for (let rangeCursor = 0; rangeCursor < prepared.dependencies.newRangeCount; rangeCursor += 1) {
      const rangeIndex = prepared.dependencies.newRangeIndices[rangeCursor]!
      const dependencySources = args.state.ranges.getDependencySourceEntities(rangeIndex)
      const rangeEntity = makeRangeEntity(rangeIndex)
      for (let index = 0; index < dependencySources.length; index += 1) {
        appendReverseEdge(dependencySources[index]!, rangeEntity)
      }
    }
    const formulaEntity = makeCellEntity(cellIndex)
    for (let index = 0; index < prepared.dependencies.dependencyEntities.length; index += 1) {
      appendReverseEdge(prepared.dependencies.dependencyEntities[index]!, formulaEntity)
    }
    runtimeFormula.compiled.symbolicNames.forEach((name) => {
      appendDefinedNameReverseEdge(name, cellIndex)
    })
    runtimeFormula.compiled.symbolicTables.forEach((name) => {
      appendTrackedReverseEdge(args.reverseState.reverseTableEdges, tableDependencyKey(name), cellIndex)
    })
    runtimeFormula.compiled.symbolicSpills.forEach((key) => {
      appendTrackedReverseEdge(
        args.reverseState.reverseSpillEdges,
        spillDependencyKeyFromRef(key, args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)),
        cellIndex,
      )
    })
    if (prepared.directLookup) {
      const lookupInfo = directLookupColumnInfo(prepared.directLookup)
      const lookupSheet = args.state.workbook.getSheet(lookupInfo.sheetName)
      if (lookupSheet) {
        const lookupEntity = lookupInfo.isExact
          ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
          : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col)
        appendReverseEdge(lookupEntity, formulaEntity)
      }
    }
    if (prepared.directAggregate) {
      const aggregateSheet = args.state.workbook.getSheet(prepared.directAggregate.sheetName)
      if (aggregateSheet) {
        appendTrackedReverseEdge(
          args.reverseState.reverseAggregateColumnEdges,
          aggregateColumnDependencyKey(aggregateSheet.id, prepared.directAggregate.col),
          cellIndex,
        )
      }
    }
    args.scheduleWasmProgramSync()

    primeLookupCandidatesNow(
      ownerSheetName,
      prepared.directLookup,
      prepared.indexedExactLookupCandidates,
      prepared.directApproximateLookupCandidates,
    )
  }

  const primeLookupCandidatesNow = (
    ownerSheetName: string,
    directLookup: RuntimeFormula['directLookup'],
    indexedExactLookupCandidates: ReturnType<typeof collectIndexedExactLookupCandidates>,
    directApproximateLookupCandidates: ReturnType<typeof collectDirectApproximateLookupCandidates>,
  ): void => {
    if (directLookup) {
      return
    }
    indexedExactLookupCandidates.forEach((candidate) => {
      if (candidate.startCol !== candidate.endCol) {
        return
      }
      args.exactLookup.primeColumnIndex({
        sheetName: candidate.sheetName ?? ownerSheetName,
        rowStart: candidate.startRow,
        rowEnd: candidate.endRow,
        col: candidate.startCol,
      })
    })
    directApproximateLookupCandidates.forEach((candidate) => {
      if (candidate.startCol !== candidate.endCol) {
        return
      }
      args.sortedLookup.primeColumnIndex({
        sheetName: candidate.sheetName ?? ownerSheetName,
        rowStart: candidate.startRow,
        rowEnd: candidate.endRow,
        col: candidate.startCol,
      })
    })
  }

  const prepareFormulaBindingFromCompiledNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiledInput: ParsedCompiledFormula,
    templateId?: number,
  ) => {
    const ownerSheetId = args.state.workbook.getSheet(ownerSheetName)?.id
    let compiled = normalizeLookupCompileMode(compiledInput)
    const hasLookupInstruction = hasLookupPlanInstruction(compiled.jsPlan)
    const directLookupBinding = hasLookupInstruction ? resolveRuntimeDirectLookupBinding(compiled.jsPlan, ownerSheetName) : undefined
    const directAggregateCandidate = buildDirectAggregateDescriptor({
      compiled,
      ownerSheetName,
    })
    const directAggregate = directAggregateCandidate
    const directCriteria = buildDirectCriteriaDescriptor({
      compiled,
      ownerSheetName,
      workbook: args.state.workbook,
      ensureCellTracked: args.ensureCellTracked,
    })
    const indexedExactLookupCandidates =
      hasLookupInstruction && args.state.getUseColumnIndex() ? collectIndexedExactLookupCandidates(compiled.optimizedAst) : []
    const directApproximateLookupCandidates = hasLookupInstruction ? collectDirectApproximateLookupCandidates(compiled.optimizedAst) : []
    const dependencies = materializeDependencies(ownerSheetName, compiled, directAggregate, directLookupBinding)
    const directLookup = directLookupBinding
      ? buildDirectLookupDescriptor({
          compiled,
          ownerSheetName,
          workbook: args.state.workbook,
          ensureCellTracked: args.ensureCellTracked,
          exactLookup: args.exactLookup,
          sortedLookup: args.sortedLookup,
        })
      : undefined

    ensureDependencyBuildCapacity(
      args.state.workbook.cellStore.size + 1,
      compiled.deps.length + 1,
      compiled.symbolicRefs.length + 1,
      compiled.symbolicRanges.length + 1,
    )
    for (let index = 0; index < compiled.symbolicRefs.length; index += 1) {
      const parsedRef = compiled.parsedSymbolicRefs?.[index]
      if (parsedRef && parsedRef.sheetName === undefined) {
        args.getSymbolicRefBindings()[index] =
          parsedRef.row !== undefined && parsedRef.col !== undefined && ownerSheetId !== undefined
            ? args.ensureCellTrackedByCoords(ownerSheetId, parsedRef.row, parsedRef.col)
            : args.ensureCellTracked(ownerSheetName, parsedRef.address)
        continue
      }
      const ref = compiled.symbolicRefs[index]!
      const [qualifiedSheetName, qualifiedAddress] = ref.includes('!') ? ref.split('!') : [undefined, ref]
      const fallbackAddress = parseCellAddress(qualifiedAddress, qualifiedSheetName).text
      const sheetName =
        parsedRef?.sheetName ??
        qualifiedSheetName ??
        args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if ((parsedRef?.sheetName ?? qualifiedSheetName) && !args.state.workbook.getSheet(sheetName)) {
        args.getSymbolicRefBindings()[index] = UNRESOLVED_WASM_OPERAND
        continue
      }
      args.getSymbolicRefBindings()[index] = args.ensureCellTracked(sheetName, parsedRef?.address ?? fallbackAddress)
    }

    const literalStringIds = compiled.symbolicStrings.map((value) => args.state.strings.intern(value))
    const runtimeProgram = new Uint32Array(compiled.program.length)
    runtimeProgram.set(compiled.program)
    compiled.program.forEach((instruction, index) => {
      const opcode = instruction >>> 24
      const operand = instruction & 0x00ff_ffff
      if (opcode === PUSH_CELL_OPCODE) {
        const targetIndex = operand < compiled.symbolicRefs.length ? (args.getSymbolicRefBindings()[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_CELL_OPCODE << 24) | (targetIndex & 0x00ff_ffff)
        return
      }
      if (opcode === PUSH_RANGE_OPCODE) {
        const targetIndex = operand < dependencies.symbolicRangeCount ? (dependencies.symbolicRangeIndices[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_RANGE_OPCODE << 24) | (targetIndex & 0x00ff_ffff)
        return
      }
      if (opcode === PUSH_STRING_OPCODE) {
        const stringId = operand < literalStringIds.length ? (literalStringIds[operand] ?? 0) : 0
        runtimeProgram[index] = (PUSH_STRING_OPCODE << 24) | (stringId & 0x00ff_ffff)
      }
    })

    return {
      compiled,
      dependencies,
      directLookup,
      directAggregate,
      directCriteria,
      runtimeProgram,
      plan: args.compiledPlans.intern(source, compiled, templateId),
      templateId,
      indexedExactLookupCandidates,
      directApproximateLookupCandidates,
    }
  }

  const prepareFormulaBindingNow = (cellIndex: number, ownerSheetName: string, source: string) => {
    const { compiled, templateResolution } = compileFormulaForCell(cellIndex, ownerSheetName, source)
    return prepareFormulaBindingFromCompiledNow(cellIndex, ownerSheetName, source, compiled, templateResolution.templateId)
  }

  const invalidateFormulaNow = (cellIndex: number): void => {
    clearFormulaNow(cellIndex)
    args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Value))
    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
  }

  const rebindFormulaCellsNow = (candidates: readonly number[], formulaChangedCount: number): number => {
    candidates.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (formula && ownerSheetName) {
        bindFormulaNow(cellIndex, ownerSheetName, formula.source)
      }
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
    })
    return formulaChangedCount
  }

  const rebindTrackedDependentsNow = (registry: Map<string, Set<number>>, keys: readonly string[], formulaChangedCount: number): number =>
    rebindFormulaCellsNow(collectTrackedDependents(registry, keys), formulaChangedCount)

  const rebindFormulasForSheetNow = (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32): number => {
    if (candidates) {
      for (let index = 0; index < candidates.length; index += 1) {
        const cellIndex = candidates[index]!
        const formula = args.state.formulas.get(cellIndex)
        if (!formula) {
          continue
        }
        const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
        if (!ownerSheetName) {
          continue
        }
        const touchesSheet = formula.compiled.deps.some((dep) => {
          if (!dep.includes('!')) {
            return false
          }
          const [qualifiedSheet] = dep.split('!')
          return qualifiedSheet?.replace(/^'(.*)'$/, '$1') === sheetName
        })
        if (!touchesSheet) {
          continue
        }
        bindFormulaNow(cellIndex, ownerSheetName, formula.source)
        formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
      }
      return formulaChangedCount
    }

    args.state.formulas.forEach((formula, cellIndex) => {
      if (!formula) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (!ownerSheetName) {
        return
      }
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes('!')) {
          return false
        }
        const [qualifiedSheet] = dep.split('!')
        return qualifiedSheet?.replace(/^'(.*)'$/, '$1') === sheetName
      })
      if (!touchesSheet) {
        return
      }
      bindFormulaNow(cellIndex, ownerSheetName, formula.source)
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
    })

    return formulaChangedCount
  }

  return {
    bindFormula(cellIndex, ownerSheetName, source) {
      return Effect.try({
        try: () => {
          return bindFormulaNow(cellIndex, ownerSheetName, source)
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to bind formula', cause),
            cause,
          }),
      })
    },
    clearFormula(cellIndex) {
      return Effect.try({
        try: () => clearFormulaNow(cellIndex),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to clear formula', cause),
            cause,
          }),
      })
    },
    invalidateFormula(cellIndex) {
      return Effect.try({
        try: () => {
          invalidateFormulaNow(cellIndex)
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to invalidate formula', cause),
            cause,
          }),
      })
    },
    rewriteCellFormulasForSheetRename(oldSheetName, newSheetName, formulaChangedCount) {
      return Effect.try({
        try: () => {
          args.state.formulas.forEach((formula, cellIndex) => {
            if (!formula) {
              return
            }
            const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
            if (!ownerSheetName) {
              return
            }
            const nextSource = renameFormulaSheetReferences(formula.source, oldSheetName, newSheetName)
            if (nextSource === formula.source && ownerSheetName !== newSheetName) {
              return
            }
            bindFormulaNow(cellIndex, ownerSheetName, nextSource)
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
          })
          return formulaChangedCount
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rewrite formulas for sheet rename', cause),
            cause,
          }),
      })
    },
    rebuildAllFormulaBindings() {
      return Effect.try({
        try: () => {
          const pending = [...args.state.formulas.entries()].map(([cellIndex, formula]) => ({
            cellIndex,
            source: formula.source,
            dependencyIndices: [...formula.dependencyIndices],
            planId: formula.planId,
          }))
          pending.forEach(({ planId }) => {
            args.compiledPlans.release(planId)
          })
          args.state.formulas.clear()
          args.formulaInstances.clear()
          args.state.ranges.reset()
          args.edgeArena.reset()
          args.programArena.reset()
          args.constantArena.reset()
          args.rangeListArena.reset()
          args.reverseState.reverseCellEdges.length = 0
          args.reverseState.reverseRangeEdges.length = 0
          args.reverseState.reverseDefinedNameEdges.clear()
          args.reverseState.reverseTableEdges.clear()
          args.reverseState.reverseSpillEdges.clear()
          args.reverseState.reverseAggregateColumnEdges.clear()
          args.reverseState.reverseExactLookupColumnEdges.clear()
          args.reverseState.reverseSortedLookupColumnEdges.clear()
          formulaColumnCounts.clear()

          const activeCellIndices: number[] = []
          pending.forEach(({ cellIndex, source }) => {
            if (!isCellIndexMappedNow(cellIndex)) {
              args.state.workbook.pruneCellIfEmpty(cellIndex)
              return
            }
            const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
            if (!ownerSheetName || !args.state.workbook.getSheet(ownerSheetName)) {
              return
            }
            try {
              bindFormulaNow(cellIndex, ownerSheetName, source)
            } catch {
              invalidateFormulaNow(cellIndex)
            }
            activeCellIndices.push(cellIndex)
          })
          pruneOrphanedDependencyCells(pending.flatMap(({ dependencyIndices }) => dependencyIndices))
          return activeCellIndices
        },
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rebuild formula bindings', cause),
            cause,
          }),
      })
    },
    rebindFormulaCells(candidates, formulaChangedCount) {
      return Effect.try({
        try: () => rebindFormulaCellsNow(candidates, formulaChangedCount),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rebind formula cells', cause),
            cause,
          }),
      })
    },
    rebindDefinedNameDependents(names, formulaChangedCount) {
      return Effect.try({
        try: () =>
          rebindTrackedDependentsNow(
            args.reverseState.reverseDefinedNameEdges,
            names.map((name) => normalizeDefinedName(name)),
            formulaChangedCount,
          ),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rebind defined name dependents', cause),
            cause,
          }),
      })
    },
    rebindTableDependents(tableNames, formulaChangedCount) {
      return Effect.try({
        try: () => rebindTrackedDependentsNow(args.reverseState.reverseTableEdges, tableNames, formulaChangedCount),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rebind table dependents', cause),
            cause,
          }),
      })
    },
    rebindFormulasForSheet(sheetName, formulaChangedCount, candidates) {
      return Effect.try({
        try: () => rebindFormulasForSheetNow(sheetName, formulaChangedCount, candidates),
        catch: (cause) =>
          new EngineFormulaBindingError({
            message: formulaBindingErrorMessage('Failed to rebind formulas for sheet', cause),
            cause,
          }),
      })
    },
    bindFormulaNow,
    bindPreparedFormulaNow,
    rewriteFormulaCompiledPreservingBindingNow,
    bindInitialFormulaNow,
    clearFormulaNow,
    invalidateFormulaNow,
    refreshRangeDependenciesNow,
    retargetRangeDependenciesNow,
    rebindFormulaCellsNow,
    rebindDefinedNameDependentsNow(names, formulaChangedCount) {
      return rebindFormulaCellsNow(collectTrackedDependents(args.reverseState.reverseDefinedNameEdges, names), formulaChangedCount)
    },
    rebindTableDependentsNow(tableNames, formulaChangedCount) {
      const normalized = tableNames.map((name) => tableDependencyKey(name))
      return rebindFormulaCellsNow(collectTrackedDependents(args.reverseState.reverseTableEdges, normalized), formulaChangedCount)
    },
    rebindFormulasForSheetNow,
  }
}
