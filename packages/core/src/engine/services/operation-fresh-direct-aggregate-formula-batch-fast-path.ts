import { ErrorCode, ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { formatAddress } from '@bilig/formula'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE = 32
const FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT = 4096

type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'strings'
  | 'events'
  | 'formulas'
  | 'counters'
  | 'replicaState'
  | 'getLastMetrics'
  | 'setLastMetrics'
  | 'getSyncClientConnection'
  | 'trackReplicaVersions'
>

interface FreshDirectAggregateFormulaEntry {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: NonNullable<CreateEngineOperationServiceArgs['compileTemplateFormula']> extends (...args: never[]) => infer Result
    ? Result extends { readonly compiled: infer Compiled }
      ? Compiled
      : never
    : never
  readonly templateId: number
  readonly result: DirectScalarCurrentOperand
}

export interface OperationFreshDirectAggregateFormulaBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly bindPreparedFormula: NonNullable<CreateEngineOperationServiceArgs['bindPreparedFormula']>
  readonly compileTemplateFormula: NonNullable<CreateEngineOperationServiceArgs['compileTemplateFormula']>
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly getChangedInputBuffer: () => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
}

export function createOperationFreshDirectAggregateFormulaBatchFastPath(args: OperationFreshDirectAggregateFormulaBatchFastPathArgs): {
  readonly tryApplyFreshDirectAggregateFormulaBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyFreshDirectAggregateFormulaBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (
      source !== 'local' ||
      firstRef === undefined ||
      refs.length < FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE ||
      potentialNewCells !== refs.length ||
      args.state.workbook.metadata.definedNames.size > 0
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (!sheet) {
      return false
    }
    const entries = collectFreshDirectAggregateFormulaEntries(args, refs, firstRef, sheet.name)
    if (entries === null) {
      return false
    }
    if (args.state.trackReplicaVersions && batch === null) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + refs.length)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(firstRef.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(firstRef.sheetId, 'column')
    const colIdsByColumn = new Map<number, string>()
    const colId = (col: number): string => {
      let cached = colIdsByColumn.get(col)
      if (cached === undefined) {
        cached = ensureColId(col)
        colIdsByColumn.set(col, cached)
      }
      return cached
    }
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < entries.length; index += 1) {
          const entry = entries[index]!
          const cellIndex = args.state.workbook.cellStore.allocate(firstRef.sheetId, entry.row, entry.col)
          args.state.workbook.attachAllocatedCellWithLogicalAxisIds(
            firstRef.sheetId,
            entry.row,
            entry.col,
            cellIndex,
            ensureRowId(entry.row),
            colId(entry.col),
          )
          args.bindPreparedFormula(cellIndex, sheet.name, entry.source, entry.compiled, entry.templateId, {
            assumeFreshFormula: true,
          })
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(sheet.name, formatAddress(entry.row, entry.col), batchOpOrder(batch, index))
          }
        }
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed: requiresChangedSet ? changedInputArray : EMPTY_CHANGED_CELLS,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  return { tryApplyFreshDirectAggregateFormulaBatch }
}

function collectFreshDirectAggregateFormulaEntries(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
  ownerSheetName: string,
): readonly FreshDirectAggregateFormulaEntry[] | null {
  const entries: FreshDirectAggregateFormulaEntry[] = []
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }
  for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (ref.sheetId !== firstRef.sheetId || ref.cellIndex !== undefined || mutation.kind !== 'setCellFormula') {
      return null
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined ||
      args.hasTrackedExactLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedSortedLookupDependents(ref.sheetId, mutation.col) ||
      args.hasTrackedDirectRangeDependents(ref.sheetId, mutation.col)
    ) {
      return null
    }
    let template: ReturnType<OperationFreshDirectAggregateFormulaBatchFastPathArgs['compileTemplateFormula']>
    try {
      template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
    } catch {
      return null
    }
    const compiled = template.compiled
    if (
      compiled.volatile ||
      compiled.producesSpill ||
      compiled.symbolicNames.length !== 0 ||
      compiled.symbolicTables.length !== 0 ||
      compiled.symbolicSpills.length !== 0
    ) {
      return null
    }
    const aggregate = compiled.directAggregateCandidate
    const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
    if (
      aggregate === undefined ||
      range === undefined ||
      range.refKind !== 'cells' ||
      (range.sheetName ?? ownerSheetName) !== ownerSheetName ||
      range.startRow !== range.endRow ||
      range.startRow !== mutation.row ||
      range.startCol > range.endCol ||
      (mutation.col >= range.startCol && mutation.col <= range.endCol) ||
      range.endCol - range.startCol + 1 > FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT
    ) {
      return null
    }
    const result = evaluateFreshDirectAggregateRow(args, {
      sheetId: ref.sheetId,
      row: range.startRow,
      colStart: range.startCol,
      colEnd: range.endCol,
      aggregateKind: aggregate.aggregateKind,
      resultOffset: aggregate.resultOffset,
    })
    if (result === undefined) {
      return null
    }
    entries.push({
      row: mutation.row,
      col: mutation.col,
      source: mutation.formula,
      compiled,
      templateId: template.templateId,
      result,
    })
  }
  return entries
}

function evaluateFreshDirectAggregateRow(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  request: {
    readonly sheetId: number
    readonly row: number
    readonly colStart: number
    readonly colEnd: number
    readonly aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
    readonly resultOffset: number | undefined
  },
): DirectScalarCurrentOperand | undefined {
  const sheet = args.state.workbook.getSheetById(request.sheetId)
  if (!sheet) {
    return undefined
  }
  const cellStore = args.state.workbook.cellStore
  let sum = 0
  let count = 0
  let averageCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  for (let col = request.colStart; col <= request.colEnd; col += 1) {
    const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(request.row, col) : sheet.grid.get(request.row, col)
    if (memberCellIndex === -1) {
      continue
    }
    if (args.state.formulas.has(memberCellIndex)) {
      return undefined
    }
    const tag = (cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    switch (tag) {
      case ValueTag.Number: {
        const value = cellStore.numbers[memberCellIndex] ?? 0
        sum += value
        count += 1
        averageCount += 1
        minimum = Math.min(minimum, value)
        maximum = Math.max(maximum, value)
        break
      }
      case ValueTag.Boolean: {
        const value = (cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
        sum += value
        count += 1
        averageCount += 1
        minimum = Math.min(minimum, value)
        maximum = Math.max(maximum, value)
        break
      }
      case ValueTag.Error:
        if (request.aggregateKind === 'sum' || request.aggregateKind === 'average') {
          return { kind: 'error', code: (cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
        }
        break
      case ValueTag.Empty:
      case ValueTag.String:
        break
    }
  }
  const result =
    request.aggregateKind === 'sum'
      ? sum
      : request.aggregateKind === 'count'
        ? count
        : request.aggregateKind === 'average'
          ? averageCount === 0
            ? undefined
            : sum / averageCount
          : request.aggregateKind === 'min'
            ? minimum === Number.POSITIVE_INFINITY
              ? 0
              : minimum
            : maximum === Number.NEGATIVE_INFINITY
              ? 0
              : maximum
  if (result === undefined) {
    return { kind: 'error', code: ErrorCode.Div0 }
  }
  return { kind: 'number', value: result + (request.resultOffset ?? 0) }
}
