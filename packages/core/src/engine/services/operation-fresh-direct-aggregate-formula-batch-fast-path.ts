import { ErrorCode, ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { formatAddress, type CompiledFormula } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import { buildFormulaFamilyShapeKey } from '../../formula/formula-family-deps.js'
import type { FormulaFamilyMember } from '../../formula/formula-family-store.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { EngineRuntimeState, RuntimeDirectAggregateDescriptor, RuntimeFormula, U32 } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'
import { freshMatrixOverlapsFormulaDependencies } from './operation-fresh-matrix-dependency-overlap.js'
import {
  attachFreshDenseDirectAggregateMatrixCells,
  createFreshFormulaCellAttacher,
  createFreshMatrixDirectAggregateTemplate,
  evaluateFreshDirectAggregateMatrixRow,
  materializeFreshMatrixAxisIds,
  normalizeFreshMatrixDirectAggregateOffset,
  tryTranslateFreshMatrixDirectAggregateTemplate,
  type FreshMatrixDirectAggregateTemplate,
} from './operation-fresh-direct-aggregate-matrix-helpers.js'
import { tryEvaluateNativeFreshDirectAggregateMatrixResults } from './operation-fresh-direct-aggregate-native-batch.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const FRESH_DIRECT_AGGREGATE_FORMULA_BATCH_MIN_SIZE = 32
const FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT = 4096
type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'wasm'
  | 'ranges'
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
  readonly aggregateKind: RuntimeDirectAggregateDescriptor['aggregateKind']
  readonly aggregateRowStart: number
  readonly aggregateRowEnd: number
  readonly aggregateColStart: number
  readonly aggregateColEnd: number
  readonly resultOffset: number | undefined
  readonly result: DirectScalarCurrentOperand
}

type FreshDirectAggregateFormulaEntrySeed = Omit<FreshDirectAggregateFormulaEntry, 'result'>

interface ContiguousSingleColumnFormulaBatch {
  readonly rowStart: number
  readonly col: number
}

interface FreshDirectAggregateMatrixBatch {
  readonly sheet: NonNullable<ReturnType<FastPathState['workbook']['getSheetById']>>
  readonly sheetId: number
  readonly rowStart: number
  readonly rowCount: number
  readonly colStart: number
  readonly inputColCount: number
  readonly formulaCol: number
  readonly values: Float64Array
  readonly formulaEntries: readonly FreshDirectAggregateFormulaEntry[]
}

export interface OperationFreshDirectAggregateFormulaBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsOverlappingRange?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly getRegionFormulaSubscriptionCount?: () => number
  readonly hasRegionFormulaSubscriptionsIntersectingRect:
    | ((sheetId: number, rowStart: number, rowEnd: number, colStart: number, colEnd: number) => boolean)
    | undefined
  readonly bindPreparedFormula: NonNullable<CreateEngineOperationServiceArgs['bindPreparedFormula']>
  readonly bindFreshDirectAggregateFormulaRun: CreateEngineOperationServiceArgs['bindFreshDirectAggregateFormulaRun']
  readonly registerFreshFormulaFamilyRun: CreateEngineOperationServiceArgs['registerFreshFormulaFamilyRun']
  readonly upsertFormulaFamilyRun: CreateEngineOperationServiceArgs['upsertFormulaFamilyRun']
  readonly upsertFreshFormulaInstances: CreateEngineOperationServiceArgs['upsertFreshFormulaInstances']
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
  readonly tryApplyFreshDirectAggregateFormulaMatrixBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyFreshDirectAggregateFormulaMatrixBatch = (
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
    if (args.state.trackReplicaVersions && batch === null) {
      return false
    }
    const matrix = collectFreshDirectAggregateMatrixBatch(args, refs, firstRef)
    if (matrix === null || freshMatrixOverlapsFormulaDependencies(args, matrix)) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'column')
    const totalColCount = matrix.inputColCount + 1
    const rowIds = materializeFreshMatrixAxisIds(matrix.rowCount, matrix.rowStart, ensureRowId)
    const columnIds = materializeFreshMatrixAxisIds(totalColCount, matrix.colStart, ensureColId)
    const firstCellIndex = args.state.workbook.cellStore.allocateDenseRowMajorAtReserved(
      matrix.sheetId,
      matrix.rowStart,
      matrix.rowCount,
      matrix.colStart,
      totalColCount,
    )
    attachFreshDenseDirectAggregateMatrixCells(matrix.sheet, firstCellIndex, matrix.rowStart, matrix.colStart, rowIds, columnIds)
    const usesBulkFormulaBinding = args.bindFreshDirectAggregateFormulaRun !== undefined
    const formulaCellIndices =
      args.upsertFormulaFamilyRun === undefined && !usesBulkFormulaBinding ? undefined : new Uint32Array(matrix.rowCount)
    const formulaInstances = args.upsertFreshFormulaInstances === undefined ? undefined : createFreshFormulaInstanceList(matrix.rowCount)
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        let valueIndex = 0
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          const row = matrix.rowStart + rowOffset
          for (let colOffset = 0; colOffset < matrix.inputColCount; colOffset += 1) {
            const col = matrix.colStart + colOffset
            const cellIndex = firstCellIndex + rowOffset * totalColCount + colOffset
            writeFreshNumericLiteralToCellStore(args, cellIndex, matrix.values[valueIndex]!)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            if (requiresChangedSet) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            }
            if (args.state.trackReplicaVersions && batch !== null) {
              args.setCellEntityVersion(matrix.sheet.name, formatAddress(row, col), batchOpOrder(batch, valueIndex))
            }
            valueIndex += 1
          }
        }
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          const entry = matrix.formulaEntries[rowOffset]!
          const cellIndex = firstCellIndex + rowOffset * totalColCount + matrix.inputColCount
          if (!usesBulkFormulaBinding) {
            args.bindPreparedFormula(cellIndex, matrix.sheet.name, entry.source, entry.compiled, entry.templateId, {
              assumeFreshFormula: true,
              assumeFreshDirectAggregateLiteralInputs: true,
              deferFamilyRegistration: formulaCellIndices !== undefined,
              deferFormulaInstanceRegistration: formulaInstances !== undefined,
              ownerPosition: { sheetName: matrix.sheet.name, row: entry.row, col: entry.col },
            })
          }
          if (formulaCellIndices !== undefined) {
            formulaCellIndices[rowOffset] = cellIndex
          }
          if (formulaInstances !== undefined) {
            formulaInstances[rowOffset] = {
              cellIndex,
              sheetName: matrix.sheet.name,
              row: entry.row,
              col: entry.col,
              source: entry.source,
              templateId: entry.templateId,
            }
          }
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(
              matrix.sheet.name,
              formatAddress(entry.row, entry.col),
              batchOpOrder(batch, matrix.values.length + rowOffset),
            )
          }
        }
        if (usesBulkFormulaBinding && formulaCellIndices !== undefined) {
          args.bindFreshDirectAggregateFormulaRun?.({
            sheetId: matrix.sheetId,
            ownerSheetName: matrix.sheet.name,
            cellIndices: formulaCellIndices,
            members: matrix.formulaEntries,
          })
        }
        if (args.upsertFormulaFamilyRun !== undefined && formulaCellIndices !== undefined) {
          registerFreshDirectAggregateFormulaFamilyRun(args, matrix.sheetId, matrix.formulaEntries, formulaCellIndices)
        }
        if (formulaInstances !== undefined) {
          args.upsertFreshFormulaInstances?.(formulaInstances)
        }
        const writtenColumns = new Uint32Array(matrix.inputColCount)
        for (let index = 0; index < matrix.inputColCount; index += 1) {
          writtenColumns[index] = matrix.colStart + index
        }
        args.state.workbook.notifyColumnsWritten(matrix.sheetId, writtenColumns)
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    if (requiresChangedSet && hasTrackedEventListeners && matrix.values.length > 0 && matrix.values.length < changedInputArray.length) {
      tagTrustedPhysicalTrackedChanges(changedInputArray, matrix.sheetId, matrix.values.length)
    }
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
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
    const contiguousBatch = getContiguousSingleColumnFormulaBatch(entries)
    const firstContiguousCellIndex =
      contiguousBatch === undefined
        ? undefined
        : args.state.workbook.cellStore.allocateDenseSingleColumnReserved(
            firstRef.sheetId,
            contiguousBatch.rowStart,
            entries.length,
            contiguousBatch.col,
          )
    if (firstContiguousCellIndex === undefined) {
      args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + refs.length)
    }
    const attachFreshCell = createFreshFormulaCellAttacher(sheet)
    const formulaCellIndices = args.upsertFormulaFamilyRun === undefined ? undefined : new Uint32Array(entries.length)
    const formulaInstances = args.upsertFreshFormulaInstances === undefined ? undefined : createFreshFormulaInstanceList(entries.length)
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < entries.length; index += 1) {
          const entry = entries[index]!
          const cellIndex =
            firstContiguousCellIndex === undefined
              ? args.state.workbook.cellStore.allocateReserved(firstRef.sheetId, entry.row, entry.col)
              : firstContiguousCellIndex + index
          attachFreshCell(entry.row, entry.col, cellIndex, ensureRowId(entry.row), colId(entry.col))
          args.bindPreparedFormula(cellIndex, sheet.name, entry.source, entry.compiled, entry.templateId, {
            assumeFreshFormula: true,
            deferFamilyRegistration: formulaCellIndices !== undefined,
            deferFormulaInstanceRegistration: formulaInstances !== undefined,
            ownerPosition: { sheetName: sheet.name, row: entry.row, col: entry.col },
          })
          if (formulaCellIndices !== undefined) {
            formulaCellIndices[index] = cellIndex
          }
          if (formulaInstances !== undefined) {
            formulaInstances[index] = {
              cellIndex,
              sheetName: sheet.name,
              row: entry.row,
              col: entry.col,
              source: entry.source,
              templateId: entry.templateId,
            }
          }
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(sheet.name, formatAddress(entry.row, entry.col), batchOpOrder(batch, index))
          }
        }
        if (formulaCellIndices !== undefined) {
          registerFreshDirectAggregateFormulaFamilyRun(args, firstRef.sheetId, entries, formulaCellIndices)
        }
        if (formulaInstances !== undefined) {
          args.upsertFreshFormulaInstances?.(formulaInstances)
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

  return { tryApplyFreshDirectAggregateFormulaBatch, tryApplyFreshDirectAggregateFormulaMatrixBatch }
}

function createFreshFormulaInstanceList(count: number): FormulaInstanceSnapshot[] {
  const records: FormulaInstanceSnapshot[] = []
  records.length = count
  return records
}

function registerFreshDirectAggregateFormulaFamilyRun(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  sheetId: number,
  entries: readonly FreshDirectAggregateFormulaEntry[],
  cellIndices: readonly number[] | Uint32Array,
): void {
  const upsertFormulaFamilyRun = args.upsertFormulaFamilyRun
  if (upsertFormulaFamilyRun === undefined || entries.length === 0) {
    return
  }
  const firstRegistration = readBoundFormulaFamilyRegistration(args, cellIndices[0])
  if (firstRegistration === undefined) {
    return
  }

  let uniformSingleColumnRun = true
  let sameFamily = true
  for (let index = 1; index < entries.length; index += 1) {
    const entry = entries[index]!
    const registration = readBoundFormulaFamilyRegistration(args, cellIndices[index])
    if (
      registration === undefined ||
      registration.templateId !== firstRegistration.templateId ||
      registration.shapeKey !== firstRegistration.shapeKey
    ) {
      sameFamily = false
      uniformSingleColumnRun = false
      break
    }
    const firstEntry = entries[0]!
    if (entry.col !== firstEntry.col || entry.row !== firstEntry.row + index) {
      uniformSingleColumnRun = false
    }
  }

  const firstEntry = entries[0]!
  if (
    sameFamily &&
    uniformSingleColumnRun &&
    args.registerFreshFormulaFamilyRun?.({
      sheetId,
      templateId: firstRegistration.templateId,
      shapeKey: firstRegistration.shapeKey,
      axis: 'row',
      fixedIndex: firstEntry.col,
      start: firstEntry.row,
      step: 1,
      cellIndices,
    })
  ) {
    return
  }

  if (sameFamily) {
    upsertFormulaFamilyRun({
      sheetId,
      templateId: firstRegistration.templateId,
      shapeKey: firstRegistration.shapeKey,
      members: materializeFormulaFamilyMembers(entries, cellIndices, 0, entries.length),
    })
    return
  }

  const groups = new Map<string, { templateId: number; shapeKey: string; members: FormulaFamilyMember[] }>()
  for (let index = 0; index < entries.length; index += 1) {
    const registration = readBoundFormulaFamilyRegistration(args, cellIndices[index])
    if (registration === undefined) {
      continue
    }
    const key = `${registration.templateId}\t${registration.shapeKey}`
    let group = groups.get(key)
    if (group === undefined) {
      group = { templateId: registration.templateId, shapeKey: registration.shapeKey, members: [] }
      groups.set(key, group)
    }
    const entry = entries[index]!
    group.members.push({ cellIndex: cellIndices[index]!, row: entry.row, col: entry.col })
  }
  groups.forEach((group) => {
    upsertFormulaFamilyRun({
      sheetId,
      templateId: group.templateId,
      shapeKey: group.shapeKey,
      members: group.members,
    })
  })
}

function readBoundFormulaFamilyRegistration(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  cellIndex: number | undefined,
): { readonly templateId: number; readonly shapeKey: string } | undefined {
  if (cellIndex === undefined) {
    return undefined
  }
  const formula = args.state.formulas.get(cellIndex)
  if (formula === undefined || formula.templateId === undefined) {
    return undefined
  }
  return {
    templateId: formula.templateId,
    shapeKey: directAggregateRuntimeFormulaFamilyShapeKey(formula),
  }
}

function directAggregateRuntimeFormulaFamilyShapeKey(formula: RuntimeFormula): string {
  return buildFormulaFamilyShapeKey({
    compiled: formula.compiled,
    dependencyCount: formula.dependencyIndices.length,
    rangeDependencyCount: formula.rangeDependencies.length,
    directAggregateKind: formula.directAggregate?.aggregateKind,
    directLookupKind: formula.directLookup?.kind,
    directScalarKind: formula.directScalar?.kind,
    directCriteriaKind: formula.directCriteria?.aggregateKind,
  })
}

function materializeFormulaFamilyMembers(
  entries: readonly FreshDirectAggregateFormulaEntry[],
  cellIndices: readonly number[] | Uint32Array,
  start: number,
  end: number,
): FormulaFamilyMember[] {
  const members: FormulaFamilyMember[] = []
  members.length = end - start
  for (let index = start; index < end; index += 1) {
    const entry = entries[index]!
    members[index - start] = { cellIndex: cellIndices[index]!, row: entry.row, col: entry.col }
  }
  return members
}

function getContiguousSingleColumnFormulaBatch(
  entries: readonly FreshDirectAggregateFormulaEntry[],
): ContiguousSingleColumnFormulaBatch | undefined {
  const first = entries[0]
  if (first === undefined) {
    return undefined
  }
  for (let index = 1; index < entries.length; index += 1) {
    const entry = entries[index]!
    if (entry.col !== first.col || entry.row !== first.row + index) {
      return undefined
    }
  }
  return { rowStart: first.row, col: first.col }
}

function collectFreshDirectAggregateMatrixBatch(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
): FreshDirectAggregateMatrixBatch | null {
  const firstMutation = firstRef.mutation
  if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }

  let firstFormulaRefIndex = -1
  for (let index = 0; index < refs.length; index += 1) {
    if (refs[index]!.mutation.kind === 'setCellFormula') {
      firstFormulaRefIndex = index
      break
    }
  }
  if (firstFormulaRefIndex <= 0) {
    return null
  }

  const values = new Float64Array(firstFormulaRefIndex)
  let currentRow = firstMutation.row
  let currentWidth = 0
  let inputColCount = 0
  let rowCount = 1
  for (let refIndex = 0; refIndex < firstFormulaRefIndex; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellValue' ||
      typeof mutation.value !== 'number' ||
      Object.is(mutation.value, -0)
    ) {
      return null
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstMutation.col + currentWidth) {
        return null
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstMutation.col || currentWidth === 0) {
        return null
      }
      if (inputColCount === 0) {
        inputColCount = currentWidth
      } else if (currentWidth !== inputColCount) {
        return null
      }
      currentRow = mutation.row
      currentWidth = 1
      rowCount += 1
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined
    ) {
      return null
    }
    values[refIndex] = mutation.value
  }
  if (inputColCount === 0) {
    inputColCount = currentWidth
  } else if (currentWidth !== inputColCount) {
    return null
  }
  if (rowCount < 2 || inputColCount < 2 || rowCount * inputColCount !== firstFormulaRefIndex) {
    return null
  }
  if (refs.length - firstFormulaRefIndex !== rowCount) {
    return null
  }
  for (let col = firstMutation.col; col < firstMutation.col + inputColCount; col += 1) {
    if (args.hasTrackedExactLookupDependents(firstRef.sheetId, col) || args.hasTrackedSortedLookupDependents(firstRef.sheetId, col)) {
      return null
    }
  }

  const formulaCol = firstMutation.col + inputColCount
  const formulaEntrySeeds: FreshDirectAggregateFormulaEntrySeed[] = []
  let directAggregateTemplate: FreshMatrixDirectAggregateTemplate | undefined
  for (let refIndex = firstFormulaRefIndex; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    const rowOffset = refIndex - firstFormulaRefIndex
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellFormula' ||
      mutation.row !== firstMutation.row + rowOffset ||
      mutation.col !== formulaCol
    ) {
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
    let compiled: CompiledFormula
    let templateId: number
    const translated = directAggregateTemplate
      ? tryTranslateFreshMatrixDirectAggregateTemplate(directAggregateTemplate, mutation.formula, mutation.row, mutation.col)
      : undefined
    if (translated) {
      compiled = translated
      templateId = directAggregateTemplate!.templateId
    } else {
      let template: ReturnType<OperationFreshDirectAggregateFormulaBatchFastPathArgs['compileTemplateFormula']>
      try {
        template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
      } catch {
        return null
      }
      compiled = template.compiled
      templateId = template.templateId
    }
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
      (range.sheetName ?? sheet.name) !== sheet.name ||
      range.startRow !== range.endRow ||
      range.startRow !== mutation.row ||
      range.startCol < firstMutation.col ||
      range.endCol >= formulaCol ||
      range.startCol > range.endCol ||
      range.endCol - range.startCol + 1 > FRESH_DIRECT_AGGREGATE_FORMULA_SCAN_LIMIT
    ) {
      return null
    }
    if (directAggregateTemplate === undefined) {
      directAggregateTemplate = createFreshMatrixDirectAggregateTemplate({
        aggregate,
        compiled,
        formulaCol,
        range,
        row: mutation.row,
        templateId,
      })
    }
    formulaEntrySeeds.push({
      row: mutation.row,
      col: mutation.col,
      source: mutation.formula,
      compiled,
      templateId,
      aggregateKind: aggregate.aggregateKind,
      aggregateRowStart: range.startRow,
      aggregateRowEnd: range.endRow,
      aggregateColStart: range.startCol,
      aggregateColEnd: range.endCol,
      resultOffset: normalizeFreshMatrixDirectAggregateOffset(aggregate.resultOffset),
    })
  }
  const formulaEntries = materializeFreshDirectAggregateFormulaEntries(args, {
    inputColCount,
    matrixColStart: firstMutation.col,
    seeds: formulaEntrySeeds,
    values,
  })

  return {
    sheet,
    sheetId: firstRef.sheetId,
    rowStart: firstMutation.row,
    rowCount,
    colStart: firstMutation.col,
    inputColCount,
    formulaCol,
    values,
    formulaEntries,
  }
}

function materializeFreshDirectAggregateFormulaEntries(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  input: {
    readonly inputColCount: number
    readonly matrixColStart: number
    readonly seeds: readonly FreshDirectAggregateFormulaEntrySeed[]
    readonly values: Float64Array
  },
): FreshDirectAggregateFormulaEntry[] {
  const nativeResults = tryEvaluateNativeFreshDirectAggregateMatrixResults(args, input)
  if (nativeResults !== undefined) {
    return input.seeds.map((seed, index) => ({
      ...seed,
      result: { kind: 'number', value: nativeResults[index]! },
    }))
  }
  return input.seeds.map((seed, rowOffset) => ({
    ...seed,
    result: evaluateFreshDirectAggregateMatrixRow({
      aggregateKind: seed.aggregateKind,
      colEnd: seed.aggregateColEnd,
      colStart: seed.aggregateColStart,
      inputColCount: input.inputColCount,
      matrixColStart: input.matrixColStart,
      resultOffset: seed.resultOffset,
      rowOffset,
      values: input.values,
    }),
  }))
}

function writeFreshNumericLiteralToCellStore(
  args: OperationFreshDirectAggregateFormulaBatchFastPathArgs,
  cellIndex: number,
  value: number,
): void {
  const cellStore = args.state.workbook.cellStore
  cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.stringIds[cellIndex] = 0
  cellStore.tags[cellIndex] = ValueTag.Number
  cellStore.numbers[cellIndex] = value
  cellStore.errors[cellIndex] = ErrorCode.None
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
      aggregateKind: aggregate.aggregateKind,
      aggregateRowStart: range.startRow,
      aggregateRowEnd: range.endRow,
      aggregateColStart: range.startCol,
      aggregateColEnd: range.endCol,
      resultOffset: normalizeFreshMatrixDirectAggregateOffset(aggregate.resultOffset),
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
