import type { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import type { CellRangeRef, CellValue, EngineChangedCell, SelectionState } from '@bilig/protocol'
import type { EdgeSlice } from '../../edge-arena.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import type {
  EngineCellMutationRef,
  EngineExistingLiteralCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import type { EnginePatch } from '../../patches/patch-types.js'
import type { FormulaFamilyFreshUniformRunRegistrationArgs, FormulaFamilyRunUpsertArgs } from '../../formula/formula-family-store.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { WorkbookPivotRecord } from '../../workbook-store.js'
import type { EngineRuntimeState, PreparedCellAddress, U32 } from '../runtime-state.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { EngineMutationError } from '../errors.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { OperationDerivedOp } from './operation-derived-op-helpers.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'
import type {
  BindPreparedFormulaOptions,
  FreshDirectAggregateFormulaBindingRun,
  FreshDirectScalarFormulaBindingRun,
} from './formula-binding-service-types.js'

export const ENGINE_OPERATION_TEST_HOOKS_ENABLED = readNodeEnv() === 'test'

export type MutationSource = 'local' | 'remote' | 'restore' | 'undo' | 'redo'

function readNodeEnv(): string | undefined {
  const maybeProcess = Reflect.get(globalThis, 'process')
  if (typeof maybeProcess !== 'object' || maybeProcess === null) {
    return undefined
  }
  const maybeEnv = Reflect.get(maybeProcess, 'env')
  if (typeof maybeEnv !== 'object' || maybeEnv === null) {
    return undefined
  }
  const nodeEnv = Reflect.get(maybeEnv, 'NODE_ENV')
  return typeof nodeEnv === 'string' ? nodeEnv : undefined
}

export type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

export interface EngineOperationService {
  readonly __testHooks: Record<string, unknown>
  readonly applyBatchNow: (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
    options?: { readonly emitTracked?: boolean },
  ) => void
  readonly applyLocalSingleStructuralAxisOpWithoutBatchNow: (op: StructuralAxisOp, options?: { readonly emitTracked?: boolean }) => boolean
  readonly applyBatch: (
    batch: EngineOpBatch,
    source: MutationSource,
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ) => Effect.Effect<void, EngineMutationError>
  readonly applyCellMutationsAt: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly applyCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => void
  readonly applyExistingNumericCellMutationAtNow: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyExistingLiteralCellMutationAtNow: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyDerivedOp: (op: OperationDerivedOp) => Effect.Effect<number[], EngineMutationError>
}

export interface CreateEngineOperationServiceArgs {
  readonly state: Pick<
    EngineRuntimeState,
    | 'workbook'
    | 'ranges'
    | 'strings'
    | 'wasm'
    | 'events'
    | 'formulas'
    | 'counters'
    | 'replicaState'
    | 'entityVersions'
    | 'sheetDeleteVersions'
    | 'batchListeners'
    | 'redoStack'
    | 'trackReplicaVersions'
    | 'getSyncClientConnection'
    | 'getLastMetrics'
    | 'setLastMetrics'
  >
  readonly reverseState: {
    readonly reverseCellEdges: Array<EdgeSlice | undefined>
    readonly reverseSpillEdges: Map<string, Set<number>>
    readonly reverseAggregateColumnEdges: Map<number, Set<number>>
    readonly reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    readonly reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
  readonly getSelectionState: () => SelectionState
  readonly setSelection: (sheetName: string, address: string) => void
  readonly rewriteDefinedNamesForSheetRename: (oldSheetName: string, newSheetName: string) => void
  readonly rewriteCellFormulasForSheetRename: (oldSheetName: string, newSheetName: string, formulaChangedCount: number) => number
  readonly rebindDefinedNameDependents: (names: readonly string[], formulaChangedCount: number) => number
  readonly collectFormulaCellsForDefinedNames: (names: readonly string[]) => readonly number[]
  readonly rebindTableDependents: (tableNames: readonly string[], formulaChangedCount: number) => number
  readonly rebindFormulaCells: (candidates: readonly number[], formulaChangedCount: number) => number
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly rebindFormulasForSheet: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly removeSheetRuntime: (
    sheetName: string,
    explicitChangedCount: number,
  ) => { changedInputCount: number; formulaChangedCount: number; explicitChangedCount: number }
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => {
    transaction: StructuralTransaction
    changedCellIndices: number[]
    precomputedChangedInputCellIndices: number[]
    formulaCellIndices: number[]
    topologyChanged: boolean
    graphRefreshRequired: boolean
  }
  readonly clearOwnedSpill: (cellIndex: number) => number[]
  readonly clearPivotForCell: (cellIndex: number) => number[]
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
  readonly removeFormula: (cellIndex: number) => boolean
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly rewriteFormulaSourcePreservingBinding?: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly bindPreparedFormula?: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    options?: BindPreparedFormulaOptions,
  ) => boolean
  readonly bindFreshDirectAggregateFormulaRun?: (run: FreshDirectAggregateFormulaBindingRun) => void
  readonly bindFreshDirectScalarFormulaRun?: (run: FreshDirectScalarFormulaBindingRun) => void
  readonly registerFreshFormulaFamilyRun?: (run: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly upsertFormulaFamilyRun?: (run: FormulaFamilyRunUpsertArgs) => void
  readonly upsertFreshFormulaInstances?: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly compileTemplateFormula?: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly setInvalidFormulaValue: (cellIndex: number) => void
  readonly beginEvaluationBudget: (startedAtMs: number) => void
  readonly endEvaluationBudget: () => void
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly beginMutationCollection: (options?: { readonly ensureScratch?: boolean }) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChanged: (count: number) => number
  readonly hasVolatileFormulas?: () => boolean
  readonly markSpillRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markPivotRootsChanged: (cellIndices: readonly number[], count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly composeDisjointEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly captureChangedPatches: (
    changedCellIndices: readonly number[] | U32,
    request?: {
      invalidation?: 'cells' | 'full'
      invalidatedRanges?: readonly CellRangeRef[]
      invalidatedRows?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
      invalidatedColumns?: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    },
  ) => readonly EnginePatch[]
  readonly getChangedInputBuffer: () => U32
  readonly getChangedFormulaBuffer: () => U32
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly ensureCellTracked: (sheetName: string, address: string) => number
  readonly estimatePotentialNewCells: (ops: readonly EngineOp[]) => number
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly syncDynamicRanges: (formulaChangedCount: number) => number
  readonly rebuildTopoRanks: () => void
  readonly repairTopoRanks: (changedFormulaCells: readonly number[] | U32) => boolean
  readonly detectCycles: () => void
  readonly recalculate: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly evaluateFormulaCell: (cellIndex: number) => readonly number[]
  readonly exactLookup: Pick<ExactColumnIndexService, 'findPreparedVectorMatch'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'findPreparedVectorMatch'>
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly prepareRegionQueryIndices: () => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly getCellDependents?: (cellIndex: number) => Uint32Array
  readonly getSingleCellDependent?: (cellIndex: number) => number
  readonly collectFormulaDependents: (entityId: number) => Uint32Array
  readonly hasRegionFormulaSubscriptionsForColumn: (sheetName: string, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsForColumnAt?: (sheetId: number, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsIntersectingRect?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly hasRegionFormulaSubscriptions?: () => boolean
  readonly hasRegionFormulaSubscriptionsOverlappingRange?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly getRegionFormulaSubscriptionCount?: () => number
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => Uint32Array
  readonly collectSingleRegionFormulaDependentForCell: (sheetName: string, row: number, col: number) => number
  readonly collectSingleRegionFormulaDependentForCellAt?: (sheetId: number, row: number, col: number) => number
  readonly noteExactLookupLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
    oldStringId?: number
    newStringId?: number
  }) => void
  readonly noteAggregateLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
  }) => void
  readonly noteSortedLookupLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
    oldStringId?: number
    newStringId?: number
  }) => void
  readonly invalidateExactLookupColumn: (request: { sheetName: string; col: number }) => void
  readonly invalidateSortedLookupColumn: (request: { sheetName: string; col: number }) => void
  readonly invalidateAggregateColumn: (request: { sheetName: string; col: number }) => void
}
