import type { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef, EngineFormulaSourceRefs } from '../../cell-mutations-at.js'
import type { FormulaFamilyFreshUniformRunRegistrationArgs, FormulaFamilyRunUpsertArgs } from '../../formula/formula-family-store.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import type { EngineMutationError } from '../errors.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import type { FormulaOwnerPosition, FreshDirectScalarFormulaBindingRun } from './formula-binding-service-types.js'
import type { DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import type { InitialFormulaEntryRefSource } from './formula-initialization-refs.js'

export interface EngineFormulaInitializationService {
  readonly initializeCellFormulasAt: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeCellFormulasAtNow: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializeFormulaSourcesAtNow: (refs: EngineFormulaSourceRefs, potentialNewCells?: number) => void
  readonly initializePreparedCellFormulasAt: (
    refs: InitialFormulaEntryRefSource<PreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializePreparedCellFormulasAtNow: (
    refs: InitialFormulaEntryRefSource<PreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ) => void
  readonly initializeHydratedPreparedCellFormulasAt: (
    refs: InitialFormulaEntryRefSource<HydratedPreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeHydratedPreparedCellFormulasAtNow: (
    refs: InitialFormulaEntryRefSource<HydratedPreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ) => void
  readonly initializeCachedFormulaSourcesAtNow: (refs: readonly CachedFormulaInitializationRef[], potentialNewCells?: number) => void
}

export interface PreparedFormulaInitializationRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
  readonly cellIndex?: number
}

export interface HydratedPreparedFormulaInitializationRef extends PreparedFormulaInitializationRef {
  readonly value: CellValue
  readonly preserveCachedValueOnFullRecalc?: boolean
}

export interface CachedFormulaInitializationRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly value: CellValue
  readonly cellIndex?: number
}

export interface EngineFormulaInitializationServiceArgs {
  readonly state: Pick<
    EngineRuntimeState,
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'ranges' | 'counters' | 'getLastMetrics' | 'setLastMetrics'
  >
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => void
  readonly withInitialFormulaCells: <T>(cellIndices: readonly number[] | U32, callback: () => T) => T
  readonly bindPreparedFormula: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    options?: {
      readonly deferFamilyRegistration?: boolean
      readonly deferFormulaInstanceRegistration?: boolean
      readonly assumeFreshFormula?: boolean
      readonly preserveCachedValueOnFullRecalc?: boolean
      readonly resolveWorkbookDateSystem?: () => string | undefined
      readonly ownerPosition?: FormulaOwnerPosition
    },
  ) => boolean
  readonly bindFreshDirectScalarFormulaRun?: (run: FreshDirectScalarFormulaBindingRun) => void
  readonly upsertFormulaFamilyRun: (args: FormulaFamilyRunUpsertArgs) => void
  readonly registerFreshFormulaFamilyRun: (args: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly deferFormulaFamilyIndexRebuild?: () => void
  readonly deferFormulaFamilyIndexRuns?: (runs: readonly DeferredInitialFormulaFamilyRun[]) => void
  readonly deferFormulaInstanceTableRebuild?: () => void
  readonly hydrateFreshFormulaInstances?: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly compileTemplateFormula: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly clearTemplateFormulaCache: () => void
  readonly removeFormula: (cellIndex: number) => boolean
  readonly setInvalidFormulaValue: (cellIndex: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChanged: (count: number) => number
  readonly hasVolatileFormulas?: () => boolean
  readonly syncDynamicRanges: (formulaChangedCount: number) => number
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32
  readonly getChangedInputBuffer: () => U32
  readonly getChangedFormulaBuffer: () => U32
  readonly rebuildTopoRanks: () => void
  readonly repairTopoRanks: (changedFormulaCells: readonly number[] | U32) => boolean
  readonly detectCycles: () => void
  readonly recalculate: (changedRoots: readonly number[] | U32, kernelSyncRoots?: readonly number[] | U32) => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly recalculatePreordered: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32
  readonly beginEvaluationBudget: (startedAtMs: number) => void
  readonly endEvaluationBudget: () => void
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly prepareRegionQueryIndices: () => void
  readonly writeHydratedFormulaValue: (cellIndex: number, value: CellValue) => void
}
