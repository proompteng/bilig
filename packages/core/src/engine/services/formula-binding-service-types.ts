import type { Effect } from 'effect'
import type { CompiledFormula, Float64Arena, FormulaNode, StructuralAxisTransform, Uint32Arena } from '@bilig/formula'
import type { EdgeArena } from '../../edge-arena.js'
import type { RegionGraph } from '../../deps/region-graph.js'
import type {
  FormulaFamily,
  FormulaFamilyFreshUniformRunRegistrationArgs,
  FormulaFamilyRunUpsertArgs,
  FormulaFamilyStats,
  FormulaFamilyStore,
  FormulaFamilyStructuralSourceTransform,
  FormulaFamilyStructuralSourceTransformEntry,
} from '../../formula/formula-family-store.js'
import type { FormulaInstanceSnapshot, FormulaInstanceTable } from '../../formula/formula-instance-table.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import type { EngineCounters } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { EngineFormulaBindingError } from '../errors.js'
import type { EngineCompiledPlanService } from './compiled-plan-service.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import type { FormulaBindingReverseEdgeState } from './formula-binding-reverse-edges.js'
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
    options?: BindPreparedFormulaOptions,
  ) => boolean
  readonly bindFreshDirectAggregateFormulaRunNow: (run: FreshDirectAggregateFormulaBindingInput) => void
  readonly bindFreshDirectScalarFormulaRunNow: (run: FreshDirectScalarFormulaBindingInput) => void
  readonly rewriteFormulaSourcePreservingBindingNow: (cellIndex: number, ownerSheetName: string, source: string) => boolean
  readonly rewriteFormulaCompiledPreservingBindingNow: (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    ownerPosition?: FormulaOwnerPosition,
  ) => boolean
  readonly rewriteFormulaMetadataPreservingRuntimeNow: (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    ownerPosition?: FormulaOwnerPosition,
  ) => boolean
  readonly deferCellFormulasForSheetRenameNow: (oldSheetName: string, newSheetName: string) => number
  readonly rewriteCellFormulasForSheetRenameNow: (oldSheetName: string, newSheetName: string, formulaChangedCount: number) => number
  readonly retargetDirectAggregateFormulaForStructuralTransformNow: (
    cellIndex: number,
    ownerSheetName: string,
    targetSheetName: string,
    transform: StructuralAxisTransform,
    preservesValue: boolean,
  ) => boolean
  readonly retargetDirectAggregateFormulasForStructuralTransformNow: (
    inputs: readonly {
      readonly cellIndex: number
      readonly ownerSheetName: string
      readonly preservesValue: boolean
    }[],
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ) => readonly number[]
  readonly bindInitialFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => void
  readonly withInitialFormulaCellsNow: <T>(cellIndices: readonly number[] | U32, callback: () => T) => T
  readonly clearFormulaNow: (cellIndex: number) => boolean
  readonly invalidateFormulaNow: (cellIndex: number) => void
  readonly clearFormulaBookkeepingNow: () => void
  readonly deferFormulaFamilyIndexRebuildNow: () => void
  readonly deferFormulaFamilyIndexRunsNow: (runs: readonly DeferredInitialFormulaFamilyRun[]) => void
  readonly registerFreshFormulaFamilyRunNow: (run: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly upsertFormulaFamilyRunNow: (run: FormulaFamilyRunUpsertArgs) => void
  readonly deferFormulaInstanceTableRebuildNow: () => void
  readonly upsertFreshFormulaInstancesNow: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly hydrateFreshFormulaInstancesNow: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly exportFormulaInstancesNow: () => FormulaInstanceSnapshot[]
  readonly refreshRangeDependenciesNow: (rangeIndices: readonly number[]) => void
  readonly retargetRangeDependenciesNow: (transaction: StructuralTransaction, rangeIndices: readonly number[]) => void
  readonly rebindFormulaCellsNow: (candidates: readonly number[], formulaChangedCount: number) => number
  readonly rebindDefinedNameDependentsNow: (names: readonly string[], formulaChangedCount: number) => number
  readonly rebindTableDependentsNow: (tableNames: readonly string[], formulaChangedCount: number) => number
  readonly rebindFormulasForSheetNow: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
  readonly forEachFormulaCellOwnedBySheetNow: (sheetName: string, fn: (cellIndex: number) => void) => void
  readonly countFormulaSheetMembersNow: (sheetId: number) => number
  readonly countFormulaFamilySheetMembersNow: (sheetId: number) => number
  readonly canUseFormulaFamilyIndexNow: () => boolean
  readonly isFormulaFamilyIndexReadyNow: () => boolean
  readonly tryDeferFormulaFamilyStructuralSourceTransformsNow: (
    sheetId: number,
    transform: FormulaFamilyStructuralSourceTransform,
    canDeferCellIndex: (cellIndex: number) => boolean,
  ) => number | undefined
  readonly forEachFormulaFamilyNow: (fn: (family: FormulaFamily) => void) => void
  readonly setFormulaFamilyStructuralSourceTransformNow: (familyId: number, transform: FormulaFamilyStructuralSourceTransform) => void
  readonly getFormulaFamilyStructuralSourceTransformNow: (cellIndex: number) => FormulaFamilyStructuralSourceTransform | undefined
  readonly hasFormulaFamilyStructuralSourceTransformsNow: () => boolean
  readonly consumeFormulaFamilyStructuralSourceTransformsNow: () => FormulaFamilyStructuralSourceTransformEntry[]
  readonly collectFormulaCellsOwnedBySheetNow: (sheetName: string) => readonly number[]
  readonly collectFormulaCellsReferencingSheetNow: (sheetName: string) => readonly number[]
  readonly collectFormulaCellsForDefinedNamesNow: (names: readonly string[]) => readonly number[]
  readonly collectFormulaCellsForTablesNow: (tableNames: readonly string[]) => readonly number[]
  readonly getFormulaFamilyStatsNow: () => FormulaFamilyStats
}

export interface BindPreparedFormulaOptions {
  readonly deferFamilyRegistration?: boolean
  readonly deferFormulaInstanceRegistration?: boolean
  readonly assumeFreshFormula?: boolean
  readonly preserveCachedValueOnFullRecalc?: boolean
  readonly assumeFreshDirectAggregateLiteralInputs?: boolean
  readonly resolveWorkbookDateSystem?: () => string | undefined
  readonly ownerPosition?: FormulaOwnerPosition
}

export interface FormulaOwnerPosition {
  readonly sheetName: string
  readonly row: number
  readonly col: number
}

export interface FreshDirectAggregateFormulaBindingMember {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId: number
  readonly aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
  readonly aggregateRowStart: number
  readonly aggregateRowEnd: number
  readonly aggregateColStart: number
  readonly aggregateColEnd: number
  readonly resultOffset: number | undefined
}

export interface FreshDirectAggregateFormulaBindingRun {
  readonly sheetId: number
  readonly ownerSheetName: string
  readonly cellIndices: readonly number[] | Uint32Array
  readonly members: readonly FreshDirectAggregateFormulaBindingMember[]
}

export interface FreshDirectAggregateFormulaBinding {
  readonly sheetId: number
  readonly ownerSheetName: string
  readonly cellIndex: number
  readonly member: FreshDirectAggregateFormulaBindingMember
}

export type FreshDirectAggregateFormulaBindingInput = FreshDirectAggregateFormulaBindingRun | FreshDirectAggregateFormulaBinding

export interface FreshDirectScalarFormulaBindingMember {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId: number
}

export interface FreshDirectScalarFormulaBindingRun {
  readonly sheetId: number
  readonly ownerSheetName: string
  readonly cellIndices: readonly number[] | Uint32Array
  readonly members: readonly FreshDirectScalarFormulaBindingMember[]
}

export interface FreshDirectScalarFormulaBinding {
  readonly sheetId: number
  readonly ownerSheetName: string
  readonly cellIndex: number
  readonly member: FreshDirectScalarFormulaBindingMember
}

export type FreshDirectScalarFormulaBindingInput = FreshDirectScalarFormulaBindingRun | FreshDirectScalarFormulaBinding

export interface CreateEngineFormulaBindingServiceArgs {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'ranges' | 'getUseColumnIndex'> & {
    counters?: EngineCounters
  }
  readonly regionGraph: RegionGraph
  readonly compiledPlans: EngineCompiledPlanService
  readonly formulaInstances: FormulaInstanceTable
  readonly formulaFamilies: FormulaFamilyStore
  readonly volatileFormulaCells?: Set<number>
  readonly resolveTemplateForCell: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly exactLookup: Pick<ExactColumnIndexService, 'primeColumnIndex' | 'prepareUniformNumericVectorLookup' | 'prepareVectorLookup'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'primeColumnIndex' | 'prepareVectorLookup'>
  readonly edgeArena: EdgeArena
  readonly programArena: Uint32Arena
  readonly constantArena: Float64Arena
  readonly rangeListArena: Uint32Arena
  readonly reverseState: FormulaBindingReverseEdgeState
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
}
