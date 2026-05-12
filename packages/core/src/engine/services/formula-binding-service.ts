import { Effect } from 'effect'
import type { CompiledFormula, StructuralAxisTransform } from '@bilig/formula'
import { FormulaMode, ErrorCode } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EdgeSlice } from '../../edge-arena.js'
import { entityPayload, isRangeEntity, makeCellEntity, makeRangeEntity } from '../../entity-ids.js'
import { tableDependencyKey } from '../../engine-metadata-utils.js'
import { errorValue } from '../../engine-value-utils.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectScalarDescriptor, RuntimeFormula } from '../runtime-state.js'
import { EngineFormulaBindingError } from '../errors.js'
import {
  canRewriteCompiledPreservingBindings,
  canRewriteCompiledPreservingDirectAggregate,
  canRewriteCompiledPreservingDirectScalar,
  directAggregateStructureEqual,
  directCriteriaStructureEqual,
  directLookupStructureEqual,
  hasInPlaceDependencyRebindShape,
  stringArrayEqual,
  uint32ArrayEqual,
} from './formula-binding-shape-helpers.js'
import type { collectDirectApproximateLookupCandidates, collectIndexedExactLookupCandidates } from './formula-binding-lookup-candidates.js'
import { buildDirectScalarDescriptor } from './formula-binding-direct-scalar.js'
import {
  aggregateColumnDependencyKey,
  appendTrackedReverseEdge,
  collectTrackedDependents,
  directCriteriaAggregateColumn,
  directRegionIdsForFormula,
  hasQualifiedDependencies,
  removeTrackedReverseEdge,
} from './formula-binding-dependency-helpers.js'
import {
  buildDirectAggregateDescriptor,
  rewriteDirectAggregateDescriptorForStructuralTransform,
  type ParsedCompiledFormula,
} from './formula-binding-direct-descriptors.js'
import {
  appendFormulaBindingReverseEdge,
  getFormulaBindingReverseEdgeSlice,
  removeFormulaBindingReverseEdge,
  setFormulaBindingReverseEdgeSlice,
  syncFormulaBindingRangeDependencyEdges,
} from './formula-binding-reverse-edges.js'
import { createFormulaBindingSheetIndex } from './formula-binding-sheet-index.js'
import { createFormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import { createFormulaBindingDependencyMaterializer } from './formula-binding-dependency-materializer.js'
import { createPendingInitialFormulaCellTracker } from './formula-binding-pending-formula-cells.js'
import { compileFormulaBindingForCell } from './formula-binding-compile.js'
import { prepareFormulaBindingFromCompiled } from './formula-binding-prepare.js'
import { createFormulaBindingInstanceTracker } from './formula-binding-instance-tracker.js'
import { clearFormulaBindingNow } from './formula-binding-clear.js'
import { installFreshFormulaBindingNow } from './formula-binding-install.js'
import { createFormulaBindingSheetRenameHandler } from './formula-binding-sheet-rename.js'
import { createFormulaBindingRebinds } from './formula-binding-rebind.js'
import { rebuildDeferredFormulaFamilyIndex } from './formula-family-index-rebuild.js'
import { canRetainUnmanagedCompiledPlan, formulaBindingErrorMessage, makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import { normalizeFormulaBindingLookupCompileMode } from './formula-binding-lookup-mode.js'
import { primeFormulaBindingLookupCandidates } from './formula-binding-lookup-primer.js'
import { directAggregateContainsFormulaOwnerCell } from './formula-binding-direct-aggregate-owner.js'
import { ensureFormulaBindingDependencyBuildCapacity } from './formula-binding-dependency-build-capacity.js'
import type {
  BindPreparedFormulaOptions,
  CreateEngineFormulaBindingServiceArgs,
  EngineFormulaBindingService,
  FormulaOwnerPosition,
} from './formula-binding-service-types.js'
export { formulaBindingServiceTestHooks } from './formula-binding-service-test-hooks.js'

export type {
  BindPreparedFormulaOptions,
  CreateEngineFormulaBindingServiceArgs,
  EngineFormulaBindingService,
  FormulaOwnerPosition,
} from './formula-binding-service-types.js'

export function createEngineFormulaBindingService(args: CreateEngineFormulaBindingServiceArgs): EngineFormulaBindingService {
  const resolvedCompiledCache = new Map<string, ParsedCompiledFormula>()
  const formulaMemberCounts = createFormulaBindingMemberCounts()
  const formulaSheetIndex = createFormulaBindingSheetIndex()
  const formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache = new Map()
  let formulaFamilyIndexNeedsRebuild = false
  const { recordFormulaInstanceNow, registerFormulaFamilyNow } = createFormulaBindingInstanceTracker({
    serviceArgs: args,
    formulaFamilyShapeKeyCache,
  })

  const clearFormulaBookkeepingNow = (): void => {
    resolvedCompiledCache.clear()
    formulaMemberCounts.clear()
    formulaSheetIndex.clear()
    formulaFamilyShapeKeyCache.clear()
    formulaFamilyIndexNeedsRebuild = false
  }

  const ensureFormulaFamilyIndexNow = (): void => {
    if (!formulaFamilyIndexNeedsRebuild) {
      return
    }
    rebuildDeferredFormulaFamilyIndex({
      state: args.state,
      store: args.formulaFamilies,
      shapeKeyCache: formulaFamilyShapeKeyCache,
    })
    formulaFamilyIndexNeedsRebuild = false
  }

  const updateVolatileFormulaIndex = (cellIndex: number, formula: RuntimeFormula | undefined): void => {
    if (!args.volatileFormulaCells) {
      return
    }
    if (formula?.compiled.volatile) {
      args.volatileFormulaCells.add(cellIndex)
      return
    }
    args.volatileFormulaCells.delete(cellIndex)
  }

  const trackFormulaSheetIndexes = (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'>,
  ): void => {
    formulaSheetIndex.trackFormula(cellIndex, ownerSheetName, compiled)
  }

  const untrackFormulaSheetIndexes = (
    cellIndex: number,
    ownerSheetName: string | undefined,
    compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'> | undefined,
  ): void => {
    formulaSheetIndex.untrackFormula(cellIndex, ownerSheetName, compiled)
  }

  const ensureDependencyBuildCapacity = (
    cellCapacity: number,
    dependencyCapacity: number,
    symbolicRefCapacity = 0,
    symbolicRangeCapacity = 0,
  ): void => ensureFormulaBindingDependencyBuildCapacity(args, cellCapacity, dependencyCapacity, symbolicRefCapacity, symbolicRangeCapacity)

  const pendingInitialFormulaCells = createPendingInitialFormulaCellTracker({
    getCellCapacity: () => args.state.workbook.cellStore.capacity,
    getSheetId: (cellIndex) => args.state.workbook.cellStore.sheetIds[cellIndex],
    getCol: (cellIndex) => args.state.workbook.getCellPosition(cellIndex)?.col ?? args.state.workbook.cellStore.cols[cellIndex],
    isBoundFormulaCell: (cellIndex) => (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0,
    hasBoundColumnMembers: (sheetId, col) => formulaMemberCounts.hasColumnMembers(sheetId, col),
  })

  const {
    materializeDependencies,
    materializeDirectScalarDependencies,
    materializeDirectAggregateDependencies,
    materializeDirectCriteriaDependencies,
  } = createFormulaBindingDependencyMaterializer({
    serviceArgs: args,
    hasFormulaColumnMembers: pendingInitialFormulaCells.hasColumnMembers,
    isFormulaCell: pendingInitialFormulaCells.isFormulaCell,
    ensureDependencyBuildCapacity,
  })

  const setReverseEdgeSlice = (entityId: number, slice: EdgeSlice): void => {
    setFormulaBindingReverseEdgeSlice(args.reverseState, entityId, slice)
  }

  const getReverseEdgeSlice = (entityId: number) => getFormulaBindingReverseEdgeSlice(args.reverseState, entityId)

  const appendReverseEdge = (entityId: number, dependentEntityId: number): void => {
    appendFormulaBindingReverseEdge(args.reverseState, args.edgeArena, entityId, dependentEntityId)
  }

  const removeReverseEdge = (entityId: number, dependentEntityId: number): void => {
    removeFormulaBindingReverseEdge(args.reverseState, args.edgeArena, entityId, dependentEntityId)
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
    if (refreshed.size > 0) {
      args.scheduleWasmProgramSync()
    }
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
    if (touched.length > 0) {
      args.scheduleWasmProgramSync()
    }
  }

  const syncRangeDependencyEdges = (
    rangeIndex: number,
    deps: { oldDependencySources: Uint32Array; newDependencySources: Uint32Array },
  ): void => {
    syncFormulaBindingRangeDependencyEdges(args.reverseState, args.edgeArena, makeRangeEntity(rangeIndex), deps)
  }

  const rangeDependenciesHaveNoFormulaMembers = (rangeDependencies: Uint32Array): boolean =>
    rangeDependencies.every((rangeIndex) => args.state.ranges.getFormulaMembersView(rangeIndex).length === 0)

  const appendDefinedNameReverseEdge = (name: string, dependentCellIndex: number): void => {
    appendTrackedReverseEdge(args.reverseState.reverseDefinedNameEdges, normalizeDefinedName(name), dependentCellIndex)
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

    const plan = canRetainUnmanagedCompiledPlan(existing.planId, prepared.plan.compiled, prepared.directScalar)
      ? prepared.plan
      : args.compiledPlans.replace(existing.planId, source, prepared.plan.compiled, prepared.templateId)
    args.compiledPlans.release(prepared.plan.id)
    untrackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    existing.source = source
    existing.structuralSourceTransform = undefined
    existing.sourceRenameTransforms = undefined
    existing.planId = plan.id
    existing.templateId = prepared.templateId
    existing.compiled = plan.compiled
    existing.plan = plan
    existing.dependencyIndices = prepared.dependencies.dependencyIndices
    existing.dependencyEntities = args.edgeArena.replace(existing.dependencyEntities, nextDependencies)
    existing.rangeDependencies = prepared.dependencies.rangeDependencies
    existing.graphRangeDependencies = prepared.dependencies.graphRangeDependencies
    existing.runtimeProgram = prepared.runtimeProgram
    existing.constants = prepared.compiled.constants
    existing.programLength = prepared.runtimeProgram.length
    existing.constNumberLength = prepared.compiled.constants.length
    existing.directLookup = undefined
    existing.directAggregate = undefined
    existing.directScalar = prepared.directScalar
    existing.directCriteria = undefined
    updateVolatileFormulaIndex(cellIndex, existing)
    args.state.workbook.cellStore.flags[cellIndex] =
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
    if (existing.compiled.mode === FormulaMode.JsOnly) {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
    } else {
      args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
    }
    if (prepared.compiled.mode === FormulaMode.WasmFastPath && prepared.runtimeProgram.length > 0) {
      args.scheduleWasmProgramSync()
    }
    recordFormulaInstanceNow(cellIndex, source, prepared.templateId)
    registerFormulaFamilyNow(cellIndex, existing)
    trackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    args.regionGraph.replaceFormulaSubscriptions(cellIndex, directRegionIdsForFormula(existing))

    primeLookupCandidatesNow(ownerSheetName, undefined, prepared.indexedExactLookupCandidates, prepared.directApproximateLookupCandidates)
  }

  const rewriteFormulaCompiledPreservingBindingNow = (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    ownerPosition?: FormulaOwnerPosition,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (!existing) {
      return false
    }
    const ownerSheetName =
      ownerPosition?.sheetName ?? args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    if (!ownerSheetName) {
      return false
    }
    const ownerSheetId = ownerPosition?.sheetName
      ? args.state.workbook.getSheet(ownerPosition.sheetName)?.id
      : args.state.workbook.cellStore.sheetIds[cellIndex]
    let nextDirectAggregate: RuntimeDirectAggregateDescriptor | undefined
    let nextDirectScalar: RuntimeDirectScalarDescriptor | undefined
    if (canRewriteCompiledPreservingBindings(existing, compiled) && rangeDependenciesHaveNoFormulaMembers(existing.rangeDependencies)) {
      nextDirectAggregate = undefined
      nextDirectScalar =
        existing.directScalar === undefined
          ? undefined
          : buildDirectScalarDescriptor({
              compiled: compiled as ParsedCompiledFormula,
              ownerSheetName,
              ownerSheetId,
              workbook: args.state.workbook,
              ensureCellTracked: args.ensureCellTracked,
              ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
            })
      if (existing.directScalar !== undefined && !nextDirectScalar) {
        return false
      }
    } else if (canRewriteCompiledPreservingDirectAggregate(existing, compiled)) {
      nextDirectAggregate = buildDirectAggregateDescriptor({
        compiled: compiled as ParsedCompiledFormula,
        ownerSheetName,
        regionGraph: args.regionGraph,
      })
      if (!nextDirectAggregate) {
        return false
      }
    } else if (canRewriteCompiledPreservingDirectScalar(existing, compiled)) {
      updateFormulaDependenciesInPlaceNow(
        cellIndex,
        existing,
        prepareFormulaBindingFromCompiledNow(cellIndex, ownerSheetName, source, compiled as ParsedCompiledFormula, templateId),
        ownerSheetName,
        source,
      )
      return true
    } else {
      return false
    }
    const nextTemplateId = templateId ?? existing.templateId
    const plan = canRetainUnmanagedCompiledPlan(existing.planId, compiled, nextDirectScalar)
      ? makeUnmanagedCompiledPlan(source, compiled, nextTemplateId)
      : args.compiledPlans.replace(existing.planId, source, compiled, nextTemplateId)
    const previousDirectCriteriaAggregate = directCriteriaAggregateColumn(existing.directCriteria)
    const shouldRefreshSheetIndexes = hasQualifiedDependencies(existing.compiled) || hasQualifiedDependencies(compiled)
    if (shouldRefreshSheetIndexes) {
      untrackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    }
    const hadDirectRegions = existing.directAggregate !== undefined || existing.directCriteria !== undefined
    existing.source = source
    existing.structuralSourceTransform = undefined
    existing.sourceRenameTransforms = undefined
    existing.planId = plan.id
    existing.templateId = nextTemplateId
    existing.compiled = plan.compiled
    existing.plan = plan
    existing.constants = compiled.constants
    existing.programLength = compiled.program.length
    existing.constNumberLength = compiled.constants.length
    existing.directAggregate = nextDirectAggregate
    existing.directScalar = nextDirectScalar
    existing.directCriteria = undefined
    updateVolatileFormulaIndex(cellIndex, existing)
    if (previousDirectCriteriaAggregate) {
      const previousCriteriaSheet = args.state.workbook.getSheet(previousDirectCriteriaAggregate.sheetName)
      if (previousCriteriaSheet) {
        removeTrackedReverseEdge(
          args.reverseState.reverseAggregateColumnEdges,
          aggregateColumnDependencyKey(previousCriteriaSheet.id, previousDirectCriteriaAggregate.col),
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
    recordFormulaInstanceNow(cellIndex, source, nextTemplateId, ownerPosition)
    registerFormulaFamilyNow(cellIndex, existing)
    if (shouldRefreshSheetIndexes) {
      trackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    }
    if (hadDirectRegions || existing.directAggregate !== undefined || existing.directCriteria !== undefined) {
      args.regionGraph.replaceFormulaSubscriptions(cellIndex, directRegionIdsForFormula(existing))
    }
    return true
  }

  const rewriteFormulaMetadataPreservingRuntimeNow = (
    cellIndex: number,
    source: string,
    compiled: CompiledFormula,
    templateId?: number,
    ownerPosition?: FormulaOwnerPosition,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (
      !existing ||
      existing.directLookup !== undefined ||
      existing.directAggregate !== undefined ||
      existing.directCriteria !== undefined ||
      !rangeDependenciesHaveNoFormulaMembers(existing.rangeDependencies) ||
      !canRewriteCompiledPreservingBindings(existing, compiled)
    ) {
      return false
    }
    const nextTemplateId = templateId ?? existing.templateId
    const plan = canRetainUnmanagedCompiledPlan(existing.planId, compiled, existing.directScalar)
      ? makeUnmanagedCompiledPlan(source, compiled, nextTemplateId)
      : args.compiledPlans.replace(existing.planId, source, compiled, nextTemplateId)
    const ownerSheetName =
      ownerPosition?.sheetName ?? args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    const shouldRefreshSheetIndexes =
      ownerSheetName !== undefined && (hasQualifiedDependencies(existing.compiled) || hasQualifiedDependencies(compiled))
    if (shouldRefreshSheetIndexes) {
      untrackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    }
    existing.source = source
    existing.structuralSourceTransform = undefined
    existing.sourceRenameTransforms = undefined
    existing.planId = plan.id
    existing.templateId = nextTemplateId
    existing.compiled = plan.compiled
    existing.plan = plan
    existing.constants = compiled.constants
    existing.programLength = existing.runtimeProgram.length
    existing.constNumberLength = compiled.constants.length
    updateVolatileFormulaIndex(cellIndex, existing)
    if (!args.formulaInstances.get(cellIndex)) {
      recordFormulaInstanceNow(cellIndex, source, nextTemplateId, ownerPosition)
    }
    if (shouldRefreshSheetIndexes) {
      trackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    }
    return true
  }

  const retargetDirectAggregateFormulaForStructuralTransformNow = (
    cellIndex: number,
    ownerSheetName: string,
    targetSheetName: string,
    transform: StructuralAxisTransform,
    preservesValue: boolean,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (!existing?.directAggregate) {
      return false
    }
    const previousDirectAggregate = existing.directAggregate
    const nextDirectAggregate = rewriteDirectAggregateDescriptorForStructuralTransform({
      descriptor: previousDirectAggregate,
      targetSheetName,
      transform,
      regionGraph: args.regionGraph,
    })
    if (!nextDirectAggregate) {
      return false
    }
    existing.directAggregate = nextDirectAggregate
    existing.structuralSourceTransform = {
      ownerSheetName,
      targetSheetName,
      transform,
      preservesValue,
    }
    args.regionGraph.replaceSingleFormulaSubscription(cellIndex, previousDirectAggregate.regionId, nextDirectAggregate.regionId)
    return true
  }

  const isCellIndexMappedNow = (cellIndex: number): boolean => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const position = args.state.workbook.getCellPosition(cellIndex)
    if (sheetId === undefined || !position) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    return sheet?.grid.get(position.row, position.col) === cellIndex
  }

  const compileFormulaForCell = (cellIndex: number, currentSheetName: string, source: string) =>
    compileFormulaBindingForCell({
      serviceArgs: args,
      cellIndex,
      currentSheetName,
      source,
      resolvedCompiledCache,
      normalizeLookupCompileMode: normalizeFormulaBindingLookupCompileMode,
    })

  const clearFormulaNow = (cellIndex: number): boolean => {
    return clearFormulaBindingNow({
      serviceArgs: args,
      formulaMemberCounts,
      untrackFormulaSheetIndexes,
      removeReverseEdge,
      setReverseEdgeSlice,
      pruneTrackedDependencyCell,
      updateVolatileFormulaIndex,
      cellIndex,
    })
  }

  const bindPreparedFormulaPreparedNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    prepared: ReturnType<typeof prepareFormulaBindingFromCompiledNow>,
    options: BindPreparedFormulaOptions = {},
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
      existing.structuralSourceTransform = undefined
      existing.sourceRenameTransforms = undefined
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
      existing.directScalar = prepared.directScalar
      existing.directCriteria = prepared.directCriteria
      updateVolatileFormulaIndex(cellIndex, existing)
      args.state.workbook.cellStore.flags[cellIndex] =
        ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)) | CellFlags.HasFormula
      if (existing.compiled.mode === FormulaMode.JsOnly) {
        args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.JsOnly
      } else {
        args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.JsOnly
      }
      if (prepared.compiled.mode === FormulaMode.WasmFastPath && prepared.runtimeProgram.length > 0) {
        args.scheduleWasmProgramSync()
      }
      if (options.deferFormulaInstanceRegistration !== true) {
        recordFormulaInstanceNow(cellIndex, source, prepared.templateId)
      }
      if (options.deferFamilyRegistration !== true) {
        registerFormulaFamilyNow(cellIndex, existing)
      }

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
      return topologyChanged
    }
    if (existing) {
      clearFormulaNow(cellIndex)
    }
    installFreshFormulaNow(cellIndex, ownerSheetName, source, prepared, options)
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
    options: BindPreparedFormulaOptions = {},
  ): boolean =>
    bindPreparedFormulaPreparedNow(
      cellIndex,
      ownerSheetName,
      source,
      prepareFormulaBindingFromCompiledNow(cellIndex, ownerSheetName, source, compiled as ParsedCompiledFormula, templateId),
      options,
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
    options: BindPreparedFormulaOptions = {},
  ): void => {
    installFreshFormulaBindingNow({
      serviceArgs: args,
      formulaMemberCounts,
      appendReverseEdge,
      appendDefinedNameReverseEdge,
      trackFormulaSheetIndexes,
      updateVolatileFormulaIndex,
      recordFormulaInstanceNow,
      registerFormulaFamilyNow,
      primeLookupCandidatesNow,
      cellIndex,
      ownerSheetName,
      source,
      prepared,
      options,
    })
  }

  const primeLookupCandidatesNow = (
    ownerSheetName: string,
    directLookup: RuntimeFormula['directLookup'],
    indexedExactLookupCandidates: ReturnType<typeof collectIndexedExactLookupCandidates>,
    directApproximateLookupCandidates: ReturnType<typeof collectDirectApproximateLookupCandidates>,
  ): void => {
    primeFormulaBindingLookupCandidates({
      serviceArgs: args,
      ownerSheetName,
      directLookup,
      indexedExactLookupCandidates,
      directApproximateLookupCandidates,
    })
  }

  const prepareFormulaBindingFromCompiledNow = (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiledInput: ParsedCompiledFormula,
    templateId?: number,
  ) => {
    return prepareFormulaBindingFromCompiled({
      serviceArgs: args,
      cellIndex,
      ownerSheetName,
      source,
      compiledInput,
      templateId,
      normalizeLookupCompileMode: normalizeFormulaBindingLookupCompileMode,
      dependencyMaterializer: {
        materializeDependencies,
        materializeDirectAggregateDependencies,
        materializeDirectCriteriaDependencies,
        materializeDirectScalarDependencies,
      },
      ensureDependencyBuildCapacity,
      directAggregateContainsOwnerCell: (directAggregate, ownerCellIndex) =>
        directAggregateContainsFormulaOwnerCell(args, directAggregate, ownerCellIndex),
      makeUnmanagedCompiledPlan,
    })
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

  const { rebindFormulaCellsNow, rebindTrackedDependentsNow, rebindFormulasForSheetNow } = createFormulaBindingRebinds({
    serviceArgs: args,
    bindFormulaNow,
  })

  const { deferCellFormulasForSheetRenameNow, rewriteCellFormulasForSheetRenameNow } = createFormulaBindingSheetRenameHandler({
    serviceArgs: args,
    formulaSheetIndex,
    rangeDependenciesHaveNoFormulaMembers,
    untrackFormulaSheetIndexes,
    trackFormulaSheetIndexes,
    canRetainUnmanagedCompiledPlan,
    makeUnmanagedCompiledPlan,
    rewriteFormulaMetadataPreservingRuntimeNow,
    bindPreparedFormulaNow,
  })

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
        try: () => rewriteCellFormulasForSheetRenameNow(oldSheetName, newSheetName, formulaChangedCount),
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
          clearFormulaBookkeepingNow()
          args.regionGraph.reset()

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
    rewriteFormulaMetadataPreservingRuntimeNow,
    deferCellFormulasForSheetRenameNow,
    rewriteCellFormulasForSheetRenameNow,
    retargetDirectAggregateFormulaForStructuralTransformNow,
    bindInitialFormulaNow,
    withInitialFormulaCellsNow: (cellIndices, callback) => pendingInitialFormulaCells.withCells(cellIndices, callback),
    clearFormulaNow,
    invalidateFormulaNow,
    clearFormulaBookkeepingNow,
    deferFormulaFamilyIndexRebuildNow() {
      formulaFamilyShapeKeyCache.clear()
      formulaFamilyIndexNeedsRebuild = true
    },
    deferFormulaInstanceTableRebuildNow() {},
    exportFormulaInstancesNow() {
      return args.formulaInstances.list()
    },
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
    forEachFormulaCellOwnedBySheetNow(sheetName, fn) {
      formulaSheetIndex.getOwnedBySheetSet(sheetName)?.forEach(fn)
    },
    countFormulaSheetMembersNow(sheetId) {
      return formulaMemberCounts.countSheetMembers(sheetId)
    },
    countFormulaFamilySheetMembersNow(sheetId) {
      ensureFormulaFamilyIndexNow()
      return args.formulaFamilies.countSheetMembers(sheetId)
    },
    forEachFormulaFamilyNow(fn) {
      ensureFormulaFamilyIndexNow()
      args.formulaFamilies.forEachFamily(fn)
    },
    setFormulaFamilyStructuralSourceTransformNow(familyId, transform) {
      ensureFormulaFamilyIndexNow()
      args.formulaFamilies.setStructuralSourceTransform(familyId, transform)
    },
    getFormulaFamilyStructuralSourceTransformNow(cellIndex) {
      ensureFormulaFamilyIndexNow()
      return args.formulaFamilies.getStructuralSourceTransform(cellIndex)
    },
    consumeFormulaFamilyStructuralSourceTransformsNow() {
      ensureFormulaFamilyIndexNow()
      return args.formulaFamilies.consumeStructuralSourceTransforms()
    },
    collectFormulaCellsOwnedBySheetNow(sheetName) {
      return formulaSheetIndex.collectOwnedBySheet(sheetName)
    },
    collectFormulaCellsReferencingSheetNow(sheetName) {
      return formulaSheetIndex.collectReferencingSheet(sheetName)
    },
    collectFormulaCellsForDefinedNamesNow(names) {
      return collectTrackedDependents(
        args.reverseState.reverseDefinedNameEdges,
        names.map((name) => normalizeDefinedName(name)),
      )
    },
    collectFormulaCellsForTablesNow(tableNames) {
      return collectTrackedDependents(
        args.reverseState.reverseTableEdges,
        tableNames.map((name) => tableDependencyKey(name)),
      )
    },
    getFormulaFamilyStatsNow() {
      ensureFormulaFamilyIndexNow()
      return args.formulaFamilies.getStats()
    },
  }
}
