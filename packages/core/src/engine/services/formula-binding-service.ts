import type { CompiledFormula, StructuralAxisTransform } from '@bilig/formula'
import { FormulaMode, ErrorCode } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EdgeSlice } from '../../edge-arena.js'
import { entityPayload, isRangeEntity, makeCellEntity } from '../../entity-ids.js'
import { tableDependencyKey } from '../../engine-metadata-utils.js'
import { errorValue } from '../../engine-value-utils.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectScalarDescriptor, RuntimeFormula } from '../runtime-state.js'
import {
  canRewriteCompiledPreservingBindings,
  canRewriteCompiledPreservingDirectAggregate,
  canRewriteCompiledPreservingDirectScalar,
  directScalarDependencyCellsEqual,
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
  appendKnownUniqueFormulaBindingReverseEdge,
  getFormulaBindingReverseEdgeSlice,
  removeFormulaBindingReverseEdge,
  setFormulaBindingReverseEdgeSlice,
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
import { canRetainUnmanagedCompiledPlan, makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import { normalizeFormulaBindingLookupCompileMode } from './formula-binding-lookup-mode.js'
import { primeFormulaBindingLookupCandidates } from './formula-binding-lookup-primer.js'
import { directAggregateContainsFormulaOwnerCell } from './formula-binding-direct-aggregate-owner.js'
import { ensureFormulaBindingDependencyBuildCapacity } from './formula-binding-dependency-build-capacity.js'
import { createFormulaBindingRangeDependencyUpdater } from './formula-binding-range-dependencies.js'
import { clearFormulaRuntimeFlags, markFormulaCellBound } from './formula-binding-cell-flags.js'
import { formulaBindingEffect } from './formula-binding-effect.js'
import { rebuildAllFormulaBindingsNow } from './formula-binding-rebuild.js'
import { createFormulaBindingFamilyIndexController } from './formula-binding-family-index-controller.js'
import { applyFormulaRuntimePlanFields } from './formula-binding-runtime-update.js'
import { rebuildDeferredFormulaFamilyIndex } from './formula-family-index-rebuild.js'
import type {
  BindPreparedFormulaOptions,
  CreateEngineFormulaBindingServiceArgs,
  EngineFormulaBindingService,
  FormulaOwnerPosition,
} from './formula-binding-service-types.js'
export { formulaBindingServiceTestHooks } from './formula-binding-service-test-hooks.js'
export type * from './formula-binding-service-types.js'

export function createEngineFormulaBindingService(args: CreateEngineFormulaBindingServiceArgs): EngineFormulaBindingService {
  const resolvedCompiledCache = new Map<string, ParsedCompiledFormula>()
  const formulaMemberCounts = createFormulaBindingMemberCounts()
  const formulaSheetIndex = createFormulaBindingSheetIndex()
  const formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache = new Map()
  const {
    rebuildFormulaInstancesNow,
    recordFormulaInstanceNow,
    registerFormulaFamilyNow: registerFormulaFamilyInStoreNow,
  } = createFormulaBindingInstanceTracker({
    serviceArgs: args,
    formulaFamilyShapeKeyCache,
  })
  const formulaFamilyIndex = createFormulaBindingFamilyIndexController({
    formulaFamilies: args.formulaFamilies,
    formulaFamilyShapeKeyCache,
    registerFormulaFamilyInStoreNow,
    countFormulaSheetMembersNow: (sheetId) => formulaMemberCounts.countSheetMembers(sheetId),
    rebuildFormulaFamilyIndexNow: () =>
      rebuildDeferredFormulaFamilyIndex({
        state: args.state,
        store: args.formulaFamilies,
        shapeKeyCache: formulaFamilyShapeKeyCache,
      }),
  })
  const { registerFormulaFamilyNow } = formulaFamilyIndex

  const clearFormulaBookkeepingNow = (): void => {
    resolvedCompiledCache.clear()
    formulaMemberCounts.clear()
    formulaSheetIndex.clear()
    formulaFamilyIndex.clearNow()
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

  const dependencyMaterializer = createFormulaBindingDependencyMaterializer({
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

  const appendKnownUniqueReverseEdge = (entityId: number, dependentEntityId: number): void => {
    appendKnownUniqueFormulaBindingReverseEdge(args.reverseState, args.edgeArena, entityId, dependentEntityId)
  }

  const removeReverseEdge = (entityId: number, dependentEntityId: number): void => {
    removeFormulaBindingReverseEdge(args.reverseState, args.edgeArena, entityId, dependentEntityId)
  }

  const { refreshRangeDependenciesNow, retargetRangeDependenciesNow } = createFormulaBindingRangeDependencyUpdater(args)

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
    const previousDependencies = args.edgeArena.readView(existing.dependencyEntities)
    const nextDependencies = prepared.dependencies.dependencyEntities
    const dependencyEntitiesChanged = !uint32ArrayEqual(previousDependencies, nextDependencies)

    if (dependencyEntitiesChanged) {
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
    }

    const plan = canRetainUnmanagedCompiledPlan(existing.planId, prepared.plan.compiled, prepared.directScalar)
      ? prepared.plan
      : args.compiledPlans.replace(existing.planId, source, prepared.plan.compiled, prepared.templateId)
    args.compiledPlans.release(prepared.plan.id)
    untrackFormulaSheetIndexes(cellIndex, ownerSheetName, existing.compiled)
    applyFormulaRuntimePlanFields(existing, {
      source,
      plan,
      templateId: prepared.templateId,
      runtimeProgram: prepared.runtimeProgram,
      programLength: prepared.runtimeProgram.length,
    })
    existing.dependencyIndices = prepared.dependencies.dependencyIndices
    if (dependencyEntitiesChanged) {
      existing.dependencyEntities = args.edgeArena.replace(existing.dependencyEntities, nextDependencies)
    }
    existing.rangeDependencies = prepared.dependencies.rangeDependencies
    existing.graphRangeDependencies = prepared.dependencies.graphRangeDependencies
    existing.directLookup = undefined
    existing.directAggregate = undefined
    existing.directScalar = prepared.directScalar
    existing.directCriteria = undefined
    updateVolatileFormulaIndex(cellIndex, existing)
    markFormulaCellBound(args.state.workbook.cellStore, cellIndex, existing.compiled.mode)
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
        workbook: args.state.workbook,
        regionGraph: args.regionGraph,
      })
      if (!nextDirectAggregate) {
        return false
      }
    } else if (canRewriteCompiledPreservingDirectScalar(existing, compiled)) {
      const replacementDirectScalar = buildDirectScalarDescriptor({
        compiled: compiled as ParsedCompiledFormula,
        ownerSheetName,
        ownerSheetId,
        workbook: args.state.workbook,
        ensureCellTracked: args.ensureCellTracked,
        ensureCellTrackedByCoords: args.ensureCellTrackedByCoords,
      })
      if (replacementDirectScalar && directScalarDependencyCellsEqual(existing.directScalar, replacementDirectScalar)) {
        const nextTemplateId = templateId ?? existing.templateId
        const plan = canRetainUnmanagedCompiledPlan(existing.planId, compiled, replacementDirectScalar)
          ? makeUnmanagedCompiledPlan(source, compiled, nextTemplateId)
          : args.compiledPlans.replace(existing.planId, source, compiled, nextTemplateId)
        applyFormulaRuntimePlanFields(existing, {
          source,
          plan,
          templateId: nextTemplateId,
          runtimeProgram: compiled.program,
          programLength: compiled.program.length,
        })
        existing.directScalar = replacementDirectScalar
        updateVolatileFormulaIndex(cellIndex, existing)
        markFormulaCellBound(args.state.workbook.cellStore, cellIndex, compiled.mode)
        if (compiled.mode === FormulaMode.WasmFastPath && compiled.program.length > 0) {
          args.scheduleWasmProgramSync()
        }
        recordFormulaInstanceNow(cellIndex, source, nextTemplateId, ownerPosition)
        registerFormulaFamilyNow(cellIndex, existing, ownerPosition)
        return true
      }
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
    applyFormulaRuntimePlanFields(existing, {
      source,
      plan,
      templateId: nextTemplateId,
      programLength: compiled.program.length,
    })
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
    markFormulaCellBound(args.state.workbook.cellStore, cellIndex, compiled.mode)
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
    applyFormulaRuntimePlanFields(existing, {
      source,
      plan,
      templateId: nextTemplateId,
      programLength: existing.runtimeProgram.length,
    })
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
    if (!formulaFamilyIndex.isReadyNow()) {
      formulaFamilyIndex.deferRebuildNow()
    }
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
      applyFormulaRuntimePlanFields(existing, {
        source,
        plan: prepared.plan,
        templateId: prepared.templateId,
        runtimeProgram: prepared.runtimeProgram,
        programLength: prepared.runtimeProgram.length,
      })
      existing.dependencyIndices = prepared.dependencies.dependencyIndices
      existing.directLookup = prepared.directLookup
      existing.directAggregate = prepared.directAggregate
      existing.directScalar = prepared.directScalar
      existing.directCriteria = prepared.directCriteria
      updateVolatileFormulaIndex(cellIndex, existing)
      markFormulaCellBound(args.state.workbook.cellStore, cellIndex, existing.compiled.mode)
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

  const rewriteFormulaSourcePreservingBindingNow = (cellIndex: number, ownerSheetName: string, source: string): boolean => {
    const { compiled, templateResolution } = compileFormulaForCell(cellIndex, ownerSheetName, source)
    return rewriteFormulaCompiledPreservingBindingNow(cellIndex, source, compiled, templateResolution.templateId)
  }

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
      appendKnownUniqueReverseEdge,
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
      dependencyMaterializer,
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
    clearFormulaRuntimeFlags(args.state.workbook.cellStore, cellIndex)
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
      return formulaBindingEffect('Failed to bind formula', () => bindFormulaNow(cellIndex, ownerSheetName, source))
    },
    clearFormula(cellIndex) {
      return formulaBindingEffect('Failed to clear formula', () => clearFormulaNow(cellIndex))
    },
    invalidateFormula(cellIndex) {
      return formulaBindingEffect('Failed to invalidate formula', () => {
        invalidateFormulaNow(cellIndex)
      })
    },
    rewriteCellFormulasForSheetRename(oldSheetName, newSheetName, formulaChangedCount) {
      return formulaBindingEffect('Failed to rewrite formulas for sheet rename', () =>
        rewriteCellFormulasForSheetRenameNow(oldSheetName, newSheetName, formulaChangedCount),
      )
    },
    rebuildAllFormulaBindings() {
      return formulaBindingEffect('Failed to rebuild formula bindings', () =>
        rebuildAllFormulaBindingsNow({
          serviceArgs: args,
          clearFormulaBookkeepingNow,
          bindFormulaNow,
          invalidateFormulaNow,
          isCellIndexMappedNow,
          pruneOrphanedDependencyCells,
        }),
      )
    },
    rebindFormulaCells(candidates, formulaChangedCount) {
      return formulaBindingEffect('Failed to rebind formula cells', () => rebindFormulaCellsNow(candidates, formulaChangedCount))
    },
    rebindDefinedNameDependents(names, formulaChangedCount) {
      return formulaBindingEffect('Failed to rebind defined name dependents', () =>
        rebindTrackedDependentsNow(
          args.reverseState.reverseDefinedNameEdges,
          names.map((name) => normalizeDefinedName(name)),
          formulaChangedCount,
        ),
      )
    },
    rebindTableDependents(tableNames, formulaChangedCount) {
      return formulaBindingEffect('Failed to rebind table dependents', () =>
        rebindTrackedDependentsNow(args.reverseState.reverseTableEdges, tableNames, formulaChangedCount),
      )
    },
    rebindFormulasForSheet(sheetName, formulaChangedCount, candidates) {
      return formulaBindingEffect('Failed to rebind formulas for sheet', () =>
        rebindFormulasForSheetNow(sheetName, formulaChangedCount, candidates),
      )
    },
    bindFormulaNow,
    bindPreparedFormulaNow,
    rewriteFormulaSourcePreservingBindingNow,
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
      formulaFamilyIndex.deferRebuildNow()
    },
    deferFormulaFamilyIndexRunsNow(runs) {
      formulaFamilyIndex.deferRunsNow(runs)
    },
    deferFormulaInstanceTableRebuildNow: rebuildFormulaInstancesNow,
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
      return formulaFamilyIndex.countSheetMembersNow(sheetId)
    },
    canUseFormulaFamilyIndexNow() {
      return formulaFamilyIndex.canUseNow()
    },
    isFormulaFamilyIndexReadyNow() {
      return formulaFamilyIndex.isReadyNow()
    },
    tryDeferFormulaFamilyStructuralSourceTransformsNow(sheetId, transform, canDeferCellIndex) {
      return formulaFamilyIndex.tryDeferStructuralSourceTransformsNow(sheetId, transform, canDeferCellIndex)
    },
    forEachFormulaFamilyNow(fn) {
      formulaFamilyIndex.forEachFamilyNow(fn)
    },
    setFormulaFamilyStructuralSourceTransformNow(familyId, transform) {
      formulaFamilyIndex.setStructuralSourceTransformNow(familyId, transform)
    },
    getFormulaFamilyStructuralSourceTransformNow(cellIndex) {
      return formulaFamilyIndex.getStructuralSourceTransformNow(cellIndex)
    },
    hasFormulaFamilyStructuralSourceTransformsNow() {
      return formulaFamilyIndex.hasStructuralSourceTransformsNow()
    },
    consumeFormulaFamilyStructuralSourceTransformsNow() {
      return formulaFamilyIndex.consumeStructuralSourceTransformsNow()
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
      formulaFamilyIndex.ensureNow()
      return args.formulaFamilies.getStats()
    },
  }
}
