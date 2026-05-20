import { FormulaMode } from '@bilig/protocol'
import { makeCellEntity, makeExactLookupColumnEntity, makeRangeEntity, makeSortedLookupColumnEntity } from '../../entity-ids.js'
import { spillDependencyKeyFromRef, tableDependencyKey } from '../../engine-metadata-utils.js'
import type { RuntimeFormula } from '../runtime-state.js'
import {
  aggregateColumnDependencyKey,
  appendDirectAggregateColumnReverseEdges,
  appendTrackedReverseEdge,
  appendUnindexedAggregateColumnReverseEdge,
  directCriteriaAggregateColumn,
  directLookupColumnInfo,
  directRegionIdsForFormula,
} from './formula-binding-dependency-helpers.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import type { PreparedFormulaBinding } from './formula-binding-prepare.js'
import type {
  BindPreparedFormulaOptions,
  CreateEngineFormulaBindingServiceArgs,
  FormulaOwnerPosition,
} from './formula-binding-service-types.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'

export function installFreshFormulaBindingNow(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly appendReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly appendDefinedNameReverseEdge: (name: string, dependentCellIndex: number) => void
  readonly trackFormulaSheetIndexes: (cellIndex: number, ownerSheetName: string, compiled: RuntimeFormula['compiled']) => void
  readonly updateVolatileFormulaIndex: (cellIndex: number, formula: RuntimeFormula | undefined) => void
  readonly recordFormulaInstanceNow: (
    cellIndex: number,
    source: string,
    templateId: number | undefined,
    ownerPosition?: FormulaOwnerPosition,
  ) => void
  readonly registerFormulaFamilyNow: (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition) => void
  readonly primeLookupCandidatesNow: (
    ownerSheetName: string,
    directLookup: PreparedFormulaBinding['directLookup'],
    indexedExactLookupCandidates: PreparedFormulaBinding['indexedExactLookupCandidates'],
    directApproximateLookupCandidates: PreparedFormulaBinding['directApproximateLookupCandidates'],
  ) => void
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly source: string
  readonly prepared: PreparedFormulaBinding
  readonly options?: BindPreparedFormulaOptions
}): void {
  const serviceArgs = args.serviceArgs
  const cellStore = serviceArgs.state.workbook.cellStore
  const sheetId = cellStore.sheetIds[args.cellIndex]
  const sheet =
    args.options?.ownerPosition === undefined && sheetId !== undefined ? serviceArgs.state.workbook.getSheetById(sheetId) : undefined
  const physicalOwnerPosition =
    args.options?.ownerPosition ??
    (sheet && sheet.structureVersion === 1
      ? {
          sheetName: args.ownerSheetName,
          row: cellStore.rows[args.cellIndex] ?? 0,
          col: cellStore.cols[args.cellIndex] ?? 0,
        }
      : undefined)
  const dependencyEntities = serviceArgs.edgeArena.replace(serviceArgs.edgeArena.empty(), args.prepared.dependencies.dependencyEntities)
  const runtimeFormula: RuntimeFormula = {
    cellIndex: args.cellIndex,
    formulaSlotId: 0,
    planId: args.prepared.plan.id,
    templateId: args.prepared.templateId,
    source: args.source,
    compiled: args.prepared.plan.compiled,
    plan: args.prepared.plan,
    dependencyIndices: args.prepared.dependencies.dependencyIndices,
    dependencyEntities,
    rangeDependencies: args.prepared.dependencies.rangeDependencies,
    graphRangeDependencies: args.prepared.dependencies.graphRangeDependencies,
    runtimeProgram: args.prepared.runtimeProgram,
    constants: args.prepared.compiled.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: args.prepared.runtimeProgram.length,
    constNumberOffset: 0,
    constNumberLength: args.prepared.compiled.constants.length,
    rangeListOffset: 0,
    rangeListLength: args.prepared.dependencies.rangeDependencies.length,
    directLookup: args.prepared.directLookup,
    directAggregate: args.prepared.directAggregate,
    directScalar: args.prepared.directScalar,
    directCriteria: args.prepared.directCriteria,
    ...(args.options?.preserveCachedValueOnFullRecalc === true ? { preserveCachedValueOnFullRecalc: true } : {}),
    inlineScalarFastPlanKind: args.prepared.inlineScalarFastPlanKind,
    inlineScalarPlanCellIndices: args.prepared.inlineScalarPlanCellIndices,
  }
  if (args.prepared.inlineScalarArithmeticDeltaCoefficients !== undefined) {
    runtimeFormula.inlineScalarArithmeticDeltaCoefficients = args.prepared.inlineScalarArithmeticDeltaCoefficients
  }
  if (args.prepared.inlineScalarFastPlanStringIds !== undefined) {
    runtimeFormula.inlineScalarFastPlanStringIds = args.prepared.inlineScalarFastPlanStringIds
  }
  const formulaSlotId = serviceArgs.state.formulas.set(args.cellIndex, runtimeFormula)
  runtimeFormula.formulaSlotId = formulaSlotId
  args.updateVolatileFormulaIndex(args.cellIndex, runtimeFormula)
  const col = physicalOwnerPosition?.col ?? serviceArgs.state.workbook.getCellPosition(args.cellIndex)?.col
  if (sheetId !== undefined) {
    args.formulaMemberCounts.increment(sheetId, col)
  }
  markFormulaCellBound(serviceArgs.state.workbook.cellStore, args.cellIndex, runtimeFormula.compiled.mode)
  if (args.options?.deferFormulaInstanceRegistration !== true) {
    args.recordFormulaInstanceNow(args.cellIndex, args.source, args.prepared.templateId, physicalOwnerPosition)
  }
  if (args.options?.deferFamilyRegistration !== true) {
    args.registerFormulaFamilyNow(args.cellIndex, runtimeFormula, physicalOwnerPosition)
  }

  for (let rangeCursor = 0; rangeCursor < args.prepared.dependencies.newRangeCount; rangeCursor += 1) {
    const rangeIndex = args.prepared.dependencies.newRangeIndices[rangeCursor]!
    const dependencySources = serviceArgs.state.ranges.getDependencySourceEntities(rangeIndex)
    const rangeEntity = makeRangeEntity(rangeIndex)
    for (let index = 0; index < dependencySources.length; index += 1) {
      args.appendReverseEdge(dependencySources[index]!, rangeEntity)
    }
  }
  const formulaEntity = makeCellEntity(args.cellIndex)
  appendFreshFormulaDependencyReverseEdges(args.prepared.dependencies.dependencyEntities, formulaEntity, args.appendKnownUniqueReverseEdge)
  runtimeFormula.compiled.symbolicNames.forEach((name) => {
    args.appendDefinedNameReverseEdge(name, args.cellIndex)
  })
  runtimeFormula.compiled.symbolicTables.forEach((name) => {
    appendTrackedReverseEdge(serviceArgs.reverseState.reverseTableEdges, tableDependencyKey(name), args.cellIndex)
  })
  runtimeFormula.compiled.symbolicSpills.forEach((key) => {
    appendTrackedReverseEdge(
      serviceArgs.reverseState.reverseSpillEdges,
      spillDependencyKeyFromRef(
        key,
        serviceArgs.state.workbook.getSheetNameById(serviceArgs.state.workbook.cellStore.sheetIds[args.cellIndex]!),
      ),
      args.cellIndex,
    )
  })
  if (args.prepared.directLookup) {
    const lookupInfo = directLookupColumnInfo(args.prepared.directLookup)
    const lookupSheet = serviceArgs.state.workbook.getSheet(lookupInfo.sheetName)
    if (lookupSheet) {
      const lookupEntity = lookupInfo.isExact
        ? makeExactLookupColumnEntity(lookupSheet.id, lookupInfo.col)
        : makeSortedLookupColumnEntity(lookupSheet.id, lookupInfo.col)
      args.appendKnownUniqueReverseEdge(lookupEntity, formulaEntity)
    }
  }
  const directCriteriaAggregate = directCriteriaAggregateColumn(args.prepared.directCriteria)
  if (directCriteriaAggregate) {
    const aggregateSheet = serviceArgs.state.workbook.getSheet(directCriteriaAggregate.sheetName)
    if (aggregateSheet) {
      appendUnindexedAggregateColumnReverseEdge(
        serviceArgs.reverseState.reverseAggregateColumnEdges,
        aggregateColumnDependencyKey(aggregateSheet.id, directCriteriaAggregate.col),
        args.cellIndex,
      )
    }
  }
  appendDirectAggregateColumnReverseEdges(
    serviceArgs.reverseState.reverseAggregateColumnEdges,
    serviceArgs.state.workbook,
    runtimeFormula.directAggregate,
    args.cellIndex,
  )
  args.trackFormulaSheetIndexes(args.cellIndex, args.ownerSheetName, runtimeFormula.compiled)
  if (runtimeFormula.directAggregate !== undefined || runtimeFormula.directCriteria !== undefined) {
    serviceArgs.regionGraph.replaceFormulaSubscriptions(args.cellIndex, directRegionIdsForFormula(runtimeFormula))
  }
  if (args.prepared.compiled.mode === FormulaMode.WasmFastPath && args.prepared.runtimeProgram.length > 0) {
    serviceArgs.scheduleWasmProgramSync()
  }

  args.primeLookupCandidatesNow(
    args.ownerSheetName,
    args.prepared.directLookup,
    args.prepared.indexedExactLookupCandidates,
    args.prepared.directApproximateLookupCandidates,
  )
}

export function appendFreshFormulaDependencyReverseEdges(
  dependencyEntities: Uint32Array,
  formulaEntity: number,
  appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void,
): void {
  let largerSeen: Set<number> | undefined
  for (let index = 0; index < dependencyEntities.length; index += 1) {
    const dependencyEntity = dependencyEntities[index]!
    if (largerSeen) {
      if (largerSeen.has(dependencyEntity)) {
        continue
      }
      largerSeen.add(dependencyEntity)
      appendKnownUniqueReverseEdge(dependencyEntity, formulaEntity)
      continue
    }
    let seen = false
    for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
      if (dependencyEntities[previousIndex] === dependencyEntity) {
        seen = true
        break
      }
    }
    if (seen) {
      continue
    }
    if (index === 8 && dependencyEntities.length > 12) {
      largerSeen = new Set(dependencyEntities.subarray(0, index + 1))
    }
    appendKnownUniqueReverseEdge(dependencyEntity, formulaEntity)
  }
}
