import { Effect } from 'effect'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { errorValue } from '../../engine-value-utils.js'
import { mapStructuralAxisIndex, mapStructuralAxisInterval, structuralTransformForOp } from '../../engine-structural-utils.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { EngineStructureError } from '../errors.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { captureAxisRangeCellState, captureSheetCellState, shouldCaptureStoredCell } from './structure-cell-state.js'
import {
  dependencyTouchesSheet,
  rangeDependencyAxisAffected,
  runtimeDirectRangeAxisAffected,
  structuralDirectAggregateRewritePreservesValue,
  structuralRewritePreservesBinding,
  structuralRewritePreservesValue,
} from './structure-formula-rewrite-guards.js'
import {
  rewriteFormulaFromTemplate,
  rewriteFormulaSourceFallback,
  rewriteStructuralFormulaCompiled,
  structuralRewritePreservesDirectCellDependencies,
  type StructuralFormulaRewriteCache,
} from './structure-formula-rewrite.js'
import { rewriteDefinedNamesForStructuralTransform, rewriteWorkbookMetadataForStructuralTransform } from './structure-metadata-rewrite.js'
import {
  clearPivotOutputsForSheet,
  clearRemovedCellRuntimeState,
  clearSpillMetadataForSheet,
  collectStructuralRangeDependencies,
  isCellIndexMapped,
  structuralAxisIndexAffected,
} from './structure-runtime-cleanup.js'
import {
  canDeferSimpleDeleteRefErrorFormulaSource,
  canDeferSimpleDeleteStructuralFormulaSource,
  canDeferSimpleStructuralFormulaSource,
} from './structure-formula-source-deferral.js'
import { materializeDeferredStructuralFormulaSources as materializeDeferredStructuralFormulaSourcesNow } from './structure-deferred-formula-sources.js'
import type { CreateEngineStructureServiceArgs, EngineStructureService, StructuralFormulaRebindInput } from './structure-service-types.js'

export type {
  CreateEngineStructureServiceArgs,
  EngineStructureService,
  EngineStructureState,
  StructuralAxisOp,
  StructuralFormulaRebindInput,
} from './structure-service-types.js'

export function createEngineStructureService(args: CreateEngineStructureServiceArgs): EngineStructureService {
  let hasDeferredStructuralFormulaSources = false

  const resolveStructuralFormulaRebindInputs = (argsForResolve: {
    readonly formulaCellIndices: readonly number[]
    readonly sheetName: string
    readonly transform: StructuralAxisTransform
    readonly transaction: StructuralTransaction
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
    readonly ownerPositions: ReadonlyMap<number, { sheetName: string; row: number; col: number }>
    readonly precomputedDirectAggregateValueCellIndices: readonly number[]
  }) => {
    const inputs: StructuralFormulaRebindInput[] = []
    const preservedCellIndices: number[] = []
    const templateRewriteCache: StructuralFormulaRewriteCache = new Map()
    const remappedCellsByIndex = new Map(argsForResolve.transaction.remappedCells.map((entry) => [entry.cellIndex, entry] as const))
    const precomputedDirectAggregateValueCellIndices = new Set(argsForResolve.precomputedDirectAggregateValueCellIndices)
    argsForResolve.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const previousOwnerPosition = argsForResolve.ownerPositions.get(cellIndex)
      if (!previousOwnerPosition) {
        return
      }
      const ownerSheetName = previousOwnerPosition.sheetName
      const touchesChangedName = formula.compiled.symbolicNames.some((name) =>
        argsForResolve.changedDefinedNames.has(normalizeDefinedName(name)),
      )
      const touchesChangedTable = formula.compiled.symbolicTables.some((name) => argsForResolve.changedTableNames.has(name))
      const touchesTargetSheetDependency = formula.compiled.deps.some((dependency) =>
        dependencyTouchesSheet(dependency, argsForResolve.sheetName),
      )
      const shouldBypassTemplateStructuralRewrite = ownerSheetName !== argsForResolve.sheetName && touchesTargetSheetDependency
      const representative = remappedCellsByIndex.get(cellIndex)
      const previousOwnerRow = representative?.fromRow ?? previousOwnerPosition.row
      const previousOwnerCol = representative?.fromCol ?? previousOwnerPosition.col
      const ownerRow =
        representative?.toRow ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'row'
          ? mapStructuralAxisIndex(previousOwnerRow, argsForResolve.transform)
          : previousOwnerRow)
      const ownerCol =
        representative?.toCol ??
        (ownerSheetName === argsForResolve.sheetName && argsForResolve.transform.axis === 'column'
          ? mapStructuralAxisIndex(previousOwnerCol, argsForResolve.transform)
          : previousOwnerCol)
      if (ownerRow === undefined || ownerCol === undefined) {
        return
      }
      if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(args, formula, argsForResolve.transform)) {
        formula.structuralSourceTransform = {
          ownerSheetName,
          targetSheetName: argsForResolve.sheetName,
          transform: argsForResolve.transform,
          preservesValue: true,
        }
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.push(cellIndex)
        return
      }
      const templateRewrite =
        !touchesChangedName &&
        !touchesChangedTable &&
        !shouldBypassTemplateStructuralRewrite &&
        formula.templateId !== undefined &&
        previousOwnerRow !== undefined &&
        previousOwnerCol !== undefined
          ? rewriteFormulaFromTemplate(
              templateRewriteCache,
              formula,
              {
                templateId: formula.templateId,
                ownerSheetName,
                targetSheetName: argsForResolve.sheetName,
                representativeRow: previousOwnerRow,
                representativeCol: previousOwnerCol,
                ownerRow,
                ownerCol,
              },
              argsForResolve.sheetName,
              argsForResolve.transform,
            )
          : undefined
      const compiledRewrite =
        templateRewrite === undefined
          ? rewriteStructuralFormulaCompiled(formula, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform)
          : undefined
      const rewritten = !touchesChangedName && !touchesChangedTable ? (compiledRewrite ?? templateRewrite) : compiledRewrite
      if (!rewritten) {
        if (!touchesChangedName && !touchesChangedTable && formula.directAggregate !== undefined) {
          return
        }
        const canReuseCompiled =
          formula.compiled.symbolicNames.length === 0 &&
          formula.compiled.symbolicTables.length === 0 &&
          formula.compiled.symbolicSpills.length === 0
        inputs.push(
          canReuseCompiled
            ? {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: formula.source,
                compiled: formula.compiled,
                ...(formula.templateId === undefined ? {} : { templateId: formula.templateId }),
              }
            : {
                cellIndex,
                ownerSheetName,
                ownerRow,
                ownerCol,
                source: formula.source,
              },
        )
        return
      }
      if (touchesChangedName || touchesChangedTable) {
        inputs.push({
          cellIndex,
          ownerSheetName,
          ownerRow,
          ownerCol,
          source: rewriteFormulaSourceFallback(formula.source, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform),
        })
        return
      }
      const preservesDirectCellDependencies = structuralRewritePreservesDirectCellDependencies(args, formula, rewritten, ownerSheetName)
      const preservesBinding =
        structuralRewritePreservesBinding(
          formula,
          rewritten,
          formula.rangeDependencies.every((rangeIndex) => args.state.ranges.getFormulaMembersView(rangeIndex).length === 0),
        ) || preservesDirectCellDependencies
      const preservesValue =
        precomputedDirectAggregateValueCellIndices.has(cellIndex) ||
        structuralRewritePreservesValue(formula, rewritten, argsForResolve.transform) ||
        structuralDirectAggregateRewritePreservesValue(formula, rewritten, argsForResolve.transform)
      const hasOnlyPlaceholderDirectDependencies =
        formula.dependencyIndices.length > 0 &&
        !formula.dependencyIndices.every((dependencyCellIndex) => shouldCaptureStoredCell(args, dependencyCellIndex))
      const rewrittenDirectDependenciesChanged =
        formula.compiled.deps.length !== rewritten.compiled.deps.length ||
        formula.compiled.deps.some((dependency, index) => dependency !== rewritten.compiled.deps[index])
      const rewrittenPlaceholderDependencyNeedsRebind =
        preservesBinding && rewrittenDirectDependenciesChanged && hasOnlyPlaceholderDirectDependencies
      inputs.push({
        cellIndex,
        ownerSheetName,
        ownerRow,
        ownerCol,
        source: rewritten.source,
        compiled: rewritten.compiled,
        ...(formula.templateId === undefined || rewritten.source !== formula.source ? {} : { templateId: formula.templateId }),
        preservesBinding: preservesBinding && !rewrittenPlaceholderDependencyNeedsRebind,
        preservesValue,
      })
    })
    return { inputs, preservedCellIndices }
  }

  const collectStructuralFormulaImpacts = (argsForImpact: {
    readonly targetSheetId: number | undefined
    readonly transform: StructuralAxisTransform
    readonly sheetName: string
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
  }): {
    formulaCellIndices: number[]
    rebindCellIndices: number[]
    preservedCellIndices: number[]
    precomputedChangedInputCellIndices: number[]
    ownerPositions: Map<number, { sheetName: string; row: number; col: number }>
    precomputedDirectAggregateValueCellIndices: number[]
    directAggregateRetargetCellIndices: number[]
  } => {
    const formulaCellIndices = new Set<number>()
    const rebindCellIndices = new Set<number>()
    const preservedCellIndices = new Set<number>()
    const precomputedChangedInputCellIndices = new Set<number>()
    const candidateCellIndices = new Set<number>()
    const ownerPositions = new Map<number, { sheetName: string; row: number; col: number }>()
    const precomputedDirectAggregateValueCellIndices = new Set<number>()
    const directAggregateRetargetCellIndices = new Set<number>()
    let sharedOwnedPreservingSourceTransform: RuntimeFormula['structuralSourceTransform']
    let deferredOwnedFormulaFamilyMemberCount = 0
    const ownedPreservingSourceTransform = (): NonNullable<RuntimeFormula['structuralSourceTransform']> =>
      (sharedOwnedPreservingSourceTransform ??= {
        ownerSheetName: argsForImpact.sheetName,
        targetSheetName: argsForImpact.sheetName,
        transform: argsForImpact.transform,
        preservesValue: true,
      })
    const tryDeferOwnedFormulaFamilies = (): boolean => {
      if (
        argsForImpact.targetSheetId === undefined ||
        argsForImpact.changedDefinedNames.size > 0 ||
        argsForImpact.changedTableNames.size > 0 ||
        argsForImpact.transform.kind === 'delete' ||
        argsForImpact.transform.axis !== 'column'
      ) {
        return false
      }
      const ownedFormulaCount = args.countFormulaSheetMembers(argsForImpact.targetSheetId)
      if (ownedFormulaCount === 0) {
        return false
      }
      const familyIds: number[] = []
      let familyMemberCount = 0
      let canDeferFamilies = true
      args.forEachFormulaFamily((family) => {
        if (!canDeferFamilies || family.sheetId !== argsForImpact.targetSheetId) {
          return
        }
        const representativeCellIndex = family.runs.find((run) => run.cellIndices.length > 0)?.cellIndices[0]
        const representative = representativeCellIndex === undefined ? undefined : args.state.formulas.get(representativeCellIndex)
        if (!representative || !canDeferSimpleStructuralFormulaSource(args, representative, argsForImpact.transform)) {
          canDeferFamilies = false
          return
        }
        familyIds.push(family.id)
        family.runs.forEach((run) => {
          familyMemberCount += run.cellIndices.length
        })
      })
      if (!canDeferFamilies || familyIds.length === 0 || familyMemberCount !== ownedFormulaCount) {
        return false
      }
      const transform = ownedPreservingSourceTransform()
      familyIds.forEach((familyId) => {
        args.setFormulaFamilyStructuralSourceTransform(familyId, transform)
      })
      hasDeferredStructuralFormulaSources = true
      deferredOwnedFormulaFamilyMemberCount = ownedFormulaCount
      return true
    }
    const canSkipOwnedDirectAggregateCandidate = (cellIndex: number): boolean => {
      if (argsForImpact.changedDefinedNames.size > 0 || argsForImpact.changedTableNames.size > 0) {
        return false
      }
      if (argsForImpact.targetSheetId === undefined) {
        return false
      }
      const formula = args.state.formulas.get(cellIndex)
      if (!formula?.directAggregate) {
        return false
      }
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      if (!ownerPosition) {
        return false
      }
      const ownerAxisIndex = argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col
      if (structuralAxisIndexAffected(ownerAxisIndex, argsForImpact.transform)) {
        return false
      }
      return !runtimeDirectRangeAxisAffected(
        argsForImpact.targetSheetId,
        argsForImpact.sheetName,
        argsForImpact.transform,
        formula.directAggregate,
      )
    }
    const tryPrecomputeDeletedDirectAggregateValue = (
      cellIndex: number,
      formula: RuntimeFormula,
      ownerPosition: { row: number; col: number },
    ): boolean => {
      if (
        argsForImpact.changedDefinedNames.size > 0 ||
        argsForImpact.changedTableNames.size > 0 ||
        argsForImpact.transform.kind !== 'delete' ||
        argsForImpact.transform.axis !== 'row' ||
        argsForImpact.targetSheetId === undefined
      ) {
        return false
      }
      const directAggregate = formula.directAggregate
      if (!directAggregate || directAggregate.sheetName !== argsForImpact.sheetName || directAggregate.aggregateKind !== 'sum') {
        return false
      }
      if (mapStructuralAxisIndex(ownerPosition.row, argsForImpact.transform) === undefined) {
        return false
      }
      const overlapStart = Math.max(directAggregate.rowStart, argsForImpact.transform.start)
      const overlapEnd = Math.min(directAggregate.rowEnd, argsForImpact.transform.start + argsForImpact.transform.count - 1)
      if (overlapStart > overlapEnd) {
        return false
      }
      const aggregateSheet = args.state.workbook.getSheet(argsForImpact.sheetName)
      if (!aggregateSheet) {
        return false
      }
      const currentValue = args.state.workbook.cellStore.getValue(cellIndex, () => '')
      if (currentValue.tag !== ValueTag.Number) {
        return false
      }
      let deletedContribution = 0
      for (let row = overlapStart; row <= overlapEnd; row += 1) {
        const memberCellIndex =
          aggregateSheet.structureVersion === 1
            ? aggregateSheet.grid.getPhysical(row, directAggregate.col)
            : aggregateSheet.grid.get(row, directAggregate.col)
        if (memberCellIndex === -1) {
          continue
        }
        const memberValue = args.state.workbook.cellStore.getValue(memberCellIndex, () => '')
        switch (memberValue.tag) {
          case ValueTag.Number:
            deletedContribution += memberValue.value
            break
          case ValueTag.Boolean:
            deletedContribution += memberValue.value ? 1 : 0
            break
          case ValueTag.Empty:
          case ValueTag.String:
            break
          case ValueTag.Error:
            return false
        }
      }
      args.state.workbook.cellStore.setValue(cellIndex, {
        tag: ValueTag.Number,
        value: currentValue.value - deletedContribution,
      })
      precomputedChangedInputCellIndices.add(cellIndex)
      precomputedDirectAggregateValueCellIndices.add(cellIndex)
      return true
    }
    const canRetargetDirectAggregateWithoutFormulaRewrite = (
      formula: RuntimeFormula,
      ownerPosition: { row: number; col: number },
    ): boolean => {
      if (
        argsForImpact.changedDefinedNames.size > 0 ||
        argsForImpact.changedTableNames.size > 0 ||
        argsForImpact.targetSheetId === undefined ||
        argsForImpact.transform.axis !== 'row'
      ) {
        return false
      }
      const directAggregate = formula.directAggregate
      if (!directAggregate || directAggregate.sheetName !== argsForImpact.sheetName) {
        return false
      }
      if (mapStructuralAxisIndex(ownerPosition.row, argsForImpact.transform) === undefined) {
        return false
      }
      return mapStructuralAxisInterval(directAggregate.rowStart, directAggregate.rowEnd, argsForImpact.transform) !== undefined
    }
    const tryQueueDirectAggregateStructuralRetarget = (
      cellIndex: number,
      formula: RuntimeFormula,
      ownerPosition: { row: number; col: number },
    ): boolean => {
      if (!canRetargetDirectAggregateWithoutFormulaRewrite(formula, ownerPosition)) {
        return false
      }
      if (argsForImpact.transform.kind === 'delete' && !tryPrecomputeDeletedDirectAggregateValue(cellIndex, formula, ownerPosition)) {
        return false
      }
      directAggregateRetargetCellIndices.add(cellIndex)
      return true
    }
    const tryDeferOwnedSimpleFormula = (cellIndex: number): boolean => {
      if (argsForImpact.changedDefinedNames.size > 0 || argsForImpact.changedTableNames.size > 0) {
        return false
      }
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return false
      }
      if (canDeferSimpleStructuralFormulaSource(args, formula, argsForImpact.transform)) {
        formula.structuralSourceTransform = ownedPreservingSourceTransform()
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.add(cellIndex)
        return true
      }
      const ownerAxisIndex = args.state.workbook.getCellAxisIndex(cellIndex, argsForImpact.transform.axis)
      if (ownerAxisIndex === undefined || mapStructuralAxisIndex(ownerAxisIndex, argsForImpact.transform) === undefined) {
        return false
      }
      const preservesValue = canDeferSimpleStructuralFormulaSource(args, formula, argsForImpact.transform)
      const preservesBinding =
        preservesValue || canDeferSimpleDeleteStructuralFormulaSource(args, formula, argsForImpact.targetSheetId, argsForImpact.transform)
      const becomesRefError =
        !preservesBinding && canDeferSimpleDeleteRefErrorFormulaSource(args, formula, argsForImpact.targetSheetId, argsForImpact.transform)
      const dependsOnPrecomputedRefError = formula.dependencyIndices.some((dependencyCellIndex) =>
        precomputedChangedInputCellIndices.has(dependencyCellIndex),
      )
      if (!preservesBinding && !becomesRefError && !dependsOnPrecomputedRefError) {
        return false
      }
      formula.structuralSourceTransform = {
        ownerSheetName: argsForImpact.sheetName,
        targetSheetName: argsForImpact.sheetName,
        transform: argsForImpact.transform,
        preservesValue,
      }
      hasDeferredStructuralFormulaSources = true
      if (preservesValue) {
        preservedCellIndices.add(cellIndex)
      } else if (becomesRefError || dependsOnPrecomputedRefError) {
        args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Ref))
        precomputedChangedInputCellIndices.add(cellIndex)
      } else {
        formulaCellIndices.add(cellIndex)
      }
      return true
    }
    const deferredOwnedFormulaFamilies = tryDeferOwnedFormulaFamilies()
    if (!deferredOwnedFormulaFamilies) {
      args.forEachFormulaCellOwnedBySheet(argsForImpact.sheetName, (cellIndex) => {
        if (tryDeferOwnedSimpleFormula(cellIndex)) {
          return
        }
        const formula = args.state.formulas.get(cellIndex)
        const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
        if (
          ownerPosition &&
          mapStructuralAxisIndex(
            argsForImpact.transform.axis === 'row' ? ownerPosition.row : ownerPosition.col,
            argsForImpact.transform,
          ) === undefined
        ) {
          return
        }
        if (formula && ownerPosition && tryQueueDirectAggregateStructuralRetarget(cellIndex, formula, ownerPosition)) {
          return
        }
        if (canSkipOwnedDirectAggregateCandidate(cellIndex)) {
          return
        }
        candidateCellIndices.add(cellIndex)
      })
    }
    const ownedFamilyDeferralCoversEveryFormula =
      deferredOwnedFormulaFamilies && deferredOwnedFormulaFamilyMemberCount === args.state.formulas.size
    if (!ownedFamilyDeferralCoversEveryFormula) {
      args.collectFormulaCellsReferencingSheet(argsForImpact.sheetName).forEach((cellIndex) => {
        if (directAggregateRetargetCellIndices.has(cellIndex)) {
          return
        }
        const formula = args.state.formulas.get(cellIndex)
        if (formula?.structuralSourceTransform !== undefined) {
          return
        }
        candidateCellIndices.add(cellIndex)
      })
    }
    if (argsForImpact.changedDefinedNames.size > 0) {
      args.collectFormulaCellsForDefinedNames([...argsForImpact.changedDefinedNames]).forEach((cellIndex) => {
        candidateCellIndices.add(cellIndex)
      })
    }
    if (argsForImpact.changedTableNames.size > 0) {
      args.collectFormulaCellsForTables([...argsForImpact.changedTableNames]).forEach((cellIndex) => {
        candidateCellIndices.add(cellIndex)
      })
    }
    if (args.state.counters && candidateCellIndices.size > 0) {
      addEngineCounter(args.state.counters, 'structuralFormulaImpactCandidates', candidateCellIndices.size)
    }
    candidateCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      if (formula.structuralSourceTransform !== undefined) {
        return
      }
      if (!isCellIndexMapped(args, cellIndex)) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (!ownerSheetName) {
        return
      }
      const ownerPosition = args.state.workbook.getCellPosition(cellIndex)
      if (!ownerPosition) {
        return
      }
      if (tryQueueDirectAggregateStructuralRetarget(cellIndex, formula, ownerPosition)) {
        return
      }
      ownerPositions.set(cellIndex, { sheetName: ownerSheetName, row: ownerPosition.row, col: ownerPosition.col })
      const formulaValuePrecomputed = tryPrecomputeDeletedDirectAggregateValue(cellIndex, formula, ownerPosition)
      const axisIndex = argsForImpact.transform.axis === 'row' ? ownerPosition?.row : ownerPosition?.col
      const ownerPositionAffected =
        ownerSheetName === argsForImpact.sheetName &&
        axisIndex !== undefined &&
        structuralAxisIndexAffected(axisIndex, argsForImpact.transform)
      const touchesChangedName =
        argsForImpact.changedDefinedNames.size > 0 &&
        formula.compiled.symbolicNames.some((name) => argsForImpact.changedDefinedNames.has(normalizeDefinedName(name)))
      const touchesChangedTable =
        argsForImpact.changedTableNames.size > 0 &&
        formula.compiled.symbolicTables.some((name) => argsForImpact.changedTableNames.has(name))
      if (!touchesChangedName && !touchesChangedTable && canDeferSimpleStructuralFormulaSource(args, formula, argsForImpact.transform)) {
        formula.structuralSourceTransform =
          ownerSheetName === argsForImpact.sheetName
            ? ownedPreservingSourceTransform()
            : {
                ownerSheetName,
                targetSheetName: argsForImpact.sheetName,
                transform: argsForImpact.transform,
                preservesValue: true,
              }
        hasDeferredStructuralFormulaSources = true
        preservedCellIndices.add(cellIndex)
        return
      }
      const dependencyPositionAffected =
        !ownerPositionAffected &&
        argsForImpact.targetSheetId !== undefined &&
        (formula.dependencyIndices.some((dependencyCellIndex) => {
          if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== argsForImpact.targetSheetId) {
            return false
          }
          const dependencyPosition = args.state.workbook.getCellPosition(dependencyCellIndex)
          const dependencyAxisIndex = argsForImpact.transform.axis === 'row' ? dependencyPosition?.row : dependencyPosition?.col
          return dependencyAxisIndex !== undefined && structuralAxisIndexAffected(dependencyAxisIndex, argsForImpact.transform)
        }) ||
          formula.rangeDependencies.some((rangeIndex) =>
            rangeDependencyAxisAffected(args.state.ranges.getDescriptor(rangeIndex), argsForImpact.targetSheetId!, argsForImpact.transform),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directAggregate,
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directCriteria?.aggregateRange,
          ) ||
          formula.directCriteria?.criteriaPairs.some((pair) =>
            runtimeDirectRangeAxisAffected(argsForImpact.targetSheetId, argsForImpact.sheetName, argsForImpact.transform, pair.range),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directLookup?.kind === 'exact' || formula.directLookup?.kind === 'approximate'
              ? {
                  sheetName: formula.directLookup.prepared.sheetName,
                  rowStart: formula.directLookup.prepared.rowStart,
                  rowEnd: formula.directLookup.prepared.rowEnd,
                  col: formula.directLookup.prepared.col,
                }
              : formula.directLookup?.kind === 'exact-uniform-numeric' || formula.directLookup?.kind === 'approximate-uniform-numeric'
                ? {
                    sheetName: formula.directLookup.sheetName,
                    rowStart: formula.directLookup.rowStart,
                    rowEnd: formula.directLookup.rowEnd,
                    col: formula.directLookup.col,
                  }
                : undefined,
          ))
      const touchesSheetDependency =
        !ownerPositionAffected &&
        !dependencyPositionAffected &&
        formula.compiled.deps.some((dependency) => dependencyTouchesSheet(dependency, argsForImpact.sheetName))
      if (!ownerPositionAffected && !dependencyPositionAffected && !touchesSheetDependency && !touchesChangedName && !touchesChangedTable) {
        return
      }
      formulaCellIndices.add(cellIndex)
      if (ownerPositionAffected || dependencyPositionAffected || touchesSheetDependency || touchesChangedName || touchesChangedTable) {
        rebindCellIndices.add(cellIndex)
      }
      if (formulaValuePrecomputed) {
        rebindCellIndices.add(cellIndex)
      }
    })
    return {
      formulaCellIndices: [...formulaCellIndices],
      rebindCellIndices: [...rebindCellIndices],
      preservedCellIndices: [...preservedCellIndices],
      precomputedChangedInputCellIndices: [...precomputedChangedInputCellIndices],
      ownerPositions,
      precomputedDirectAggregateValueCellIndices: [...precomputedDirectAggregateValueCellIndices],
      directAggregateRetargetCellIndices: [...directAggregateRetargetCellIndices],
    }
  }

  const materializeDeferredStructuralFormulaSources = (): void => {
    hasDeferredStructuralFormulaSources = materializeDeferredStructuralFormulaSourcesNow(args, hasDeferredStructuralFormulaSources)
  }

  return {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => captureSheetCellState(args, sheetName),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture sheet cell state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(args, sheetName, 'row', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture row state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(args, sheetName, 'column', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture column state for ${sheetName}`,
            cause,
          }),
      })
    },
    materializeDeferredStructuralFormulaSources() {
      return Effect.try({
        try: () => materializeDeferredStructuralFormulaSources(),
        catch: (cause) =>
          new EngineStructureError({
            message: 'Failed to materialize deferred structural formula sources',
            cause,
          }),
      })
    },
    applyStructuralAxisOp(op) {
      return Effect.try({
        try: () => {
          materializeDeferredStructuralFormulaSources()
          const transform = structuralTransformForOp(op)
          const sheetName = op.sheetName
          const targetSheetId = args.state.workbook.getSheet(sheetName)?.id

          clearPivotOutputsForSheet(args, sheetName)
          const changedDefinedNames = rewriteDefinedNamesForStructuralTransform(args, sheetName, transform)
          const { changedTableNames } = rewriteWorkbookMetadataForStructuralTransform(args, sheetName, transform)
          const impactedFormulas = collectStructuralFormulaImpacts({
            targetSheetId,
            transform,
            sheetName,
            changedDefinedNames,
            changedTableNames,
          })

          const transaction =
            args.state.workbook.planStructuralAxisTransform(sheetName, transform) ??
            (() => {
              throw new Error(`Missing sheet for structural op: ${sheetName}`)
            })()

          switch (op.kind) {
            case 'insertRows':
              args.state.workbook.insertRows(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteRows':
              args.state.workbook.deleteRows(sheetName, op.start, op.count)
              break
            case 'moveRows':
              args.state.workbook.moveRows(sheetName, op.start, op.count, op.target)
              break
            case 'insertColumns':
              args.state.workbook.insertColumns(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteColumns':
              args.state.workbook.deleteColumns(sheetName, op.start, op.count)
              break
            case 'moveColumns':
              args.state.workbook.moveColumns(sheetName, op.start, op.count, op.target)
              break
          }

          args.state.workbook.applyPlannedStructuralTransaction(transaction)

          const structuralRangeDependencies = collectStructuralRangeDependencies(args, impactedFormulas.formulaCellIndices)

          let hadCycleFormulas: boolean | undefined
          const hasCycleFormulas = (): boolean => {
            if (hadCycleFormulas !== undefined) {
              return hadCycleFormulas
            }
            if (args.state.counters) {
              addEngineCounter(args.state.counters, 'cycleFormulaScans')
            }
            let found = false
            args.state.formulas.forEach((_formula, cellIndex) => {
              if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
                found = true
              }
            })
            hadCycleFormulas = found
            return found
          }
          const removedFormulaCellIndices = transaction.removedCellIndices.filter((cellIndex) => args.state.formulas.has(cellIndex))
          const removedFormulaCellIndexSet = new Set<number>(removedFormulaCellIndices)
          const removedCycleFormulaCount = removedFormulaCellIndices.filter(
            (cellIndex) => ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0,
          ).length
          transaction.removedCellIndices.forEach((cellIndex) => {
            clearRemovedCellRuntimeState(args, cellIndex)
          })

          clearSpillMetadataForSheet(args, sheetName)
          args.retargetRangeDependencies(transaction, structuralRangeDependencies)
          const directRetargetedFormulaCellIndices: number[] = []
          const directRetargetedPreservedFormulaCellIndices: number[] = []
          const precomputedDirectAggregateValueCellIndices = new Set(impactedFormulas.precomputedDirectAggregateValueCellIndices)
          impactedFormulas.directAggregateRetargetCellIndices.forEach((cellIndex) => {
            if (!isCellIndexMapped(args, cellIndex)) {
              return
            }
            const retargeted = args.retargetDirectAggregateFormulaForStructuralTransform(
              {
                cellIndex,
                ownerSheetName: sheetName,
                ownerRow: 0,
                ownerCol: 0,
                source: '',
                preservesValue: true,
              },
              sheetName,
              transform,
            )
            if (!retargeted) {
              return
            }
            directRetargetedFormulaCellIndices.push(cellIndex)
            directRetargetedPreservedFormulaCellIndices.push(cellIndex)
            hasDeferredStructuralFormulaSources = true
          })
          const rebindResolution = resolveStructuralFormulaRebindInputs({
            formulaCellIndices: impactedFormulas.rebindCellIndices.filter((cellIndex) => isCellIndexMapped(args, cellIndex)),
            sheetName,
            transform,
            transaction,
            changedDefinedNames,
            changedTableNames,
            ownerPositions: impactedFormulas.ownerPositions,
            precomputedDirectAggregateValueCellIndices: [...precomputedDirectAggregateValueCellIndices],
          })
          const rebindInputs = rebindResolution.inputs
          const remainingRebindInputs: StructuralFormulaRebindInput[] = []
          rebindInputs.forEach((input) => {
            const formula = args.state.formulas.get(input.cellIndex)
            const directAggregateRetargeted =
              input.preservesBinding === true &&
              formula?.directAggregate !== undefined &&
              args.retargetDirectAggregateFormulaForStructuralTransform(input, sheetName, transform)
            if (directAggregateRetargeted) {
              hasDeferredStructuralFormulaSources = true
            }
            if (
              directAggregateRetargeted ||
              (input.preservesBinding === true &&
                formula?.directAggregate !== undefined &&
                input.compiled !== undefined &&
                args.rewriteFormulaCompiledPreservingBinding(input))
            ) {
              directRetargetedFormulaCellIndices.push(input.cellIndex)
              if (input.preservesValue) {
                directRetargetedPreservedFormulaCellIndices.push(input.cellIndex)
              }
              return
            }
            remainingRebindInputs.push(input)
          })
          if (args.state.counters && remainingRebindInputs.length > 0) {
            addEngineCounter(args.state.counters, 'structuralFormulaRebindInputs', remainingRebindInputs.length)
          }
          const formulaCellIndices = impactedFormulas.formulaCellIndices.filter((cellIndex) => isCellIndexMapped(args, cellIndex))
          const onlyDirectAggregateFormulaCells =
            formulaCellIndices.length > 0 &&
            formulaCellIndices.every((cellIndex) => args.state.formulas.get(cellIndex)?.directAggregate !== undefined)
          args.rebindFormulaCells(remainingRebindInputs)
          const reboundFormulaCellIndices = new Set([
            ...directRetargetedFormulaCellIndices,
            ...remainingRebindInputs.map((input) => input.cellIndex),
          ])
          const preservedFormulaCellIndices = new Set([
            ...impactedFormulas.preservedCellIndices,
            ...rebindResolution.preservedCellIndices,
            ...directRetargetedPreservedFormulaCellIndices,
            ...remainingRebindInputs.filter((input) => input.preservesValue).map((input) => input.cellIndex),
          ])
          const lostSurvivingFormulaCells = impactedFormulas.formulaCellIndices.some(
            (cellIndex) =>
              !reboundFormulaCellIndices.has(cellIndex) &&
              !isCellIndexMapped(args, cellIndex) &&
              !removedFormulaCellIndexSet.has(cellIndex),
          )
          const hasNonPreservedRebind = remainingRebindInputs.some((input) => input.preservesBinding !== true)
          const needsDeleteAcyclicRebindCheck =
            transform.kind === 'delete' &&
            changedDefinedNames.size === 0 &&
            changedTableNames.size === 0 &&
            (hasNonPreservedRebind || lostSurvivingFormulaCells)
          const deleteOnlyAcyclicRebind = needsDeleteAcyclicRebindCheck && !hasCycleFormulas()
          const topologyChanged = removedFormulaCellIndices.length > 0 || hasNonPreservedRebind || lostSurvivingFormulaCells
          const graphRefreshRequired =
            ((hasNonPreservedRebind || lostSurvivingFormulaCells) && !onlyDirectAggregateFormulaCells && !deleteOnlyAcyclicRebind) ||
            removedCycleFormulaCount > 0
          return {
            transaction,
            changedCellIndices: [...transaction.removedCellIndices],
            precomputedChangedInputCellIndices: impactedFormulas.precomputedChangedInputCellIndices,
            formulaCellIndices: formulaCellIndices.filter((cellIndex) => !preservedFormulaCellIndices.has(cellIndex)),
            topologyChanged,
            graphRefreshRequired,
          }
        },
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to apply structural operation ${op.kind}`,
            cause,
          }),
      })
    },
  }
}
