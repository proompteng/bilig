import { Effect } from 'effect'
import { ValueTag } from '@bilig/protocol'
import type { EngineCellMutationRef, EngineFormulaSourceRefs } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeFormula, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { evaluateInitialDirectScalar, evaluateInitialDirectScalarNumber } from './formula-initialization-direct-scalar.js'
import { evaluateInitialDirectFormulas, INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT } from './formula-initialization-direct-formulas.js'
import { canEvaluateInitialPrefixAggregateGroupsNatively } from './formula-initialization-prefix-aggregates.js'
import {
  createDeferredInitialFormulaFamilyRunMap,
  flushDeferredInitialFormulaFamilyRuns,
  noteDeferredFormulaFamilyRunMember as noteDeferredFormulaFamilyRunMemberNow,
  registerDeferredFormulaFamilyRunNow,
  type DeferredInitialFormulaFamilyRunMap,
  type DeferredInitialFormulaFamilyRun,
} from './formula-initialization-family-runs.js'
import {
  flushAlignedFreshFormulaFamilyRuns,
  readAlignedFreshFormulaFamilyRunsFromRefs,
} from './formula-initialization-restored-family-runs.js'
import { createInitialTemplateFormulaResolver } from './formula-initialization-template-resolver.js'
import { createInitialFormulaValueWriter, type InitialFormulaValueWriter } from './formula-initialization-value-writer.js'
import { noteDeferredFormulaInstance, readAlignedFreshFormulaInstancesFromRefs } from './formula-initialization-fresh-instances.js'
import {
  initialFormulaEntryRefAt,
  type InitialFormulaCellIndexList,
  type InitialFormulaEntryRefSource,
  type InitialResolvedFormulaEntry,
} from './formula-initialization-refs.js'
import { createInitialFormulaCellIndexPlan } from './formula-initialization-cell-index-plan.js'
import { recalculateFreshVolatileFormulasAfterInitialMaterialization } from './formula-initialization-volatile-pass.js'
import {
  canEvaluateInitialDirectRuntimeFormula,
  compiledFormulaRequiresWorkbookMetadataBinding,
  hasPendingFormulaDependency,
  mutationErrorMessage,
} from './formula-initialization-predicates.js'
import { rethrowFatalFormulaBindingError } from './formula-binding-error-policy.js'
import {
  createInitialNativeDirectScalarBatch,
  MAX_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE,
  MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE,
} from './formula-initialization-native-direct-scalar.js'
import { createInitialDirectScalarRunChunkCollector } from './formula-initialization-direct-scalar-run-chunks.js'
import { createInitialDirectScalarPreEvaluationTracker } from './formula-initialization-direct-scalar-pre-evaluation.js'
import {
  createInitialNativeDirectScalarRowChainBatch,
  MIN_INITIAL_NATIVE_DIRECT_SCALAR_ROW_CHAIN_BATCH_SIZE,
} from './formula-initialization-native-direct-scalar-row-chain.js'
import {
  createInitialNativeDirectLookupBatch,
  MAX_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE,
  MIN_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE,
} from './formula-initialization-native-direct-lookup.js'
import type {
  EngineFormulaInitializationService,
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'
import { tryBindHydratedFreshDirectScalarFormula } from './formula-initialization-hydrated-direct-scalar.js'
import { tryEvaluateFormulaLeafInlineScalar } from './formula-leaf-inline-scalar-evaluator.js'
import { initializeCachedFormulaSourcesAtNow as initializeCachedFormulaSourcesAtNowUnchecked } from './formula-initialization-cached-formulas.js'

export type {
  EngineFormulaInitializationService,
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'
export function createEngineFormulaInitializationService(args: EngineFormulaInitializationServiceArgs): EngineFormulaInitializationService {
  const sheetNameById = new Map<number, string>()
  const hasCycleMembersNow = (): boolean => {
    addEngineCounter(args.state.counters, 'cycleFormulaScans')
    let found = false
    args.state.formulas.forEach((_formula, cellIndex) => {
      if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
        found = true
      }
    })
    return found
  }
  const resolveSheetName = (sheetId: number): string => {
    const cached = sheetNameById.get(sheetId)
    if (cached !== undefined) {
      return cached
    }
    const sheet = args.state.workbook.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    sheetNameById.set(sheetId, sheet.name)
    return sheet.name
  }

  const noteDeferredFormulaFamilyRunMember = (
    runs: DeferredInitialFormulaFamilyRunMap | undefined,
    prepared: Parameters<typeof noteDeferredFormulaFamilyRunMemberNow>[0]['prepared'],
    runtimeFormula: RuntimeFormula | undefined,
  ): void => noteDeferredFormulaFamilyRunMemberNow({ runs, prepared, runtimeFormula })

  const registerDeferredFormulaFamilyRun = (run: DeferredInitialFormulaFamilyRun): void =>
    registerDeferredFormulaFamilyRunNow({
      run,
      registerFreshFormulaFamilyRun: args.registerFreshFormulaFamilyRun,
      upsertFormulaFamilyRun: args.upsertFormulaFamilyRun,
    })

  const initializeFormulaEntriesNow = <Entry>(
    refs: InitialFormulaEntryRefSource<Entry>,
    potentialNewCells: number | undefined,
    resolveCellIndex: (ref: Entry) => number,
    resolveEntry: (ref: Entry, cellIndex: number) => InitialResolvedFormulaEntry,
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginEvaluationBudget(performance.now())
    try {
      args.checkEvaluationBudget()
      args.beginMutationCollection()
      let hadCycleMembersBefore: boolean | undefined
      const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
      let changedInputCount = 0
      let formulaChangedCount = 0
      let topologyChanged = false
      let compileMs = 0
      let workbookDateSystem: string | undefined
      const resolveWorkbookDateSystem = (): string | undefined =>
        (workbookDateSystem ??= args.state.workbook.getCalculationSettings().dateSystem)
      const reservedNewCells = potentialNewCells ?? refs.length
      const hadExistingFormulas = args.state.formulas.size > 0
      args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
      args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
      args.resetMaterializedCellScratch(reservedNewCells)
      const { targetCellIndices, pendingInitialFormulaCellIndices, pendingFormulaCells, maxTargetCellIndex } =
        createInitialFormulaCellIndexPlan({
          refs,
          hadExistingFormulas,
          resolveCellIndex,
          checkEvaluationBudget: args.checkEvaluationBudget,
        })
      let canAssignTopoInBatch = !hadExistingFormulas
      let needsFreshTopoRebuild = false
      let nextTopoRank = 0
      let orderedPreparedCellIndices: number[] | undefined = hadExistingFormulas ? [] : undefined
      let orderedPreparedCellCount = 0
      let canUseInitialDirectEvaluation = false
      let allPreparedFormulasCanUseInitialDirectEvaluation = true
      let hasInitialPrefixAggregateCandidates = false
      let canUseNativeInitialDirectScalarOverLimit = false
      let canUseNativeInitialDirectLookupOverLimit = false
      let canUseNativeInitialPrefixAggregateOverLimit = false
      let inlineInitialDirectScalarWriter: InitialFormulaValueWriter | undefined
      const nativeInitialDirectScalarRunChunks =
        hadExistingFormulas || refs.length < MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE
          ? undefined
          : createInitialDirectScalarRunChunkCollector({ state: args.state })
      let nativeInitialDirectScalarBatch =
        hadExistingFormulas ||
        refs.length < MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE ||
        refs.length > MAX_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE
          ? undefined
          : createInitialNativeDirectScalarBatch({ state: args.state, capacity: refs.length })
      let nativeInitialDirectScalarRowChainBatch =
        hadExistingFormulas || refs.length < MIN_INITIAL_NATIVE_DIRECT_SCALAR_ROW_CHAIN_BATCH_SIZE || refs.length % 2 !== 0
          ? undefined
          : createInitialNativeDirectScalarRowChainBatch({ state: args.state, capacity: refs.length })
      let nativeInitialDirectScalarCellCount = 0
      let nativeInitialDirectLookupBatch =
        hadExistingFormulas ||
        refs.length < MIN_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE ||
        refs.length > MAX_INITIAL_NATIVE_DIRECT_LOOKUP_BATCH_SIZE
          ? undefined
          : createInitialNativeDirectLookupBatch({ state: args.state, capacity: refs.length })
      let nativeInitialDirectLookupCellCount = 0
      const shouldDeferFormulaFamilyIndex = !hadExistingFormulas && args.deferFormulaFamilyIndexRebuild !== undefined
      const shouldDeferFormulaInstanceTable =
        !hadExistingFormulas && (args.hydrateFreshFormulaInstances !== undefined || args.deferFormulaInstanceTableRebuild !== undefined)
      const deferredFormulaInstances = !hadExistingFormulas && args.hydrateFreshFormulaInstances !== undefined ? [] : undefined
      const alignedFreshFormulaFamilyRuns = readAlignedFreshFormulaFamilyRunsFromRefs({
        refs,
        hadExistingFormulas,
        counters: args.state.counters,
      })
      const canCaptureDeferredFormulaFamilyRuns = !shouldDeferFormulaFamilyIndex || args.deferFormulaFamilyIndexRuns !== undefined
      const deferredFormulaFamilyRuns =
        hadExistingFormulas || alignedFreshFormulaFamilyRuns !== undefined || !canCaptureDeferredFormulaFamilyRuns
          ? undefined
          : createDeferredInitialFormulaFamilyRunMap()
      let freshFormulaChangedBufferMaterialized = hadExistingFormulas

      const materializeOrderedPreparedCellIndices = (): number[] => {
        if (orderedPreparedCellIndices) {
          return orderedPreparedCellIndices
        }
        orderedPreparedCellIndices = []
        for (let index = 0; index < orderedPreparedCellCount; index += 1) {
          orderedPreparedCellIndices.push(targetCellIndices[index]!)
        }
        return orderedPreparedCellIndices
      }
      const pushOrderedPreparedCellIndex = (cellIndex: number): void => {
        if (!hadExistingFormulas && orderedPreparedCellIndices === undefined) {
          orderedPreparedCellCount += 1
          return
        }
        materializeOrderedPreparedCellIndices().push(cellIndex)
        orderedPreparedCellCount += 1
      }
      const noteSkippedOrderedPreparedCellIndex = (): void => {
        if (!hadExistingFormulas && orderedPreparedCellIndices === undefined) {
          materializeOrderedPreparedCellIndices()
        }
      }
      const orderedPreparedCellList = (): InitialFormulaCellIndexList => orderedPreparedCellIndices ?? targetCellIndices
      const inlineInitialDirectScalarCells = createInitialDirectScalarPreEvaluationTracker({
        state: args.state,
        refsLength: refs.length,
        targetCellIndices,
        maxTargetCellIndex,
        orderedPreparedCellList,
        orderedPreparedCellCount: () => orderedPreparedCellCount,
      })
      const materializeFreshFormulaChangedBuffer = (): number => {
        if (hadExistingFormulas || freshFormulaChangedBufferMaterialized) {
          return formulaChangedCount
        }
        const orderedPreparedCells = orderedPreparedCellList()
        for (let index = 0; index < orderedPreparedCellCount; index += 1) {
          args.checkEvaluationBudget()
          formulaChangedCount = args.markFormulaChanged(orderedPreparedCells[index]!, formulaChangedCount)
        }
        freshFormulaChangedBufferMaterialized = true
        return formulaChangedCount
      }
      const logicalFormulaChangedCount = (): number =>
        !hadExistingFormulas && !freshFormulaChangedBufferMaterialized ? orderedPreparedCellCount : formulaChangedCount

      const noteInlineInitialDirectScalarCell = (prepared: { cellIndex: number }, runtimeFormula: RuntimeFormula): void => {
        inlineInitialDirectScalarCells.noteCell(prepared.cellIndex, runtimeFormula)
      }

      const tryInlineInitialDirectScalarEvaluation = (
        prepared: { cellIndex: number; sheetId: number; row: number; col: number },
        runtimeFormula: RuntimeFormula | undefined,
      ): void => {
        if (
          hadExistingFormulas ||
          !canAssignTopoInBatch ||
          !runtimeFormula ||
          runtimeFormula.compiled.volatile ||
          runtimeFormula.compiled.producesSpill
        ) {
          return
        }
        if (runtimeFormula.directScalar !== undefined) {
          const addedToNativeRunChunks =
            nativeInitialDirectScalarRunChunks?.add({
              cellIndex: prepared.cellIndex,
              sheetId: prepared.sheetId,
              row: prepared.row,
              col: prepared.col,
              runtimeFormula,
            }) === true
          let addedToNativeRowChain = false
          if (nativeInitialDirectScalarRowChainBatch) {
            if (nativeInitialDirectScalarRowChainBatch.add(prepared, runtimeFormula)) {
              addedToNativeRowChain = true
            } else {
              nativeInitialDirectScalarRowChainBatch = undefined
            }
          }
          if (nativeInitialDirectScalarBatch) {
            if (nativeInitialDirectScalarBatch.add(prepared, runtimeFormula.directScalar)) {
              nativeInitialDirectScalarCellCount += 1
              return
            }
            nativeInitialDirectScalarBatch = undefined
            nativeInitialDirectScalarCellCount = 0
          }
          if (addedToNativeRowChain) {
            return
          }
          if (addedToNativeRunChunks) {
            return
          }
          const numericValue = evaluateInitialDirectScalarNumber(args.state, runtimeFormula.directScalar)
          inlineInitialDirectScalarWriter ??= createInitialFormulaValueWriter(args)
          if (numericValue !== undefined) {
            inlineInitialDirectScalarWriter.writeNumberAt(prepared.cellIndex, prepared.sheetId, prepared.col, numericValue)
            noteInlineInitialDirectScalarCell(prepared, runtimeFormula)
            return
          }
          const fallbackValue = evaluateInitialDirectScalar(args.state, runtimeFormula.directScalar)
          if (fallbackValue !== undefined) {
            inlineInitialDirectScalarWriter.writeValueAt(prepared.cellIndex, prepared.sheetId, prepared.col, fallbackValue)
            noteInlineInitialDirectScalarCell(prepared, runtimeFormula)
            return
          }
        }
        if (runtimeFormula.directLookup !== undefined && nativeInitialDirectLookupBatch) {
          if (nativeInitialDirectLookupBatch.add(prepared, runtimeFormula.directLookup)) {
            nativeInitialDirectLookupCellCount += 1
            return
          }
          nativeInitialDirectLookupBatch = undefined
          nativeInitialDirectLookupCellCount = 0
        }
        if (runtimeFormula.inlineScalarFastPlanKind !== undefined) {
          const inlineValue = tryEvaluateFormulaLeafInlineScalar({
            state: args.state,
            formula: runtimeFormula,
          })
          if (inlineValue !== undefined) {
            inlineInitialDirectScalarWriter ??= createInitialFormulaValueWriter(args)
            if (inlineValue.tag === ValueTag.Number) {
              inlineInitialDirectScalarWriter.writeNumberAt(prepared.cellIndex, prepared.sheetId, prepared.col, inlineValue.value)
            } else {
              inlineInitialDirectScalarWriter.writeValueAt(prepared.cellIndex, prepared.sheetId, prepared.col, inlineValue)
            }
            noteInlineInitialDirectScalarCell(prepared, runtimeFormula)
          }
        }
      }
      const noteBoundFormula = (
        prepared: { cellIndex: number; sheetId: number; row: number; col: number },
        runtimeFormula: RuntimeFormula | undefined,
      ): void => {
        if (hadExistingFormulas) {
          formulaChangedCount = args.markFormulaChanged(prepared.cellIndex, formulaChangedCount)
        }
        topologyChanged = true
        pushOrderedPreparedCellIndex(prepared.cellIndex)
        if (canAssignTopoInBatch && pendingFormulaCells) {
          if (!canEvaluateInitialDirectRuntimeFormula(runtimeFormula)) {
            allPreparedFormulasCanUseInitialDirectEvaluation = false
          }
          if (runtimeFormula?.directAggregate !== undefined) {
            hasInitialPrefixAggregateCandidates = true
          }
          const hasPendingDependency =
            runtimeFormula !== undefined &&
            hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells, (rangeIndex) => args.state.ranges.getMembersView(rangeIndex))
          if (!runtimeFormula || hasPendingDependency) {
            needsFreshTopoRebuild ||= hasPendingDependency
            canAssignTopoInBatch = false
          } else {
            args.state.workbook.cellStore.topoRanks[prepared.cellIndex] = nextTopoRank
            nextTopoRank += 1
            tryInlineInitialDirectScalarEvaluation(prepared, runtimeFormula)
          }
        }
      }

      args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
      try {
        args.clearTemplateFormulaCache()
        const compileStarted = performance.now()
        const bindFormulaEntries = (): void => {
          args.state.workbook.withBatchedColumnVersionUpdates(() => {
            for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
              args.checkEvaluationBudget()
              const ref = initialFormulaEntryRefAt(refs, refIndex)
              const cellIndex = hadExistingFormulas ? pendingInitialFormulaCellIndices[refIndex]! : targetCellIndices[refIndex]!
              try {
                const prepared = resolveEntry(ref, cellIndex)
                const requiresWorkbookMetadataBinding = compiledFormulaRequiresWorkbookMetadataBinding(prepared.compiled)
                if (requiresWorkbookMetadataBinding) {
                  args.bindFormula(prepared.cellIndex, prepared.ownerSheetName, prepared.source)
                } else {
                  args.bindPreparedFormula(
                    prepared.cellIndex,
                    prepared.ownerSheetName,
                    prepared.source,
                    prepared.compiled,
                    prepared.templateId,
                    {
                      deferFamilyRegistration:
                        shouldDeferFormulaFamilyIndex ||
                        deferredFormulaFamilyRuns !== undefined ||
                        alignedFreshFormulaFamilyRuns !== undefined,
                      deferFormulaInstanceRegistration: shouldDeferFormulaInstanceTable,
                      assumeFreshFormula: !hadExistingFormulas,
                      resolveWorkbookDateSystem,
                      ownerPosition: {
                        sheetName: prepared.ownerSheetName,
                        row: prepared.row,
                        col: prepared.col,
                      },
                    },
                  )
                }
                const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
                if (alignedFreshFormulaFamilyRuns === undefined) {
                  noteDeferredFormulaFamilyRunMember(deferredFormulaFamilyRuns, prepared, runtimeFormula)
                }
                noteDeferredFormulaInstance(deferredFormulaInstances, prepared, runtimeFormula)
                noteBoundFormula(prepared, runtimeFormula)
              } catch (error) {
                rethrowFatalFormulaBindingError(error)
                noteSkippedOrderedPreparedCellIndex()
                topologyChanged = args.removeFormula(cellIndex) || topologyChanged
                args.setInvalidFormulaValue(cellIndex)
                changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
              }
              if (pendingFormulaCells) {
                pendingFormulaCells[cellIndex] = 0
              }
            }
            if (args.state.ranges.size > 0) {
              const reboundCount = formulaChangedCount
              formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
              topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
            }
            if (
              !flushAlignedFreshFormulaFamilyRuns({
                runs: alignedFreshFormulaFamilyRuns,
                shouldDeferFormulaFamilyIndex,
                deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
                deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
                registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
                checkEvaluationBudget: args.checkEvaluationBudget,
              })
            ) {
              flushDeferredInitialFormulaFamilyRuns({
                runs: deferredFormulaFamilyRuns,
                shouldDeferFormulaFamilyIndex,
                deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
                deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
                registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
              })
            }
            if (shouldDeferFormulaInstanceTable) {
              if (deferredFormulaInstances) {
                args.hydrateFreshFormulaInstances?.(deferredFormulaInstances)
              } else {
                args.deferFormulaInstanceTableRebuild?.()
              }
            }
            inlineInitialDirectScalarWriter?.flush()
          })
        }
        args.withInitialFormulaCells(pendingInitialFormulaCellIndices, bindFormulaEntries)
        const canUseNativeInitialDirectScalarRowChain =
          !hasInitialPrefixAggregateCandidates &&
          nativeInitialDirectScalarRowChainBatch !== undefined &&
          nativeInitialDirectScalarRowChainBatch.count === orderedPreparedCellCount &&
          orderedPreparedCellCount >= MIN_INITIAL_NATIVE_DIRECT_SCALAR_ROW_CHAIN_BATCH_SIZE
        canUseNativeInitialDirectScalarOverLimit =
          !hasInitialPrefixAggregateCandidates &&
          nativeInitialDirectScalarBatch !== undefined &&
          nativeInitialDirectScalarBatch.count === orderedPreparedCellCount &&
          orderedPreparedCellCount > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT
        canUseNativeInitialDirectLookupOverLimit =
          !hasInitialPrefixAggregateCandidates &&
          nativeInitialDirectLookupBatch !== undefined &&
          nativeInitialDirectLookupBatch.count === orderedPreparedCellCount &&
          orderedPreparedCellCount > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT
        canUseNativeInitialPrefixAggregateOverLimit =
          hasInitialPrefixAggregateCandidates &&
          orderedPreparedCellCount > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
          canEvaluateInitialPrefixAggregateGroupsNatively(args, orderedPreparedCellList(), { requireAllCells: true })
        const canUseNativeInitialDirectScalarRunChunksOverLimit =
          nativeInitialDirectScalarRunChunks !== undefined &&
          orderedPreparedCellCount > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
          nativeInitialDirectScalarRunChunks.hasNativeChunks()
        canUseInitialDirectEvaluation =
          canAssignTopoInBatch &&
          !hadExistingFormulas &&
          changedInputCount === 0 &&
          allPreparedFormulasCanUseInitialDirectEvaluation &&
          (orderedPreparedCellCount <= INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT ||
            canUseNativeInitialDirectScalarRowChain ||
            canUseNativeInitialDirectScalarOverLimit ||
            canUseNativeInitialDirectLookupOverLimit ||
            canUseNativeInitialPrefixAggregateOverLimit ||
            canUseNativeInitialDirectScalarRunChunksOverLimit) &&
          orderedPreparedCellCount === refs.length
        compileMs += performance.now() - compileStarted
      } finally {
        args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
      }

      if (needsFreshTopoRebuild) {
        args.checkEvaluationBudget()
        materializeFreshFormulaChangedBuffer()
        args.prepareRegionQueryIndices()
        args.rebuildTopoRanks()
        args.detectCycles()
        canAssignTopoInBatch = false
        canUseInitialDirectEvaluation = false
        args.state.formulas.forEach((_formula, cellIndex) => {
          if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          }
        })
      } else if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
        args.checkEvaluationBudget()
        materializeFreshFormulaChangedBuffer()
        const repaired =
          !hadCycleMembersBeforeNow() &&
          formulaChangedCount > 0 &&
          args.repairTopoRanks(args.getChangedFormulaBuffer().subarray(0, formulaChangedCount))
        if (!repaired) {
          args.rebuildTopoRanks()
          args.detectCycles()
          args.state.formulas.forEach((_formula, cellIndex) => {
            if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
              changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            }
          })
        }
      }
      const hasVolatileFormulaWork = args.hasVolatileFormulas?.() !== false
      if (hasVolatileFormulaWork) {
        materializeFreshFormulaChangedBuffer()
        formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
      }
      const useInitialDirectEvaluation = canUseInitialDirectEvaluation && logicalFormulaChangedCount() === orderedPreparedCellCount
      if (!useInitialDirectEvaluation) {
        materializeFreshFormulaChangedBuffer()
        args.prepareRegionQueryIndices()
      }
      let recalculated: U32
      const allowInitialDirectEvaluationOverLimit =
        canUseNativeInitialDirectScalarOverLimit ||
        canUseNativeInitialDirectLookupOverLimit ||
        canUseNativeInitialPrefixAggregateOverLimit ||
        (nativeInitialDirectScalarRunChunks !== undefined &&
          orderedPreparedCellCount > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
          nativeInitialDirectScalarRunChunks.hasNativeChunks())
      const canEvaluateWholeBatchNatively =
        useInitialDirectEvaluation &&
        !hasInitialPrefixAggregateCandidates &&
        ((nativeInitialDirectScalarRowChainBatch !== undefined &&
          nativeInitialDirectScalarRowChainBatch.count === orderedPreparedCellCount) ||
          nativeInitialDirectScalarCellCount === orderedPreparedCellCount ||
          nativeInitialDirectLookupCellCount === orderedPreparedCellCount)
      if (useInitialDirectEvaluation && !canEvaluateWholeBatchNatively) {
        const runChunkCells = nativeInitialDirectScalarRunChunks?.evaluate()
        if (runChunkCells) {
          inlineInitialDirectScalarCells.noteReusableCells(runChunkCells)
        }
      }
      if (
        useInitialDirectEvaluation &&
        nativeInitialDirectScalarRowChainBatch !== undefined &&
        nativeInitialDirectScalarRowChainBatch.count === orderedPreparedCellCount &&
        !hasInitialPrefixAggregateCandidates
      ) {
        const native = nativeInitialDirectScalarRowChainBatch.evaluate()
        if (native) {
          recalculated = native
          args.deferKernelSync(recalculated)
          addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedPreparedCellCount)
        } else {
          const direct = evaluateInitialDirectFormulas(args, orderedPreparedCellList(), {
            alreadyValidated: true,
            hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
            allowOverLimit: allowInitialDirectEvaluationOverLimit,
          })
          if (direct) {
            recalculated = direct
          } else {
            const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
            const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
            recalculated = args.recalculate(changedRoots, changedInputArray)
          }
        }
      } else if (
        useInitialDirectEvaluation &&
        nativeInitialDirectScalarCellCount === orderedPreparedCellCount &&
        !hasInitialPrefixAggregateCandidates
      ) {
        const native = nativeInitialDirectScalarBatch?.evaluate()
        if (native) {
          recalculated = native
          args.deferKernelSync(recalculated)
          addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedPreparedCellCount)
        } else {
          const direct = evaluateInitialDirectFormulas(args, orderedPreparedCellList(), {
            alreadyValidated: true,
            hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
            allowOverLimit: allowInitialDirectEvaluationOverLimit,
          })
          if (direct) {
            recalculated = direct
          } else {
            const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
            const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
            recalculated = args.recalculate(changedRoots, changedInputArray)
          }
        }
      } else if (
        useInitialDirectEvaluation &&
        nativeInitialDirectLookupCellCount === orderedPreparedCellCount &&
        !hasInitialPrefixAggregateCandidates
      ) {
        const native = nativeInitialDirectLookupBatch?.evaluate()
        if (native) {
          recalculated = native
          args.deferKernelSync(recalculated)
          addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedPreparedCellCount)
        } else {
          const direct = evaluateInitialDirectFormulas(args, orderedPreparedCellList(), {
            alreadyValidated: true,
            hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
            allowOverLimit: allowInitialDirectEvaluationOverLimit,
          })
          if (direct) {
            recalculated = direct
          } else {
            const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
            const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
            recalculated = args.recalculate(changedRoots, changedInputArray)
          }
        }
      } else if (
        useInitialDirectEvaluation &&
        inlineInitialDirectScalarCells.cellCount === orderedPreparedCellCount &&
        !hasInitialPrefixAggregateCandidates
      ) {
        recalculated = inlineInitialDirectScalarCells.cellIndices!.subarray(0, inlineInitialDirectScalarCells.cellCount)
        args.deferKernelSync(recalculated)
        addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedPreparedCellCount)
      } else if (useInitialDirectEvaluation) {
        const direct = evaluateInitialDirectFormulas(args, orderedPreparedCellList(), {
          alreadyValidated: true,
          hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
          allowOverLimit: allowInitialDirectEvaluationOverLimit,
          ...(inlineInitialDirectScalarCells.cellCount > 0
            ? {
                preEvaluatedCellIndices: inlineInitialDirectScalarCells.cellIndices!,
                preEvaluatedCellCount: inlineInitialDirectScalarCells.cellCount,
                preEvaluatedCellsAreReusable: inlineInitialDirectScalarCells.allCellsReusable,
              }
            : {}),
        })
        if (direct) {
          recalculated = direct
        } else {
          const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
          const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
          recalculated = args.recalculate(changedRoots, changedInputArray)
        }
      } else {
        const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
        const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
        recalculated =
          canAssignTopoInBatch && !hadExistingFormulas && orderedPreparedCellCount > 0
            ? args.recalculatePreordered(changedRoots, orderedPreparedCellList(), orderedPreparedCellCount, changedInputArray)
            : args.recalculate(changedRoots, changedInputArray)
      }
      recalculated = args.reconcilePivotOutputs(recalculated, false)
      if (!hadExistingFormulas && hasVolatileFormulaWork) {
        recalculated = recalculateFreshVolatileFormulasAfterInitialMaterialization(args, recalculated)
      }
      void recalculated
      const lastMetrics = args.state.getLastMetrics()
      args.state.setLastMetrics({
        ...lastMetrics,
        batchId: lastMetrics.batchId + 1,
        changedInputCount: changedInputCount + logicalFormulaChangedCount(),
        compileMs,
      })
    } finally {
      args.endEvaluationBudget()
    }
  }

  const initializeCellFormulasAtNow = (refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void => {
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver(args.compileTemplateFormula)
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        return ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.mutation.row, ref.mutation.col)
      },
      (ref, cellIndex) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        const ownerSheetName = resolveSheetName(ref.sheetId)
        const template = resolveInitialTemplateFormula(ref.mutation.formula, ref.mutation.row, ref.mutation.col)
        return {
          cellIndex,
          sheetId: ref.sheetId,
          row: ref.mutation.row,
          col: ref.mutation.col,
          ownerSheetName,
          source: ref.mutation.formula,
          compiled: template.compiled,
          templateId: template.templateId,
        }
      },
    )
  }

  const initializeFormulaSourcesAtNow = (refs: EngineFormulaSourceRefs, potentialNewCells?: number): void => {
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver(args.compileTemplateFormula)
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      (ref, cellIndex) => {
        const ownerSheetName = resolveSheetName(ref.sheetId)
        const template = resolveInitialTemplateFormula(ref.source, ref.row, ref.col)
        return {
          cellIndex,
          sheetId: ref.sheetId,
          row: ref.row,
          col: ref.col,
          ownerSheetName,
          source: ref.source,
          compiled: template.compiled,
          templateId: template.templateId,
        }
      },
    )
  }

  const initializePreparedCellFormulasAtNow = (
    refs: InitialFormulaEntryRefSource<PreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ): void => {
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      (ref, cellIndex) => ({
        cellIndex,
        sheetId: ref.sheetId,
        row: ref.row,
        col: ref.col,
        ownerSheetName: resolveSheetName(ref.sheetId),
        source: ref.source,
        compiled: ref.compiled,
        ...(ref.templateId !== undefined ? { templateId: ref.templateId } : {}),
      }),
    )
  }

  const initializeHydratedPreparedCellFormulasAtNowUnchecked = (
    refs: InitialFormulaEntryRefSource<HydratedPreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginMutationCollection()
    args.checkEvaluationBudget()
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    let topologyChanged = false
    let compileMs = 0
    let workbookDateSystem: string | undefined
    const resolveWorkbookDateSystem = (): string | undefined =>
      (workbookDateSystem ??= args.state.workbook.getCalculationSettings().dateSystem)
    const reservedNewCells = potentialNewCells ?? refs.length
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const { targetCellIndices, pendingInitialFormulaCellIndices, pendingFormulaCells } = createInitialFormulaCellIndexPlan({
      refs,
      hadExistingFormulas,
      resolveCellIndex: (ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      checkEvaluationBudget: args.checkEvaluationBudget,
    })
    let canAssignTopoInBatch = !hadExistingFormulas
    let needsFreshTopoRebuild = false
    let nextTopoRank = 0
    const shouldDeferFormulaFamilyIndex = !hadExistingFormulas && args.deferFormulaFamilyIndexRebuild !== undefined
    const shouldDeferFormulaInstanceTable =
      !hadExistingFormulas && (args.hydrateFreshFormulaInstances !== undefined || args.deferFormulaInstanceTableRebuild !== undefined)
    const alignedFreshFormulaInstances =
      !hadExistingFormulas && args.hydrateFreshFormulaInstances !== undefined ? readAlignedFreshFormulaInstancesFromRefs(refs) : undefined
    const alignedFreshFormulaFamilyRuns = readAlignedFreshFormulaFamilyRunsFromRefs({
      refs,
      hadExistingFormulas,
      counters: args.state.counters,
    })
    const deferredFormulaInstances =
      !hadExistingFormulas && args.hydrateFreshFormulaInstances !== undefined && alignedFreshFormulaInstances === undefined ? [] : undefined
    const canCaptureDeferredFormulaFamilyRuns = !shouldDeferFormulaFamilyIndex || args.deferFormulaFamilyIndexRuns !== undefined
    const deferredFormulaFamilyRuns =
      hadExistingFormulas || alignedFreshFormulaFamilyRuns !== undefined || !canCaptureDeferredFormulaFamilyRuns
        ? undefined
        : createDeferredInitialFormulaFamilyRunMap()

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      const valueWriter = createInitialFormulaValueWriter(args)
      const bindFormulaEntries = (): void => {
        args.state.workbook.withBatchedColumnVersionUpdates(() => {
          for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
            args.checkEvaluationBudget()
            const ref = initialFormulaEntryRefAt(refs, refIndex)
            const cellIndex = hadExistingFormulas ? pendingInitialFormulaCellIndices[refIndex]! : targetCellIndices[refIndex]!
            const ownerSheetName = resolveSheetName(ref.sheetId)
            const usedHydratedDirectScalarFastBinding = tryBindHydratedFreshDirectScalarFormula(
              args,
              hadExistingFormulas,
              cellIndex,
              ownerSheetName,
              ref,
            )
            if (usedHydratedDirectScalarFastBinding) {
              topologyChanged = true
            } else {
              topologyChanged =
                args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId, {
                  deferFamilyRegistration:
                    shouldDeferFormulaFamilyIndex || deferredFormulaFamilyRuns !== undefined || alignedFreshFormulaFamilyRuns !== undefined,
                  preserveCachedValueOnFullRecalc: ref.preserveCachedValueOnFullRecalc === true,
                  deferFormulaInstanceRegistration: shouldDeferFormulaInstanceTable,
                  assumeFreshFormula: !hadExistingFormulas,
                  ownerPosition: {
                    sheetName: ownerSheetName,
                    row: ref.row,
                    col: ref.col,
                  },
                  resolveWorkbookDateSystem,
                }) || topologyChanged
            }
            const runtimeFormula = args.state.formulas.get(cellIndex)
            noteDeferredFormulaInstance(deferredFormulaInstances, { cellIndex, row: ref.row, col: ref.col, ownerSheetName }, runtimeFormula)
            if (alignedFreshFormulaFamilyRuns === undefined) {
              noteDeferredFormulaFamilyRunMember(
                deferredFormulaFamilyRuns,
                {
                  cellIndex,
                  sheetId: ref.sheetId,
                  row: ref.row,
                  col: ref.col,
                  ...(ref.templateId !== undefined ? { templateId: ref.templateId } : {}),
                },
                runtimeFormula,
              )
            }
            valueWriter.writeValueAt(cellIndex, ref.sheetId, ref.col, ref.value)
            if (canAssignTopoInBatch && pendingFormulaCells) {
              const hasPendingDependency =
                runtimeFormula !== undefined &&
                hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells, (rangeIndex) =>
                  args.state.ranges.getMembersView(rangeIndex),
                )
              if (!runtimeFormula || hasPendingDependency) {
                needsFreshTopoRebuild ||= hasPendingDependency
                canAssignTopoInBatch = false
              } else {
                args.state.workbook.cellStore.topoRanks[cellIndex] = nextTopoRank
                nextTopoRank += 1
              }
            }
            if (pendingFormulaCells) {
              pendingFormulaCells[cellIndex] = 0
            }
          }
          if (
            !flushAlignedFreshFormulaFamilyRuns({
              runs: alignedFreshFormulaFamilyRuns,
              shouldDeferFormulaFamilyIndex,
              deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
              deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
              registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
              checkEvaluationBudget: args.checkEvaluationBudget,
            })
          ) {
            flushDeferredInitialFormulaFamilyRuns({
              runs: deferredFormulaFamilyRuns,
              shouldDeferFormulaFamilyIndex,
              deferFormulaFamilyIndexRuns: args.deferFormulaFamilyIndexRuns,
              deferFormulaFamilyIndexRebuild: args.deferFormulaFamilyIndexRebuild,
              registerFormulaFamilyRun: registerDeferredFormulaFamilyRun,
              checkEvaluationBudget: args.checkEvaluationBudget,
            })
          }
          if (shouldDeferFormulaInstanceTable) {
            if (alignedFreshFormulaInstances !== undefined) {
              args.hydrateFreshFormulaInstances?.(alignedFreshFormulaInstances)
            } else if (deferredFormulaInstances) {
              args.hydrateFreshFormulaInstances?.(deferredFormulaInstances)
            } else {
              args.deferFormulaInstanceTableRebuild?.()
            }
          }
          args.checkEvaluationBudget()
          valueWriter.flush()
        })
      }
      args.checkEvaluationBudget()
      args.withInitialFormulaCells(pendingInitialFormulaCellIndices, bindFormulaEntries)
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if ((topologyChanged || needsFreshTopoRebuild) && !(canAssignTopoInBatch && !hadExistingFormulas)) {
      args.checkEvaluationBudget()
      const repaired =
        !needsFreshTopoRebuild &&
        !hadCycleMembersBeforeNow() &&
        refs.length > 0 &&
        args.repairTopoRanks(targetCellIndices.length > 0 ? targetCellIndices : pendingInitialFormulaCellIndices)
      if (!repaired) {
        args.checkEvaluationBudget()
        args.rebuildTopoRanks()
        args.checkEvaluationBudget()
        args.detectCycles()
      }
    }
    args.checkEvaluationBudget()
    const lastMetrics = args.state.getLastMetrics()
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: 0,
      compileMs,
      recalcMs: 0,
    })
  }

  const initializeHydratedPreparedCellFormulasAtNow = (
    refs: InitialFormulaEntryRefSource<HydratedPreparedFormulaInitializationRef>,
    potentialNewCells?: number,
  ): void => {
    if (refs.length === 0) {
      return
    }
    args.beginEvaluationBudget(performance.now())
    try {
      args.checkEvaluationBudget()
      initializeHydratedPreparedCellFormulasAtNowUnchecked(refs, potentialNewCells)
    } finally {
      args.endEvaluationBudget()
    }
  }

  return {
    initializeCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializeCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize cell formulas', cause),
            cause,
          }),
      })
    },
    initializePreparedCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializePreparedCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize prepared cell formulas', cause),
            cause,
          }),
      })
    },
    initializeHydratedPreparedCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializeHydratedPreparedCellFormulasAtNow(refs, potentialNewCells)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to initialize hydrated prepared cell formulas', cause),
            cause,
          }),
      })
    },
    initializeCellFormulasAtNow,
    initializeFormulaSourcesAtNow,
    initializePreparedCellFormulasAtNow,
    initializeHydratedPreparedCellFormulasAtNow,
    initializeCachedFormulaSourcesAtNow: (refs, potentialNewCells) =>
      initializeCachedFormulaSourcesAtNowUnchecked({
        serviceArgs: args,
        refs,
        ...(potentialNewCells === undefined ? {} : { potentialNewCells }),
        resolveSheetName,
      }),
  }
}
