import { Effect } from 'effect'
import type { EngineCellMutationRef, EngineFormulaSourceRefs } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeFormula, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { evaluateInitialDirectScalar, evaluateInitialDirectScalarNumber } from './formula-initialization-direct-scalar.js'
import { materializeDeferredFormulaFamilyRunMembers, type DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import { initialFormulaFamilyShapeKey } from './formula-initialization-template-keys.js'
import { createInitialTemplateFormulaResolver } from './formula-initialization-template-resolver.js'
import { createInitialFormulaValueWriter, type InitialFormulaValueWriter } from './formula-initialization-value-writer.js'
import {
  initialFormulaEntryRefAt,
  type InitialFormulaCellIndexList,
  type InitialFormulaEntryRefSource,
  type InitialResolvedFormulaEntry,
} from './formula-initialization-refs.js'
import { recalculateFreshVolatileFormulasAfterInitialMaterialization } from './formula-initialization-volatile-pass.js'
import {
  canEvaluateInitialDirectRuntimeFormula,
  compiledFormulaRequiresWorkbookMetadataBinding,
  hasPendingFormulaDependency,
  mutationErrorMessage,
} from './formula-initialization-predicates.js'
import { evaluateInitialPrefixAggregateGroups } from './formula-initialization-prefix-aggregates.js'
import type {
  EngineFormulaInitializationService,
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'

export type {
  EngineFormulaInitializationService,
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'

const INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT = 16_384
const DEFERRED_FORMULA_FAMILY_RUN_CAPTURE_LIMIT = 16_384
const EMPTY_U32 = new Uint32Array(0)

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
    runs: Map<string, DeferredInitialFormulaFamilyRun> | undefined,
    prepared: {
      readonly cellIndex: number
      readonly sheetId: number
      readonly row: number
      readonly col: number
      readonly templateId?: number
    },
  ): void => {
    if (!runs) {
      return
    }
    const templateId = prepared.templateId
    if (templateId === undefined) {
      return
    }
    const familyKey = `${prepared.sheetId}\t${templateId}\t${prepared.col}`
    let run = runs.get(familyKey)
    if (!run) {
      const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
      if (runtimeFormula === undefined) {
        return
      }
      run = {
        sheetId: prepared.sheetId,
        templateId,
        shapeKey: initialFormulaFamilyShapeKey(runtimeFormula),
        axis: 'row',
        fixedIndex: prepared.col,
        start: prepared.row,
        step: 0,
        lastIndex: prepared.row,
        ordered: true,
        cellIndices: [],
      }
      runs.set(familyKey, run)
    } else {
      const nextStep = prepared.row - run.lastIndex
      let breaksOrder = false
      if (run.cellIndices.length === 1) {
        run.step = nextStep
      } else if (run.step !== nextStep) {
        breaksOrder = true
      }
      if (prepared.row <= run.lastIndex || prepared.col !== run.fixedIndex) {
        breaksOrder = true
      }
      if (breaksOrder) {
        if (!run.rows) {
          const priorStep = run.cellIndices.length <= 1 ? 1 : run.step
          const start = run.start
          run.rows = Array.from({ length: run.cellIndices.length }, (_value, index) => start + priorStep * index)
        }
        run.ordered = false
      }
      run.lastIndex = prepared.row
    }
    run.cellIndices.push(prepared.cellIndex)
    run.rows?.push(prepared.row)
  }

  const registerDeferredFormulaFamilyRun = (run: DeferredInitialFormulaFamilyRun): void => {
    const step = run.cellIndices.length <= 1 ? 1 : run.step
    if (
      run.ordered &&
      step > 0 &&
      args.registerFreshFormulaFamilyRun({
        sheetId: run.sheetId,
        templateId: run.templateId,
        shapeKey: run.shapeKey,
        axis: run.axis,
        fixedIndex: run.fixedIndex,
        start: run.start,
        step,
        cellIndices: run.cellIndices,
      })
    ) {
      return
    }
    args.upsertFormulaFamilyRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      members: materializeDeferredFormulaFamilyRunMembers(run),
    })
  }

  const canEvaluateInitialDirectFormula = (cellIndex: number): boolean => {
    return canEvaluateInitialDirectRuntimeFormula(args.state.formulas.get(cellIndex))
  }

  const evaluateInitialDirectFormulas = (
    orderedCellIndices: InitialFormulaCellIndexList,
    options?: { readonly alreadyValidated?: boolean; readonly hasPrefixAggregateCandidates?: boolean },
  ): U32 | undefined => {
    if (
      orderedCellIndices.length === 0 ||
      orderedCellIndices.length > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT ||
      (options?.alreadyValidated !== true && !orderedCellIndices.every(canEvaluateInitialDirectFormula))
    ) {
      return undefined
    }
    let changedCellBuffer = new Uint32Array(Math.max(orderedCellIndices.length, 1))
    let changedCellCount = 0
    const pushChangedCellIndex = (cellIndex: number): void => {
      if (changedCellCount === changedCellBuffer.length) {
        const next = new Uint32Array(changedCellBuffer.length * 2)
        next.set(changedCellBuffer)
        changedCellBuffer = next
      }
      changedCellBuffer[changedCellCount] = cellIndex
      changedCellCount += 1
    }
    const valueWriter = createInitialFormulaValueWriter(args)
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      const prefixAggregateHandled =
        options?.hasPrefixAggregateCandidates === true
          ? evaluateInitialPrefixAggregateGroups(args, orderedCellIndices, pushChangedCellIndex, valueWriter.writeValue)
          : undefined
      for (let index = 0; index < orderedCellIndices.length; index += 1) {
        args.checkEvaluationBudget()
        const cellIndex = orderedCellIndices[index]!
        if (prefixAggregateHandled?.has(cellIndex)) {
          continue
        }
        const formula = args.state.formulas.get(cellIndex)
        if (formula?.directScalar !== undefined) {
          const numericValue = evaluateInitialDirectScalarNumber(args.state, formula.directScalar)
          if (numericValue !== undefined) {
            valueWriter.writeNumber(cellIndex, numericValue)
            pushChangedCellIndex(cellIndex)
            continue
          }
          const fallbackValue = evaluateInitialDirectScalar(args.state, formula.directScalar)
          if (fallbackValue !== undefined) {
            valueWriter.writeValue(cellIndex, fallbackValue)
            pushChangedCellIndex(cellIndex)
            continue
          }
        }
        const changedSpillIndices = args.evaluateDirectFormula(cellIndex)
        pushChangedCellIndex(cellIndex)
        if (changedSpillIndices) {
          for (let spillIndex = 0; spillIndex < changedSpillIndices.length; spillIndex += 1) {
            pushChangedCellIndex(changedSpillIndices[spillIndex]!)
          }
        }
      }
      valueWriter.flush()
    })
    const changedCellIndices = changedCellBuffer.subarray(0, changedCellCount)
    args.deferKernelSync(changedCellIndices)
    addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedCellIndices.length)
    return changedCellIndices
  }

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
      const reservedNewCells = potentialNewCells ?? refs.length
      const hadExistingFormulas = args.state.formulas.size > 0
      args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
      args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
      args.resetMaterializedCellScratch(reservedNewCells)
      const targetCellIndices = hadExistingFormulas ? EMPTY_U32 : new Uint32Array(refs.length)
      const pendingInitialFormulaCellIndices = hadExistingFormulas ? new Uint32Array(refs.length) : targetCellIndices
      let maxTargetCellIndex = 0
      for (let index = 0; index < refs.length; index += 1) {
        args.checkEvaluationBudget()
        const cellIndex = resolveCellIndex(initialFormulaEntryRefAt(refs, index))
        if (hadExistingFormulas) {
          pendingInitialFormulaCellIndices[index] = cellIndex
        } else {
          targetCellIndices[index] = cellIndex
          if (cellIndex > maxTargetCellIndex) {
            maxTargetCellIndex = cellIndex
          }
        }
      }
      const pendingFormulaCells = hadExistingFormulas ? undefined : new Uint8Array(maxTargetCellIndex + 1)
      if (pendingFormulaCells) {
        for (let index = 0; index < targetCellIndices.length; index += 1) {
          args.checkEvaluationBudget()
          pendingFormulaCells[targetCellIndices[index]!] = 1
        }
      }
      let canAssignTopoInBatch = !hadExistingFormulas
      let nextTopoRank = 0
      let orderedPreparedCellIndices: number[] | undefined = hadExistingFormulas ? [] : undefined
      let orderedPreparedCellCount = 0
      let canUseInitialDirectEvaluation = false
      let allPreparedFormulasCanUseInitialDirectEvaluation = true
      let hasInitialPrefixAggregateCandidates = false
      let inlineInitialDirectScalarWriter: InitialFormulaValueWriter | undefined
      let inlineInitialDirectScalarCellBuffer: Uint32Array | undefined = hadExistingFormulas
        ? new Uint32Array(Math.max(refs.length, 1))
        : undefined
      let inlineInitialDirectScalarCellCount = 0
      const shouldDeferFormulaFamilyIndex = !hadExistingFormulas && args.deferFormulaFamilyIndexRebuild !== undefined
      const shouldDeferFormulaInstanceTable = !hadExistingFormulas && args.deferFormulaInstanceTableRebuild !== undefined
      const canCaptureDeferredFormulaFamilyRuns =
        !shouldDeferFormulaFamilyIndex ||
        (args.deferFormulaFamilyIndexRuns !== undefined && refs.length <= DEFERRED_FORMULA_FAMILY_RUN_CAPTURE_LIMIT)
      const deferredFormulaFamilyRuns =
        hadExistingFormulas || !canCaptureDeferredFormulaFamilyRuns ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()
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

      const materializeInlineInitialDirectScalarCellBuffer = (): Uint32Array => {
        if (inlineInitialDirectScalarCellBuffer) {
          return inlineInitialDirectScalarCellBuffer
        }
        inlineInitialDirectScalarCellBuffer = new Uint32Array(Math.max(refs.length, 1))
        inlineInitialDirectScalarCellBuffer.set(targetCellIndices.subarray(0, inlineInitialDirectScalarCellCount))
        return inlineInitialDirectScalarCellBuffer
      }
      const pushInlineInitialDirectScalarCell = (cellIndex: number): void => {
        if (
          !hadExistingFormulas &&
          inlineInitialDirectScalarCellBuffer === undefined &&
          cellIndex === targetCellIndices[inlineInitialDirectScalarCellCount]
        ) {
          inlineInitialDirectScalarCellCount += 1
          return
        }
        let buffer = materializeInlineInitialDirectScalarCellBuffer()
        if (inlineInitialDirectScalarCellCount === buffer.length) {
          const next = new Uint32Array(buffer.length * 2)
          next.set(buffer)
          inlineInitialDirectScalarCellBuffer = next
          buffer = next
        }
        buffer[inlineInitialDirectScalarCellCount] = cellIndex
        inlineInitialDirectScalarCellCount += 1
      }

      const tryInlineInitialDirectScalarEvaluation = (
        prepared: { cellIndex: number; sheetId: number; col: number },
        runtimeFormula: RuntimeFormula | undefined,
      ): void => {
        if (
          hadExistingFormulas ||
          !canAssignTopoInBatch ||
          !runtimeFormula ||
          runtimeFormula.compiled.volatile ||
          runtimeFormula.compiled.producesSpill ||
          runtimeFormula.directScalar === undefined
        ) {
          return
        }
        const numericValue = evaluateInitialDirectScalarNumber(args.state, runtimeFormula.directScalar)
        inlineInitialDirectScalarWriter ??= createInitialFormulaValueWriter(args)
        if (numericValue !== undefined) {
          inlineInitialDirectScalarWriter.writeNumberAt(prepared.cellIndex, prepared.sheetId, prepared.col, numericValue)
          pushInlineInitialDirectScalarCell(prepared.cellIndex)
          return
        }
        const fallbackValue = evaluateInitialDirectScalar(args.state, runtimeFormula.directScalar)
        if (fallbackValue !== undefined) {
          inlineInitialDirectScalarWriter.writeValueAt(prepared.cellIndex, prepared.sheetId, prepared.col, fallbackValue)
          pushInlineInitialDirectScalarCell(prepared.cellIndex)
        }
      }
      const noteBoundFormula = (prepared: { cellIndex: number; sheetId: number; col: number }): void => {
        if (hadExistingFormulas) {
          formulaChangedCount = args.markFormulaChanged(prepared.cellIndex, formulaChangedCount)
        }
        topologyChanged = true
        pushOrderedPreparedCellIndex(prepared.cellIndex)
        if (canAssignTopoInBatch && pendingFormulaCells) {
          const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
          if (!canEvaluateInitialDirectRuntimeFormula(runtimeFormula)) {
            allPreparedFormulasCanUseInitialDirectEvaluation = false
          }
          if (runtimeFormula?.directAggregate !== undefined) {
            hasInitialPrefixAggregateCandidates = true
          }
          if (
            !runtimeFormula ||
            hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells, (rangeIndex) => args.state.ranges.getMembersView(rangeIndex))
          ) {
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
                      deferFamilyRegistration: shouldDeferFormulaFamilyIndex || deferredFormulaFamilyRuns !== undefined,
                      deferFormulaInstanceRegistration: shouldDeferFormulaInstanceTable,
                      assumeFreshFormula: !hadExistingFormulas,
                    },
                  )
                  noteDeferredFormulaFamilyRunMember(deferredFormulaFamilyRuns, prepared)
                }
                noteBoundFormula(prepared)
              } catch {
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
            if (shouldDeferFormulaFamilyIndex) {
              if (deferredFormulaFamilyRuns) {
                args.deferFormulaFamilyIndexRuns?.([...deferredFormulaFamilyRuns.values()])
              } else {
                args.deferFormulaFamilyIndexRebuild?.()
              }
            } else {
              deferredFormulaFamilyRuns?.forEach(registerDeferredFormulaFamilyRun)
            }
            if (shouldDeferFormulaInstanceTable) {
              args.deferFormulaInstanceTableRebuild?.()
            }
            inlineInitialDirectScalarWriter?.flush()
          })
        }
        args.withInitialFormulaCells(pendingInitialFormulaCellIndices, bindFormulaEntries)
        canUseInitialDirectEvaluation =
          canAssignTopoInBatch &&
          !hadExistingFormulas &&
          changedInputCount === 0 &&
          allPreparedFormulasCanUseInitialDirectEvaluation &&
          orderedPreparedCellCount <= INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
          orderedPreparedCellCount === refs.length
        compileMs += performance.now() - compileStarted
      } finally {
        args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
      }

      if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
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
      if (
        useInitialDirectEvaluation &&
        inlineInitialDirectScalarCellCount === orderedPreparedCellCount &&
        !hasInitialPrefixAggregateCandidates
      ) {
        recalculated = inlineInitialDirectScalarCellBuffer
          ? inlineInitialDirectScalarCellBuffer.subarray(0, inlineInitialDirectScalarCellCount)
          : targetCellIndices
        args.deferKernelSync(recalculated)
        addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedPreparedCellCount)
      } else if (useInitialDirectEvaluation) {
        const direct = evaluateInitialDirectFormulas(orderedPreparedCellList(), {
          alreadyValidated: true,
          hasPrefixAggregateCandidates: hasInitialPrefixAggregateCandidates,
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

  const initializePreparedCellFormulasAtNow = (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number): void => {
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
    refs: readonly HydratedPreparedFormulaInitializationRef[],
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
    const reservedNewCells = potentialNewCells ?? refs.length
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + reservedNewCells + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const targetCellIndices = hadExistingFormulas
      ? []
      : refs.map((ref) => {
          args.checkEvaluationBudget()
          return ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)
        })
    const pendingInitialFormulaCellIndices = hadExistingFormulas
      ? refs.map((ref) => {
          args.checkEvaluationBudget()
          return ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)
        })
      : targetCellIndices
    const pendingFormulaCells = hadExistingFormulas
      ? undefined
      : new Uint8Array(args.state.workbook.cellStore.capacity + reservedNewCells + 1)
    if (pendingFormulaCells) {
      for (let index = 0; index < targetCellIndices.length; index += 1) {
        args.checkEvaluationBudget()
        pendingFormulaCells[targetCellIndices[index]!] = 1
      }
    }
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0
    const shouldDeferFormulaFamilyIndex = !hadExistingFormulas && args.deferFormulaFamilyIndexRebuild !== undefined
    const canCaptureDeferredFormulaFamilyRuns =
      !shouldDeferFormulaFamilyIndex ||
      (args.deferFormulaFamilyIndexRuns !== undefined && refs.length <= DEFERRED_FORMULA_FAMILY_RUN_CAPTURE_LIMIT)
    const deferredFormulaFamilyRuns =
      hadExistingFormulas || !canCaptureDeferredFormulaFamilyRuns ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      const valueWriter = createInitialFormulaValueWriter(args)
      const bindFormulaEntries = (): void => {
        args.state.workbook.withBatchedColumnVersionUpdates(() => {
          for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
            args.checkEvaluationBudget()
            const ref = refs[refIndex]!
            const cellIndex = hadExistingFormulas ? pendingInitialFormulaCellIndices[refIndex]! : targetCellIndices[refIndex]!
            const ownerSheetName = resolveSheetName(ref.sheetId)
            topologyChanged =
              args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId, {
                deferFamilyRegistration: shouldDeferFormulaFamilyIndex || deferredFormulaFamilyRuns !== undefined,
                preserveCachedValueOnFullRecalc: ref.preserveCachedValueOnFullRecalc === true,
              }) || topologyChanged
            noteDeferredFormulaFamilyRunMember(deferredFormulaFamilyRuns, {
              cellIndex,
              sheetId: ref.sheetId,
              row: ref.row,
              col: ref.col,
              ...(ref.templateId !== undefined ? { templateId: ref.templateId } : {}),
            })
            valueWriter.writeValueAt(cellIndex, ref.sheetId, ref.col, ref.value)
            if (canAssignTopoInBatch && pendingFormulaCells) {
              const runtimeFormula = args.state.formulas.get(cellIndex)
              if (
                !runtimeFormula ||
                hasPendingFormulaDependency(runtimeFormula, pendingFormulaCells, (rangeIndex) =>
                  args.state.ranges.getMembersView(rangeIndex),
                )
              ) {
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
          if (shouldDeferFormulaFamilyIndex) {
            if (deferredFormulaFamilyRuns) {
              args.deferFormulaFamilyIndexRuns?.([...deferredFormulaFamilyRuns.values()])
            } else {
              args.deferFormulaFamilyIndexRebuild?.()
            }
          } else {
            deferredFormulaFamilyRuns?.forEach((run) => {
              args.checkEvaluationBudget()
              registerDeferredFormulaFamilyRun(run)
            })
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

    if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
      args.checkEvaluationBudget()
      const repaired =
        !hadCycleMembersBeforeNow() &&
        refs.length > 0 &&
        args.repairTopoRanks(
          Uint32Array.from(
            targetCellIndices.length > 0
              ? targetCellIndices
              : refs.map((ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)),
          ),
        )
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
    refs: readonly HydratedPreparedFormulaInitializationRef[],
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
  }
}
