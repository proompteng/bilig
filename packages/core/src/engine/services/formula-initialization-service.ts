import { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef, EngineFormulaSourceRefs } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import type { FormulaFamilyFreshUniformRunRegistrationArgs, FormulaFamilyRunUpsertArgs } from '../../formula/formula-family-store.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { translateSimpleDirectScalarFormula } from '../../formula/simple-direct-scalar-compile.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeFormula, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { evaluateInitialDirectScalar, evaluateInitialDirectScalarNumber } from './formula-initialization-direct-scalar.js'
import { materializeDeferredFormulaFamilyRunMembers, type DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import {
  initialFormulaFamilyShapeKey,
  tryBuildInitialSimpleRowRelativeBinaryTemplateKey,
  type InitialTemplateFormulaCacheEntry,
} from './formula-initialization-template-keys.js'
import { createInitialFormulaValueWriter, type InitialFormulaValueWriter } from './formula-initialization-value-writer.js'
import {
  initialFormulaEntryRefAt,
  type InitialFormulaCellIndexList,
  type InitialFormulaEntryRefSource,
  type InitialResolvedFormulaEntry,
} from './formula-initialization-refs.js'
import {
  canEvaluateInitialDirectRuntimeFormula,
  compiledFormulaRequiresWorkbookMetadataBinding,
  hasPendingFormulaDependency,
  mutationErrorMessage,
} from './formula-initialization-predicates.js'
import type { InitialPrefixAggregateGroup } from './formula-initialization-prefix-aggregates.js'
import type {
  EngineFormulaInitializationService,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'

export type {
  EngineFormulaInitializationService,
  HydratedPreparedFormulaInitializationRef,
  PreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'

const INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT = 16_384
const EMPTY_U32 = new Uint32Array(0)

export function createEngineFormulaInitializationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    'workbook' | 'strings' | 'formulas' | 'ranges' | 'counters' | 'getLastMetrics' | 'setLastMetrics'
  >
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => void
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
    },
  ) => boolean
  readonly upsertFormulaFamilyRun: (args: FormulaFamilyRunUpsertArgs) => void
  readonly registerFreshFormulaFamilyRun: (args: FormulaFamilyFreshUniformRunRegistrationArgs) => boolean
  readonly deferFormulaFamilyIndexRebuild?: () => void
  readonly deferFormulaInstanceTableRebuild?: () => void
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
}): EngineFormulaInitializationService {
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
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly row: number; readonly col: number },
  ): void => {
    if (!runs) {
      return
    }
    const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
    const templateId = runtimeFormula?.templateId
    if (runtimeFormula === undefined || templateId === undefined) {
      return
    }
    const familyKey = `${prepared.sheetId}\t${templateId}\t${prepared.col}`
    let run = runs.get(familyKey)
    if (!run) {
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

  const createInitialTemplateFormulaResolver = (): ((source: string, row: number, col: number) => FormulaTemplateResolution) => {
    const simpleTemplateCache = new Map<string, InitialTemplateFormulaCacheEntry>()
    return (source, row, col) => {
      const templateKey = tryBuildInitialSimpleRowRelativeBinaryTemplateKey(source, row, col)
      const cached = templateKey === undefined ? undefined : simpleTemplateCache.get(templateKey)
      if (cached) {
        const anchorRowDelta = row - cached.anchorRow
        const anchorColDelta = col - cached.anchorCol
        const compiled = translateSimpleDirectScalarFormula(cached.anchorCompiled, anchorRowDelta, anchorColDelta, source)
        if (compiled) {
          return {
            ...cached.resolution,
            compiled,
            translated: cached.resolution.translated || anchorRowDelta !== 0 || anchorColDelta !== 0,
            rowDelta: cached.resolution.rowDelta + anchorRowDelta,
            colDelta: cached.resolution.colDelta + anchorColDelta,
          }
        }
      }
      const resolution = args.compileTemplateFormula(source, row, col)
      if (templateKey !== undefined) {
        simpleTemplateCache.set(templateKey, {
          resolution,
          anchorRow: row,
          anchorCol: col,
          anchorCompiled: resolution.compiled,
        })
      }
      return resolution
    }
  }

  const evaluateInitialPrefixAggregateGroups = (
    orderedCellIndices: InitialFormulaCellIndexList,
    pushChangedCellIndex: (cellIndex: number) => void,
    writeFormulaValue: (cellIndex: number, value: CellValue) => void,
  ): Set<number> | undefined => {
    const groups = new Map<string, InitialPrefixAggregateGroup>()
    for (let index = 0; index < orderedCellIndices.length; index += 1) {
      args.checkEvaluationBudget()
      const cellIndex = orderedCellIndices[index]!
      const formula = args.state.formulas.get(cellIndex)
      const aggregate = formula?.directAggregate
      if (!formula || !aggregate || aggregate.rowStart !== 0 || formula.dependencyIndices.length !== 0) {
        continue
      }
      const key = `${aggregate.sheetName}\t${aggregate.col}\t${aggregate.colEnd}\t${aggregate.aggregateKind}`
      let group = groups.get(key)
      if (!group) {
        group = {
          sheetName: aggregate.sheetName,
          col: aggregate.col,
          colEnd: aggregate.colEnd,
          aggregateKind: aggregate.aggregateKind,
          maxRowEnd: aggregate.rowEnd,
          lastRowEnd: aggregate.rowEnd,
          formulasAreOrdered: true,
          formulas: [],
        }
        groups.set(key, group)
      } else {
        group.maxRowEnd = Math.max(group.maxRowEnd, aggregate.rowEnd)
        if (aggregate.rowEnd < group.lastRowEnd) {
          group.formulasAreOrdered = false
        }
        group.lastRowEnd = aggregate.rowEnd
      }
      group.formulas.push({
        cellIndex,
        rowEnd: aggregate.rowEnd,
        ...(aggregate.resultOffset !== undefined ? { resultOffset: aggregate.resultOffset } : {}),
      })
    }
    if (groups.size === 0) {
      return undefined
    }

    const handled = new Set<number>()
    groups.forEach((group) => {
      const sheet = args.state.workbook.getSheet(group.sheetName)
      if (!sheet) {
        return
      }
      const formulas = group.formulasAreOrdered ? group.formulas : group.formulas.toSorted((left, right) => left.rowEnd - right.rowEnd)
      let sum = 0
      let count = 0
      let averageCount = 0
      let errorCode = ErrorCode.None
      let errorCount = 0
      let minimum = Number.POSITIVE_INFINITY
      let maximum = Number.NEGATIVE_INFINITY
      let formulaIndex = 0
      let encounteredFormulaMember = false
      for (let row = 0; row <= group.maxRowEnd && !encounteredFormulaMember; row += 1) {
        for (let col = group.col; col <= group.colEnd; col += 1) {
          const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
          if (memberCellIndex !== -1) {
            if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
              encounteredFormulaMember = true
              break
            }
            const tag = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
            if (tag === ValueTag.Number) {
              const numeric = args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
              sum += numeric
              count += 1
              averageCount += 1
              minimum = Math.min(minimum, numeric)
              maximum = Math.max(maximum, numeric)
            } else if (tag === ValueTag.Boolean) {
              const numeric = (args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
              sum += numeric
              count += 1
              averageCount += 1
              minimum = Math.min(minimum, numeric)
              maximum = Math.max(maximum, numeric)
            } else if (tag === ValueTag.Empty) {
              minimum = Math.min(minimum, 0)
              maximum = Math.max(maximum, 0)
            } else if (tag === ValueTag.Error) {
              errorCode ||= (args.state.workbook.cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None
              errorCount += 1
            }
          }
        }
        while (formulaIndex < formulas.length && formulas[formulaIndex]!.rowEnd <= row) {
          const formula = formulas[formulaIndex]!
          const aggregateValue =
            group.aggregateKind === 'sum'
              ? errorCount > 0 && errorCode !== ErrorCode.None
                ? { tag: ValueTag.Error as const, code: errorCode }
                : { tag: ValueTag.Number as const, value: sum }
              : group.aggregateKind === 'count'
                ? { tag: ValueTag.Number as const, value: count }
                : group.aggregateKind === 'average'
                  ? errorCount > 0 && errorCode !== ErrorCode.None
                    ? { tag: ValueTag.Error as const, code: errorCode }
                    : averageCount === 0
                      ? { tag: ValueTag.Error as const, code: ErrorCode.Div0 }
                      : { tag: ValueTag.Number as const, value: sum / averageCount }
                  : group.aggregateKind === 'min'
                    ? { tag: ValueTag.Number as const, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
                    : { tag: ValueTag.Number as const, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
          const value =
            formula.resultOffset !== undefined && aggregateValue.tag === ValueTag.Number
              ? { tag: ValueTag.Number as const, value: aggregateValue.value + formula.resultOffset }
              : aggregateValue
          writeFormulaValue(formula.cellIndex, value)
          handled.add(formula.cellIndex)
          pushChangedCellIndex(formula.cellIndex)
          formulaIndex += 1
        }
      }
    })
    return handled.size === 0 ? undefined : handled
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
          ? evaluateInitialPrefixAggregateGroups(orderedCellIndices, pushChangedCellIndex, valueWriter.writeValue)
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
      let maxTargetCellIndex = 0
      if (!hadExistingFormulas) {
        for (let index = 0; index < refs.length; index += 1) {
          args.checkEvaluationBudget()
          const cellIndex = resolveCellIndex(initialFormulaEntryRefAt(refs, index))
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
      const deferredFormulaFamilyRuns =
        hadExistingFormulas || shouldDeferFormulaFamilyIndex ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()
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
              const cellIndex = hadExistingFormulas ? resolveCellIndex(ref) : targetCellIndices[refIndex]!
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
              args.deferFormulaFamilyIndexRebuild?.()
            } else {
              deferredFormulaFamilyRuns?.forEach(registerDeferredFormulaFamilyRun)
            }
            if (shouldDeferFormulaInstanceTable) {
              args.deferFormulaInstanceTableRebuild?.()
            }
            inlineInitialDirectScalarWriter?.flush()
          })
        }
        bindFormulaEntries()
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
      if (args.hasVolatileFormulas?.() !== false) {
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
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver()
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
    const resolveInitialTemplateFormula = createInitialTemplateFormulaResolver()
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

  const initializeHydratedPreparedCellFormulasAtNow = (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginMutationCollection()
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
      : refs.map((ref) => ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col))
    const pendingFormulaCells = hadExistingFormulas
      ? undefined
      : new Uint8Array(args.state.workbook.cellStore.capacity + reservedNewCells + 1)
    if (pendingFormulaCells) {
      for (let index = 0; index < targetCellIndices.length; index += 1) {
        pendingFormulaCells[targetCellIndices[index]!] = 1
      }
    }
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0
    const shouldDeferFormulaFamilyIndex = !hadExistingFormulas && args.deferFormulaFamilyIndexRebuild !== undefined
    const deferredFormulaFamilyRuns =
      hadExistingFormulas || shouldDeferFormulaFamilyIndex ? undefined : new Map<string, DeferredInitialFormulaFamilyRun>()

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      const valueWriter = createInitialFormulaValueWriter(args)
      const bindFormulaEntries = (): void => {
        args.state.workbook.withBatchedColumnVersionUpdates(() => {
          for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
            const ref = refs[refIndex]!
            const cellIndex = hadExistingFormulas
              ? (ref.cellIndex ?? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col))
              : targetCellIndices[refIndex]!
            const ownerSheetName = resolveSheetName(ref.sheetId)
            topologyChanged =
              args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId, {
                deferFamilyRegistration: shouldDeferFormulaFamilyIndex || deferredFormulaFamilyRuns !== undefined,
              }) || topologyChanged
            noteDeferredFormulaFamilyRunMember(deferredFormulaFamilyRuns, {
              cellIndex,
              sheetId: ref.sheetId,
              row: ref.row,
              col: ref.col,
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
            args.deferFormulaFamilyIndexRebuild?.()
          } else {
            deferredFormulaFamilyRuns?.forEach(registerDeferredFormulaFamilyRun)
          }
          valueWriter.flush()
        })
      }
      bindFormulaEntries()
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
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
        args.rebuildTopoRanks()
        args.detectCycles()
      }
    }
    const lastMetrics = args.state.getLastMetrics()
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: 0,
      compileMs,
      recalcMs: 0,
    })
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
