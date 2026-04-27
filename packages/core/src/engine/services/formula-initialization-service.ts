import { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'

const INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT = 16_384

type InitialPrefixAggregateKind = 'sum' | 'count'

interface InitialPrefixAggregateGroup {
  readonly sheetName: string
  readonly col: number
  readonly aggregateKind: InitialPrefixAggregateKind
  maxRowEnd: number
  readonly formulas: Array<{ cellIndex: number; rowEnd: number }>
}

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export interface EngineFormulaInitializationService {
  readonly initializeCellFormulasAt: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeCellFormulasAtNow: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializePreparedCellFormulasAt: (
    refs: readonly PreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializePreparedCellFormulasAtNow: (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeHydratedPreparedCellFormulasAtNow: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => void
}

export interface PreparedFormulaInitializationRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
}

export interface HydratedPreparedFormulaInitializationRef extends PreparedFormulaInitializationRef {
  readonly value: CellValue
}

export function createEngineFormulaInitializationService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'counters' | 'getLastMetrics' | 'setLastMetrics'>
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
  ) => boolean
  readonly compileTemplateFormula: (source: string, row: number, col: number) => FormulaTemplateResolution
  readonly clearTemplateFormulaCache: () => void
  readonly removeFormula: (cellIndex: number) => boolean
  readonly setInvalidFormulaValue: (cellIndex: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly markVolatileFormulasChanged: (count: number) => number
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

  const canEvaluateInitialDirectFormula = (cellIndex: number): boolean => {
    const formula = args.state.formulas.get(cellIndex)
    return (
      formula !== undefined &&
      !formula.compiled.volatile &&
      !formula.compiled.producesSpill &&
      (formula.directAggregate !== undefined ||
        formula.directCriteria !== undefined ||
        formula.directLookup !== undefined ||
        formula.directScalar !== undefined)
    )
  }

  const evaluateInitialPrefixAggregateGroups = (
    orderedCellIndices: readonly number[],
    changedCellIndices: number[],
  ): Set<number> | undefined => {
    const groups = new Map<string, InitialPrefixAggregateGroup>()
    for (let index = 0; index < orderedCellIndices.length; index += 1) {
      const cellIndex = orderedCellIndices[index]!
      const formula = args.state.formulas.get(cellIndex)
      const aggregate = formula?.directAggregate
      if (
        !formula ||
        !aggregate ||
        aggregate.rowStart !== 0 ||
        formula.dependencyIndices.length !== 0 ||
        (aggregate.aggregateKind !== 'sum' && aggregate.aggregateKind !== 'count')
      ) {
        continue
      }
      const key = `${aggregate.sheetName}\t${aggregate.col}\t${aggregate.aggregateKind}`
      let group = groups.get(key)
      if (!group) {
        group = {
          sheetName: aggregate.sheetName,
          col: aggregate.col,
          aggregateKind: aggregate.aggregateKind,
          maxRowEnd: aggregate.rowEnd,
          formulas: [],
        }
        groups.set(key, group)
      } else {
        group.maxRowEnd = Math.max(group.maxRowEnd, aggregate.rowEnd)
      }
      group.formulas.push({ cellIndex, rowEnd: aggregate.rowEnd })
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
      const formulas = group.formulas.toSorted((left, right) => left.rowEnd - right.rowEnd)
      const assignments: Array<{ cellIndex: number; value: number }> = []
      let sum = 0
      let count = 0
      let formulaIndex = 0
      let canMaterializeGroup = true
      for (let row = 0; row <= group.maxRowEnd; row += 1) {
        const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, group.col) : sheet.grid.get(row, group.col)
        if (memberCellIndex !== -1) {
          if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
            canMaterializeGroup = false
            break
          }
          const tag = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
          if (tag === ValueTag.Number) {
            sum += args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
            count += 1
          } else if (tag === ValueTag.Boolean) {
            sum += (args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
            count += 1
          } else if (tag === ValueTag.Error && group.aggregateKind === 'sum') {
            canMaterializeGroup = false
            break
          }
        }
        while (formulaIndex < formulas.length && formulas[formulaIndex]!.rowEnd <= row) {
          assignments.push({
            cellIndex: formulas[formulaIndex]!.cellIndex,
            value: group.aggregateKind === 'sum' ? sum : count,
          })
          formulaIndex += 1
        }
      }
      if (!canMaterializeGroup || assignments.length !== formulas.length) {
        return
      }
      for (let index = 0; index < assignments.length; index += 1) {
        const assignment = assignments[index]!
        args.writeHydratedFormulaValue(assignment.cellIndex, {
          tag: ValueTag.Number,
          value: assignment.value,
        })
        handled.add(assignment.cellIndex)
        changedCellIndices.push(assignment.cellIndex)
      }
    })
    return handled.size === 0 ? undefined : handled
  }

  const coerceInitialDirectScalarCell = (
    cellIndex: number,
  ): { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode } | undefined => {
    const cellStore = args.state.workbook.cellStore
    const tag = (cellStore.tags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    switch (tag) {
      case ValueTag.Number:
        return { kind: 'number', value: cellStore.numbers[cellIndex] ?? 0 }
      case ValueTag.Boolean:
        return { kind: 'number', value: (cellStore.numbers[cellIndex] ?? 0) !== 0 ? 1 : 0 }
      case ValueTag.Empty:
        return { kind: 'number', value: 0 }
      case ValueTag.Error:
        return { kind: 'error', code: (cellStore.errors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None }
      case ValueTag.String:
        return { kind: 'error', code: ErrorCode.Value }
      default:
        return undefined
    }
  }

  const readInitialDirectScalarOperand = (
    operand: RuntimeDirectScalarOperand,
  ): { kind: 'number'; value: number } | { kind: 'error'; code: ErrorCode } | undefined => {
    switch (operand.kind) {
      case 'literal-number':
        return { kind: 'number', value: operand.value }
      case 'error':
        return { kind: 'error', code: operand.code }
      case 'cell':
        return coerceInitialDirectScalarCell(operand.cellIndex)
    }
  }

  const evaluateInitialDirectScalar = (directScalar: RuntimeDirectScalarDescriptor): CellValue | undefined => {
    if (directScalar.kind === 'abs') {
      const operand = readInitialDirectScalarOperand(directScalar.operand)
      if (!operand) {
        return undefined
      }
      return operand.kind === 'error'
        ? { tag: ValueTag.Error, code: operand.code }
        : { tag: ValueTag.Number, value: Math.abs(operand.value) }
    }
    const left = readInitialDirectScalarOperand(directScalar.left)
    const right = readInitialDirectScalarOperand(directScalar.right)
    if (!left || !right) {
      return undefined
    }
    if (left.kind === 'error') {
      return { tag: ValueTag.Error, code: left.code }
    }
    if (right.kind === 'error') {
      return { tag: ValueTag.Error, code: right.code }
    }
    switch (directScalar.operator) {
      case '+':
        return { tag: ValueTag.Number, value: left.value + right.value }
      case '-':
        return { tag: ValueTag.Number, value: left.value - right.value }
      case '*':
        return { tag: ValueTag.Number, value: left.value * right.value }
      case '/':
        return right.value === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: left.value / right.value }
    }
  }

  const evaluateInitialDirectFormulas = (orderedCellIndices: readonly number[]): U32 | undefined => {
    if (
      orderedCellIndices.length === 0 ||
      orderedCellIndices.length > INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT ||
      !orderedCellIndices.every(canEvaluateInitialDirectFormula)
    ) {
      return undefined
    }
    const changedCellIndices: number[] = []
    args.state.workbook.withBatchedColumnVersionUpdates(() => {
      const prefixAggregateHandled = evaluateInitialPrefixAggregateGroups(orderedCellIndices, changedCellIndices)
      for (let index = 0; index < orderedCellIndices.length; index += 1) {
        const cellIndex = orderedCellIndices[index]!
        if (prefixAggregateHandled?.has(cellIndex)) {
          continue
        }
        const formula = args.state.formulas.get(cellIndex)
        if (formula?.directScalar !== undefined) {
          const value = evaluateInitialDirectScalar(formula.directScalar)
          if (value !== undefined) {
            args.writeHydratedFormulaValue(cellIndex, value)
            changedCellIndices.push(cellIndex)
            continue
          }
        }
        const changedSpillIndices = args.evaluateDirectFormula(cellIndex)
        changedCellIndices.push(cellIndex)
        if (changedSpillIndices) {
          for (let spillIndex = 0; spillIndex < changedSpillIndices.length; spillIndex += 1) {
            changedCellIndices.push(changedSpillIndices[spillIndex]!)
          }
        }
      }
    })
    args.deferKernelSync(Uint32Array.from(changedCellIndices))
    addEngineCounter(args.state.counters, 'directFormulaInitialEvaluations', orderedCellIndices.length)
    return Uint32Array.from(changedCellIndices)
  }

  const initializeFormulaEntriesNow = <Entry>(
    refs: readonly Entry[],
    potentialNewCells: number | undefined,
    resolveCellIndex: (ref: Entry) => number,
    resolveEntry: (
      ref: Entry,
      cellIndex: number,
    ) => {
      cellIndex: number
      ownerSheetName: string
      source: string
      compiled: CompiledFormula
      templateId?: number
    },
  ): void => {
    if (refs.length === 0) {
      return
    }

    args.beginMutationCollection()
    let hadCycleMembersBefore: boolean | undefined
    const hadCycleMembersBeforeNow = (): boolean => (hadCycleMembersBefore ??= hasCycleMembersNow())
    let changedInputCount = 0
    let formulaChangedCount = 0
    let topologyChanged = false
    let compileMs = 0
    const reservedNewCells = Math.max(potentialNewCells ?? refs.length, refs.length)
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const targetCellIndices = hadExistingFormulas ? [] : refs.map((ref) => resolveCellIndex(ref))
    const pendingFormulaCellIndices = hadExistingFormulas ? undefined : new Set<number>(targetCellIndices)
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0
    const orderedPreparedCellIndices: number[] = []
    let canUseInitialDirectEvaluation = false

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
          const ref = refs[refIndex]!
          const cellIndex = hadExistingFormulas ? resolveCellIndex(ref) : targetCellIndices[refIndex]!
          try {
            const prepared = resolveEntry(ref, cellIndex)
            args.bindPreparedFormula(prepared.cellIndex, prepared.ownerSheetName, prepared.source, prepared.compiled, prepared.templateId)
            formulaChangedCount = args.markFormulaChanged(prepared.cellIndex, formulaChangedCount)
            topologyChanged = true
            orderedPreparedCellIndices.push(prepared.cellIndex)
            if (canAssignTopoInBatch && pendingFormulaCellIndices) {
              const runtimeFormula = args.state.formulas.get(prepared.cellIndex)
              if (
                !runtimeFormula ||
                runtimeFormula.dependencyIndices.some((dependencyCellIndex) => pendingFormulaCellIndices.has(dependencyCellIndex))
              ) {
                canAssignTopoInBatch = false
              } else {
                args.state.workbook.cellStore.topoRanks[prepared.cellIndex] = nextTopoRank
                nextTopoRank += 1
              }
            }
          } catch {
            topologyChanged = args.removeFormula(cellIndex) || topologyChanged
            args.setInvalidFormulaValue(cellIndex)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          }
          pendingFormulaCellIndices?.delete(cellIndex)
        }
        const reboundCount = formulaChangedCount
        formulaChangedCount = args.syncDynamicRanges(formulaChangedCount)
        topologyChanged = topologyChanged || formulaChangedCount !== reboundCount
      })
      canUseInitialDirectEvaluation =
        canAssignTopoInBatch &&
        !hadExistingFormulas &&
        changedInputCount === 0 &&
        orderedPreparedCellIndices.length <= INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT &&
        orderedPreparedCellIndices.length === refs.length &&
        orderedPreparedCellIndices.every(canEvaluateInitialDirectFormula)
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (topologyChanged && !(canAssignTopoInBatch && !hadExistingFormulas)) {
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
    args.prepareRegionQueryIndices()
    formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount)
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    const changedRoots = args.composeMutationRoots(changedInputCount, formulaChangedCount)
    let recalculated =
      canUseInitialDirectEvaluation && formulaChangedCount === orderedPreparedCellIndices.length
        ? (evaluateInitialDirectFormulas(orderedPreparedCellIndices) ?? args.recalculate(changedRoots, changedInputArray))
        : canAssignTopoInBatch && !hadExistingFormulas && orderedPreparedCellIndices.length > 0
          ? args.recalculatePreordered(changedRoots, orderedPreparedCellIndices, orderedPreparedCellIndices.length, changedInputArray)
          : args.recalculate(changedRoots, changedInputArray)
    recalculated = args.reconcilePivotOutputs(recalculated, false)
    void recalculated
    const lastMetrics = args.state.getLastMetrics()
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    })
  }

  const initializeCellFormulasAtNow = (refs: readonly EngineCellMutationRef[], potentialNewCells?: number): void => {
    initializeFormulaEntriesNow(
      refs,
      potentialNewCells,
      (ref) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        return args.ensureCellTrackedByCoords(ref.sheetId, ref.mutation.row, ref.mutation.col)
      },
      (ref, cellIndex) => {
        if (ref.mutation.kind !== 'setCellFormula') {
          throw new Error('initializeCellFormulasAt only supports setCellFormula coordinate mutations')
        }
        const ownerSheetName = resolveSheetName(ref.sheetId)
        const template = args.compileTemplateFormula(ref.mutation.formula, ref.mutation.row, ref.mutation.col)
        return {
          cellIndex,
          ownerSheetName,
          source: ref.mutation.formula,
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
      (ref) => args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col),
      (ref, cellIndex) => ({
        cellIndex,
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
    const reservedNewCells = Math.max(potentialNewCells ?? refs.length, refs.length)
    const hadExistingFormulas = args.state.formulas.size > 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1)
    args.resetMaterializedCellScratch(reservedNewCells)
    const targetCellIndices = hadExistingFormulas ? [] : refs.map((ref) => args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col))
    const pendingFormulaCellIndices = hadExistingFormulas ? undefined : new Set<number>(targetCellIndices)
    let canAssignTopoInBatch = !hadExistingFormulas
    let nextTopoRank = 0

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.clearTemplateFormulaCache()
      const compileStarted = performance.now()
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
          const ref = refs[refIndex]!
          const cellIndex = hadExistingFormulas
            ? args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)
            : targetCellIndices[refIndex]!
          const ownerSheetName = resolveSheetName(ref.sheetId)
          topologyChanged = args.bindPreparedFormula(cellIndex, ownerSheetName, ref.source, ref.compiled, ref.templateId) || topologyChanged
          args.writeHydratedFormulaValue(cellIndex, ref.value)
          if (canAssignTopoInBatch && pendingFormulaCellIndices) {
            const runtimeFormula = args.state.formulas.get(cellIndex)
            if (
              !runtimeFormula ||
              runtimeFormula.dependencyIndices.some((dependencyCellIndex) => pendingFormulaCellIndices.has(dependencyCellIndex))
            ) {
              canAssignTopoInBatch = false
            } else {
              args.state.workbook.cellStore.topoRanks[cellIndex] = nextTopoRank
              nextTopoRank += 1
            }
          }
          pendingFormulaCellIndices?.delete(cellIndex)
        }
      })
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
              : refs.map((ref) => args.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)),
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
    initializePreparedCellFormulasAtNow,
    initializeHydratedPreparedCellFormulasAtNow,
  }
}
