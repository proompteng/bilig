import { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'

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
  readonly recalculatePreordered: (
    changedRoots: readonly number[] | U32,
    orderedFormulaCellIndices: readonly number[] | U32,
    orderedFormulaCount: number,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly flushWasmProgramSync: () => void
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
      compileMs += performance.now() - compileStarted
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
      args.flushWasmProgramSync()
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
      canAssignTopoInBatch && !hadExistingFormulas && orderedPreparedCellIndices.length > 0
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
