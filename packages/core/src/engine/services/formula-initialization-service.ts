import { Effect } from "effect";
import type { EngineCellMutationRef } from "../../cell-mutations-at.js";
import { CellFlags } from "../../cell-store.js";
import type { EngineRuntimeState, U32 } from "../runtime-state.js";
import { EngineMutationError } from "../errors.js";

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

export interface EngineFormulaInitializationService {
  readonly initializeCellFormulasAt: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>;
  readonly initializeCellFormulasAtNow: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => void;
}

export function createEngineFormulaInitializationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    "workbook" | "formulas" | "getLastMetrics" | "setLastMetrics"
  >;
  readonly beginMutationCollection: () => void;
  readonly ensureRecalcScratchCapacity: (size: number) => void;
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number;
  readonly resetMaterializedCellScratch: (expectedSize: number) => void;
  readonly bindFormula: (cellIndex: number, ownerSheetName: string, source: string) => void;
  readonly bindPreparedFormula: (
    cellIndex: number,
    ownerSheetName: string,
    source: string,
    compiled: import("@bilig/formula").CompiledFormula,
  ) => void;
  readonly compileTemplateFormula: (
    source: string,
    row: number,
    col: number,
  ) => import("@bilig/formula").CompiledFormula;
  readonly clearTemplateFormulaCache: () => void;
  readonly removeFormula: (cellIndex: number) => boolean;
  readonly setInvalidFormulaValue: (cellIndex: number) => void;
  readonly markInputChanged: (cellIndex: number, count: number) => number;
  readonly markFormulaChanged: (cellIndex: number, count: number) => number;
  readonly markVolatileFormulasChanged: (count: number) => number;
  readonly syncDynamicRanges: (formulaChangedCount: number) => number;
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32;
  readonly getChangedInputBuffer: () => U32;
  readonly rebuildTopoRanks: () => void;
  readonly detectCycles: () => void;
  readonly recalculate: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32;
  readonly reconcilePivotOutputs: (baseChanged: U32, forceAllPivots?: boolean) => U32;
  readonly getBatchMutationDepth: () => number;
  readonly setBatchMutationDepth: (next: number) => void;
  readonly flushWasmProgramSync: () => void;
}): EngineFormulaInitializationService {
  const initializeCellFormulasAtNow = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ): void => {
    if (refs.length === 0) {
      return;
    }

    args.beginMutationCollection();
    let changedInputCount = 0;
    let formulaChangedCount = 0;
    let topologyChanged = false;
    let compileMs = 0;
    const reservedNewCells = Math.max(potentialNewCells ?? refs.length, refs.length);
    args.state.workbook.cellStore.ensureCapacity(
      args.state.workbook.cellStore.size + reservedNewCells,
    );
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.capacity + 1);
    args.resetMaterializedCellScratch(reservedNewCells);

    const sheetNameById = new Map<number, string>();
    const resolveSheetName = (sheetId: number): string => {
      const cached = sheetNameById.get(sheetId);
      if (cached !== undefined) {
        return cached;
      }
      const sheet = args.state.workbook.getSheetById(sheetId);
      if (!sheet) {
        throw new Error(`Unknown sheet id: ${sheetId}`);
      }
      sheetNameById.set(sheetId, sheet.name);
      return sheet.name;
    };

    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1);
    try {
      args.clearTemplateFormulaCache();
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        refs.forEach((ref) => {
          if (ref.mutation.kind !== "setCellFormula") {
            throw new Error(
              "initializeCellFormulasAt only supports setCellFormula coordinate mutations",
            );
          }
          const sheetName = resolveSheetName(ref.sheetId);
          const cellIndex = args.ensureCellTrackedByCoords(
            ref.sheetId,
            ref.mutation.row,
            ref.mutation.col,
          );
          const compileStarted = performance.now();
          try {
            args.bindPreparedFormula(
              cellIndex,
              sheetName,
              ref.mutation.formula,
              args.compileTemplateFormula(ref.mutation.formula, ref.mutation.row, ref.mutation.col),
            );
            compileMs += performance.now() - compileStarted;
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
            topologyChanged = true;
          } catch {
            compileMs += performance.now() - compileStarted;
            topologyChanged = args.removeFormula(cellIndex) || topologyChanged;
            args.setInvalidFormulaValue(cellIndex);
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
          }
        });
        const reboundCount = formulaChangedCount;
        formulaChangedCount = args.syncDynamicRanges(formulaChangedCount);
        topologyChanged = topologyChanged || formulaChangedCount !== reboundCount;
      });
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1);
      args.flushWasmProgramSync();
    }

    if (topologyChanged) {
      args.rebuildTopoRanks();
      args.detectCycles();
      args.state.formulas.forEach((_formula, cellIndex) => {
        if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
        }
      });
    }
    formulaChangedCount = args.markVolatileFormulasChanged(formulaChangedCount);
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount);
    let recalculated = args.recalculate(
      args.composeMutationRoots(changedInputCount, formulaChangedCount),
      changedInputArray,
    );
    recalculated = args.reconcilePivotOutputs(recalculated, false);
    void recalculated;
    const lastMetrics = args.state.getLastMetrics();
    args.state.setLastMetrics({
      ...lastMetrics,
      batchId: lastMetrics.batchId + 1,
      changedInputCount: changedInputCount + formulaChangedCount,
      compileMs,
    });
  };

  return {
    initializeCellFormulasAt(refs, potentialNewCells) {
      return Effect.try({
        try: () => {
          initializeCellFormulasAtNow(refs, potentialNewCells);
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage("Failed to initialize cell formulas", cause),
            cause,
          }),
      });
    },
    initializeCellFormulasAtNow,
  };
}
