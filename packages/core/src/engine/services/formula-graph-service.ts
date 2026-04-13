import { Effect } from "effect";
import { ErrorCode, FormulaMode } from "@bilig/protocol";
import { CellFlags } from "../../cell-store.js";
import type { CycleDetector } from "../../cycle-detection.js";
import { makeCellEntity } from "../../entity-ids.js";
import { growUint32 } from "../../engine-buffer-utils.js";
import { errorValue } from "../../engine-value-utils.js";
import type { EngineRuntimeState, U32 } from "../runtime-state.js";
import { EngineFormulaGraphError } from "../errors.js";
import type { Float64Arena, Uint32Arena } from "@bilig/formula";

export interface EngineFormulaGraphService {
  readonly rebuildTopoRanks: () => Effect.Effect<void, EngineFormulaGraphError>;
  readonly detectCycles: () => Effect.Effect<void, EngineFormulaGraphError>;
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => Effect.Effect<void, EngineFormulaGraphError>;
  readonly scheduleWasmProgramSync: () => Effect.Effect<void, EngineFormulaGraphError>;
  readonly flushWasmProgramSync: () => Effect.Effect<void, EngineFormulaGraphError>;
  readonly rebuildTopoRanksNow: () => void;
  readonly detectCyclesNow: () => void;
  readonly scheduleWasmProgramSyncNow: () => void;
  readonly flushWasmProgramSyncNow: () => void;
}

function graphErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

const SYNC_WASM_INIT_FORMULA_THRESHOLD = 64;

export function createEngineFormulaGraphService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "formulas" | "ranges" | "wasm">;
  readonly cycleDetector: CycleDetector;
  readonly programArena: Uint32Arena;
  readonly constantArena: Float64Arena;
  readonly rangeListArena: Uint32Arena;
  readonly getTopoIndegree: () => U32;
  readonly setTopoIndegree: (next: U32) => void;
  readonly getTopoQueue: () => U32;
  readonly setTopoQueue: (next: U32) => void;
  readonly getWasmProgramTargets: () => U32;
  readonly setWasmProgramTargets: (next: U32) => void;
  readonly getWasmProgramOffsets: () => U32;
  readonly setWasmProgramOffsets: (next: U32) => void;
  readonly getWasmProgramLengths: () => U32;
  readonly setWasmProgramLengths: (next: U32) => void;
  readonly getWasmConstantOffsets: () => U32;
  readonly setWasmConstantOffsets: (next: U32) => void;
  readonly getWasmConstantLengths: () => U32;
  readonly setWasmConstantLengths: (next: U32) => void;
  readonly getWasmRangeOffsets: () => U32;
  readonly setWasmRangeOffsets: (next: U32) => void;
  readonly getWasmRangeLengths: () => U32;
  readonly setWasmRangeLengths: (next: U32) => void;
  readonly getWasmRangeRowCounts: () => U32;
  readonly setWasmRangeRowCounts: (next: U32) => void;
  readonly getWasmRangeColCounts: () => U32;
  readonly setWasmRangeColCounts: (next: U32) => void;
  readonly getBatchMutationDepth: () => number;
  readonly getWasmProgramSyncPending: () => boolean;
  readonly setWasmProgramSyncPending: (next: boolean) => void;
  readonly notifyCellValueWritten: (cellIndex: number) => void;
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => void;
  readonly collectFormulaDependents: (entityId: number) => Uint32Array;
}): EngineFormulaGraphService {
  const ensureTopoScratchCapacity = (cellSize: number): void => {
    if (cellSize > args.getTopoIndegree().length) {
      args.setTopoIndegree(growUint32(args.getTopoIndegree(), cellSize));
    }
    if (cellSize > args.getTopoQueue().length) {
      args.setTopoQueue(growUint32(args.getTopoQueue(), cellSize));
    }
  };

  const ensureWasmProgramScratchCapacity = (formulaSize: number, rangeSize: number): void => {
    if (formulaSize > args.getWasmProgramTargets().length) {
      args.setWasmProgramTargets(growUint32(args.getWasmProgramTargets(), formulaSize));
    }
    if (formulaSize > args.getWasmProgramOffsets().length) {
      args.setWasmProgramOffsets(growUint32(args.getWasmProgramOffsets(), formulaSize));
    }
    if (formulaSize > args.getWasmProgramLengths().length) {
      args.setWasmProgramLengths(growUint32(args.getWasmProgramLengths(), formulaSize));
    }
    if (formulaSize > args.getWasmConstantOffsets().length) {
      args.setWasmConstantOffsets(growUint32(args.getWasmConstantOffsets(), formulaSize));
    }
    if (formulaSize > args.getWasmConstantLengths().length) {
      args.setWasmConstantLengths(growUint32(args.getWasmConstantLengths(), formulaSize));
    }
    if (rangeSize > args.getWasmRangeOffsets().length) {
      args.setWasmRangeOffsets(growUint32(args.getWasmRangeOffsets(), rangeSize));
    }
    if (rangeSize > args.getWasmRangeLengths().length) {
      args.setWasmRangeLengths(growUint32(args.getWasmRangeLengths(), rangeSize));
    }
    if (rangeSize > args.getWasmRangeRowCounts().length) {
      args.setWasmRangeRowCounts(growUint32(args.getWasmRangeRowCounts(), rangeSize));
    }
    if (rangeSize > args.getWasmRangeColCounts().length) {
      args.setWasmRangeColCounts(growUint32(args.getWasmRangeColCounts(), rangeSize));
    }
  };

  const rebuildTopoRanksNow = (): void => {
    const requiredCellCapacity = args.state.workbook.cellStore.size + 1;
    ensureTopoScratchCapacity(requiredCellCapacity);

    let queueLength = 0;
    args.state.formulas.forEach((_formula, cellIndex) => {
      args.getTopoIndegree()[cellIndex] = 0;
      args.state.workbook.cellStore.topoRanks[cellIndex] = 0;
    });
    args.state.formulas.forEach((formula, cellIndex) => {
      for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
        const dependency = formula.dependencyIndices[index]!;
        if ((args.state.workbook.cellStore.formulaIds[dependency] ?? 0) !== 0) {
          args.getTopoIndegree()[cellIndex] = (args.getTopoIndegree()[cellIndex] ?? 0) + 1;
        }
      }
    });
    args.state.formulas.forEach((_formula, cellIndex) => {
      if ((args.getTopoIndegree()[cellIndex] ?? 0) === 0) {
        args.getTopoQueue()[queueLength] = cellIndex;
        queueLength += 1;
      }
    });

    let rank = 0;
    for (let queueIndex = 0; queueIndex < queueLength; queueIndex += 1) {
      const cellIndex = args.getTopoQueue()[queueIndex]!;
      args.state.workbook.cellStore.topoRanks[cellIndex] = rank++;
      const dependents = args.collectFormulaDependents(makeCellEntity(cellIndex));
      for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
        const dependent = dependents[dependentIndex]!;
        if ((args.state.workbook.cellStore.formulaIds[dependent] ?? 0) === 0) {
          continue;
        }
        const next = (args.getTopoIndegree()[dependent] ?? 0) - 1;
        args.getTopoIndegree()[dependent] = next;
        if (next === 0) {
          args.getTopoQueue()[queueLength] = dependent;
          queueLength += 1;
        }
      }
    }
  };

  const detectCyclesNow = (): void => {
    const result = args.cycleDetector.detect(
      args.state.formulas.keys(),
      args.state.workbook.cellStore.size + 1,
      (cellIndex, fn) => args.forEachFormulaDependencyCell(cellIndex, fn),
      (cellIndex) => args.state.formulas.has(cellIndex),
    );

    args.state.formulas.forEach((_formula, cellIndex) => {
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.InCycle;
      args.state.workbook.cellStore.cycleGroupIds[cellIndex] = -1;
    });

    for (let index = 0; index < result.cycleMemberCount; index += 1) {
      const cellIndex = result.cycleMembers[index]!;
      args.state.workbook.cellStore.flags[cellIndex] =
        (args.state.workbook.cellStore.flags[cellIndex] ?? 0) | CellFlags.InCycle;
      args.state.workbook.cellStore.cycleGroupIds[cellIndex] = result.cycleGroups[cellIndex] ?? -1;
      args.state.workbook.cellStore.setValue(cellIndex, errorValue(ErrorCode.Cycle));
      args.notifyCellValueWritten(cellIndex);
    }
  };

  const syncWasmProgramsNow = (): void => {
    args.programArena.reset();
    args.constantArena.reset();
    args.rangeListArena.reset();

    let wasmFormulaCount = 0;
    args.state.formulas.forEach((formula) => {
      if (formula.compiled.mode === FormulaMode.WasmFastPath) {
        wasmFormulaCount += 1;
      }
    });
    ensureWasmProgramScratchCapacity(
      Math.max(wasmFormulaCount, 1),
      Math.max(args.state.ranges.size, 1),
    );

    let formulaIndex = 0;
    args.state.formulas.forEach((formula) => {
      if (formula.compiled.mode !== FormulaMode.WasmFastPath) {
        return;
      }
      const programSlice = args.programArena.append(formula.runtimeProgram);
      const constantSlice = args.constantArena.append(formula.constants);
      const rangeSlice = args.rangeListArena.append(formula.rangeDependencies);

      formula.programOffset = programSlice.offset;
      formula.programLength = programSlice.length;
      formula.constNumberOffset = constantSlice.offset;
      formula.constNumberLength = constantSlice.length;
      formula.rangeListOffset = rangeSlice.offset;
      formula.rangeListLength = rangeSlice.length;

      args.getWasmProgramTargets()[formulaIndex] = formula.cellIndex;
      args.getWasmProgramOffsets()[formulaIndex] = programSlice.offset;
      args.getWasmProgramLengths()[formulaIndex] = programSlice.length;
      args.getWasmConstantOffsets()[formulaIndex] = constantSlice.offset;
      args.getWasmConstantLengths()[formulaIndex] = constantSlice.length;
      formulaIndex += 1;
    });

    args.state.wasm.uploadFormulas({
      targets: args.getWasmProgramTargets().subarray(0, wasmFormulaCount),
      programs: args.programArena.view(),
      programOffsets: args.getWasmProgramOffsets().subarray(0, wasmFormulaCount),
      programLengths: args.getWasmProgramLengths().subarray(0, wasmFormulaCount),
      constants: args.constantArena.view(),
      constantOffsets: args.getWasmConstantOffsets().subarray(0, wasmFormulaCount),
      constantLengths: args.getWasmConstantLengths().subarray(0, wasmFormulaCount),
    });

    const rangeCapacity = Math.max(args.state.ranges.size, 1);
    if (args.state.ranges.size === 0) {
      args.getWasmRangeOffsets()[0] = 0;
      args.getWasmRangeLengths()[0] = 0;
      args.getWasmRangeRowCounts()[0] = 0;
      args.getWasmRangeColCounts()[0] = 0;
    }
    for (let rangeIndex = 0; rangeIndex < args.state.ranges.size; rangeIndex += 1) {
      const descriptor = args.state.ranges.getDescriptor(rangeIndex);
      args.getWasmRangeOffsets()[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.membersOffset : 0;
      args.getWasmRangeLengths()[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.membersLength : 0;
      args.getWasmRangeRowCounts()[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.row2 - descriptor.row1 + 1 : 0;
      args.getWasmRangeColCounts()[rangeIndex] =
        descriptor.refCount > 0 ? descriptor.col2 - descriptor.col1 + 1 : 0;
    }

    args.state.wasm.uploadRanges({
      members: args.state.ranges.getMemberPoolView(),
      offsets: args.getWasmRangeOffsets().subarray(0, rangeCapacity),
      lengths: args.getWasmRangeLengths().subarray(0, rangeCapacity),
      rowCounts: args.getWasmRangeRowCounts().subarray(0, rangeCapacity),
      colCounts: args.getWasmRangeColCounts().subarray(0, rangeCapacity),
    });
  };

  const scheduleWasmProgramSyncNow = (): void => {
    if (args.getBatchMutationDepth() > 0) {
      args.setWasmProgramSyncPending(true);
      return;
    }
    if (
      !args.state.wasm.ready &&
      (args.state.formulas.size < SYNC_WASM_INIT_FORMULA_THRESHOLD ||
        !args.state.wasm.initSyncIfPossible())
    ) {
      args.setWasmProgramSyncPending(true);
      return;
    }
    syncWasmProgramsNow();
  };

  const flushWasmProgramSyncNow = (): void => {
    if (!args.getWasmProgramSyncPending()) {
      return;
    }
    if (
      !args.state.wasm.ready &&
      (args.state.formulas.size < SYNC_WASM_INIT_FORMULA_THRESHOLD ||
        !args.state.wasm.initSyncIfPossible())
    ) {
      return;
    }
    args.setWasmProgramSyncPending(false);
    syncWasmProgramsNow();
  };

  return {
    rebuildTopoRanks() {
      return Effect.try({
        try: () => {
          rebuildTopoRanksNow();
        },
        catch: (cause) =>
          new EngineFormulaGraphError({
            message: graphErrorMessage("Failed to rebuild topo ranks", cause),
            cause,
          }),
      });
    },
    detectCycles() {
      return Effect.try({
        try: () => {
          detectCyclesNow();
        },
        catch: (cause) =>
          new EngineFormulaGraphError({
            message: graphErrorMessage("Failed to detect formula cycles", cause),
            cause,
          }),
      });
    },
    forEachFormulaDependencyCell(cellIndex, fn) {
      return Effect.try({
        try: () => {
          args.forEachFormulaDependencyCell(cellIndex, fn);
        },
        catch: (cause) =>
          new EngineFormulaGraphError({
            message: graphErrorMessage("Failed to iterate formula dependencies", cause),
            cause,
          }),
      });
    },
    scheduleWasmProgramSync() {
      return Effect.try({
        try: () => {
          scheduleWasmProgramSyncNow();
        },
        catch: (cause) =>
          new EngineFormulaGraphError({
            message: graphErrorMessage("Failed to schedule wasm program sync", cause),
            cause,
          }),
      });
    },
    flushWasmProgramSync() {
      return Effect.try({
        try: () => {
          flushWasmProgramSyncNow();
        },
        catch: (cause) =>
          new EngineFormulaGraphError({
            message: graphErrorMessage("Failed to flush wasm program sync", cause),
            cause,
          }),
      });
    },
    rebuildTopoRanksNow,
    detectCyclesNow,
    scheduleWasmProgramSyncNow,
    flushWasmProgramSyncNow,
  };
}
