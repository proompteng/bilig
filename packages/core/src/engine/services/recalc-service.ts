import { Effect } from "effect";
import { FormulaMode, ValueTag, type CellSnapshot } from "@bilig/protocol";
import { makeCellKey } from "../../workbook-store.js";
import { CellFlags } from "../../cell-store.js";
import type {
  EngineRuntimeState,
  RecalcVolatileState,
  RuntimeFormula,
  U32,
} from "../runtime-state.js";
import { EngineRecalcError } from "../errors.js";
import type { WorkbookPivotRecord } from "../../workbook-store.js";
import { parseCellAddress, utcDateToExcelSerial } from "@bilig/formula";

export interface DirtyRegion {
  readonly sheetName: string;
  readonly rowStart: number;
  readonly rowEnd: number;
  readonly colStart: number;
  readonly colEnd: number;
}

export interface EngineRecalcService {
  readonly recalculateNow: () => Effect.Effect<number[], EngineRecalcError>;
  readonly recalculateDirty: (
    dirtyRegions: ReadonlyArray<DirtyRegion>,
  ) => Effect.Effect<number[], EngineRecalcError>;
  readonly recalculateDifferential: () => Effect.Effect<
    { js: CellSnapshot[]; wasm: CellSnapshot[]; drift: string[] },
    EngineRecalcError
  >;
  readonly recalculate: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => Effect.Effect<U32, EngineRecalcError>;
  readonly reconcilePivotOutputs: (
    baseChanged: U32,
    forceAllPivots?: boolean,
  ) => Effect.Effect<U32, EngineRecalcError>;
  readonly recalculateNowSync: (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots?: readonly number[] | U32,
  ) => U32;
  readonly reconcilePivotOutputsNow: (baseChanged: U32, forceAllPivots?: boolean) => U32;
}

function createRecalcVolatileState(now: () => Date): RecalcVolatileState {
  return {
    nowSerial: utcDateToExcelSerial(now()),
    randomValues: [],
    randomCursor: 0,
  };
}

function ensureVolatileRandomValues(
  state: RecalcVolatileState,
  count: number,
  random: () => number,
): void {
  const needed = state.randomCursor + count - state.randomValues.length;
  if (needed <= 0) {
    return;
  }
  for (let index = 0; index < needed; index += 1) {
    state.randomValues.push(random());
  }
}

function consumeVolatileRandomValues(
  state: RecalcVolatileState,
  count: number,
  random: () => number,
): Float64Array {
  ensureVolatileRandomValues(state, count, random);
  const values = state.randomValues.slice(state.randomCursor, state.randomCursor + count);
  state.randomCursor += count;
  return Float64Array.from(values);
}

export function createEngineRecalcService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | "workbook"
    | "strings"
    | "scheduler"
    | "wasm"
    | "formulas"
    | "ranges"
    | "events"
    | "getLastMetrics"
    | "setLastMetrics"
  >;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly exportSnapshot: () => import("@bilig/protocol").WorkbookSnapshot;
  readonly importSnapshot: (snapshot: import("@bilig/protocol").WorkbookSnapshot) => void;
  readonly beginMutationCollection: () => void;
  readonly markInputChanged: (cellIndex: number, count: number) => number;
  readonly markFormulaChanged: (cellIndex: number, count: number) => number;
  readonly markExplicitChanged: (cellIndex: number, count: number) => number;
  readonly composeMutationRoots: (changedInputCount: number, formulaChangedCount: number) => U32;
  readonly composeEventChanges: (recalculated: U32, explicitChangedCount: number) => U32;
  readonly unionChangedSets: (...sets: Array<readonly number[] | U32>) => U32;
  readonly composeChangedRootsAndOrdered: (
    changedRoots: readonly number[] | U32,
    ordered: U32,
    orderedCount: number,
  ) => U32;
  readonly emptyChangedSet: () => U32;
  readonly ensureRecalcScratchCapacity: (size: number) => void;
  readonly getPendingKernelSync: () => U32;
  readonly getWasmBatch: () => U32;
  readonly getChangedInputBuffer: () => U32;
  readonly now: () => Date;
  readonly random: () => number;
  readonly performanceNow: () => number;
  readonly materializeSpill: (
    cellIndex: number,
    arrayValue: { values: import("@bilig/protocol").CellValue[]; rows: number; cols: number },
  ) => import("../runtime-state.js").SpillMaterialization;
  readonly clearOwnedSpill: (cellIndex: number) => number[];
  readonly evaluateUnsupportedFormula: (cellIndex: number) => number[];
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[];
  readonly getEntityDependents: (entityId: number) => Uint32Array;
}): EngineRecalcService {
  const shouldRefreshPivot = (
    pivot: WorkbookPivotRecord,
    changed: readonly number[] | U32,
  ): boolean => {
    const ownerSheet = args.state.workbook.getSheet(pivot.source.sheetName);
    if (!ownerSheet) {
      return true;
    }
    const ownerStart = parseCellAddress(pivot.source.startAddress, pivot.source.sheetName);
    const ownerEnd = parseCellAddress(pivot.source.endAddress, pivot.source.sheetName);
    for (let index = 0; index < changed.length; index += 1) {
      const cellIndex = changed[index]!;
      const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex];
      if (sheetId === undefined || sheetId !== ownerSheet.id) {
        continue;
      }
      const row = args.state.workbook.cellStore.rows[cellIndex] ?? -1;
      const col = args.state.workbook.cellStore.cols[cellIndex] ?? -1;
      if (
        row >= ownerStart.row &&
        row <= ownerEnd.row &&
        col >= ownerStart.col &&
        col <= ownerEnd.col
      ) {
        return true;
      }
    }
    return false;
  };

  const refreshPivotOutputs = (changed: readonly number[] | U32, forceAll: boolean): U32 => {
    const pivots = args.state.workbook.listPivots();
    if (pivots.length === 0 || (!forceAll && changed.length === 0)) {
      return args.emptyChangedSet();
    }

    const changedCellIndices: number[] = [];
    const changedSeen = new Set<number>();
    for (let index = 0; index < pivots.length; index += 1) {
      const pivot = pivots[index]!;
      if (!forceAll && !shouldRefreshPivot(pivot, changed)) {
        continue;
      }
      const pivotChanges = args.materializePivot(pivot);
      for (let changeIndex = 0; changeIndex < pivotChanges.length; changeIndex += 1) {
        const cellIndex = pivotChanges[changeIndex]!;
        if (changedSeen.has(cellIndex)) {
          continue;
        }
        changedSeen.add(cellIndex);
        changedCellIndices.push(cellIndex);
      }
    }

    return changedCellIndices.length === 0
      ? args.emptyChangedSet()
      : Uint32Array.from(changedCellIndices);
  };

  const recalculate = (
    changedRoots: readonly number[] | U32,
    kernelSyncRoots: readonly number[] | U32 = changedRoots,
  ): U32 => {
    const started = args.performanceNow();
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + 1);
    let pendingKernelSync = args.getPendingKernelSync();
    let wasmBatch = args.getWasmBatch();
    if (args.state.wasm.ready) {
      args.state.wasm.syncStringPool(args.state.strings.exportLayout());
    }

    const allChangedRoots = [...changedRoots];
    const allOrdered: number[] = [];
    let singlePassOrdered: U32 | null = null;
    let singlePassOrderedCount = 0;
    let passRoots = [...changedRoots];
    let passKernelRoots = [...kernelSyncRoots];
    let totalOrderedCount = 0;
    let totalRangeNodeVisits = 0;
    let wasmCount = 0;
    let jsCount = 0;
    let pendingKernelSyncCount = 0;
    const volatileState = createRecalcVolatileState(args.now);

    const flushWasmBatch = (
      batchCount: number,
      hasVolatile: boolean,
      randCount: number,
    ): number => {
      if (batchCount === 0) {
        return 0;
      }
      args.state.wasm.syncFromStore(
        args.state.workbook.cellStore,
        pendingKernelSync.subarray(0, pendingKernelSyncCount),
      );
      pendingKernelSyncCount = 0;
      if (hasVolatile) {
        args.state.wasm.uploadVolatileNowSerial(volatileState.nowSerial);
        args.state.wasm.uploadVolatileRandomValues(
          consumeVolatileRandomValues(volatileState, randCount, args.random),
        );
      }
      const batchIndices = wasmBatch.subarray(0, batchCount);
      args.state.wasm.evalBatch(batchIndices);
      args.state.wasm.syncToStore(args.state.workbook.cellStore, batchIndices, args.state.strings);
      return batchCount;
    };

    while (passRoots.length > 0) {
      const scheduled = args.state.scheduler.collectDirty(
        passRoots,
        { getDependents: (entityId) => args.getEntityDependents(entityId) },
        args.state.workbook.cellStore,
        (cellIndex) => args.state.formulas.has(cellIndex),
        args.state.ranges.size,
      );
      const ordered = scheduled.orderedFormulaCellIndices;
      const orderedCount = scheduled.orderedFormulaCount;
      totalOrderedCount += orderedCount;
      totalRangeNodeVisits += scheduled.rangeNodeVisits;
      if (singlePassOrdered === null && allOrdered.length === 0) {
        singlePassOrdered = ordered;
        singlePassOrderedCount = orderedCount;
      } else {
        if (singlePassOrdered !== null) {
          for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
            const cellIndex = singlePassOrdered[orderedIndex];
            if (cellIndex !== undefined) {
              allOrdered.push(cellIndex);
            }
          }
          singlePassOrdered = null;
          singlePassOrderedCount = 0;
        }
        for (let orderedIndex = 0; orderedIndex < orderedCount; orderedIndex += 1) {
          allOrdered.push(ordered[orderedIndex]!);
        }
      }

      pendingKernelSyncCount = 0;
      for (let index = 0; index < passKernelRoots.length; index += 1) {
        pendingKernelSync[pendingKernelSyncCount] = passKernelRoots[index]!;
        pendingKernelSyncCount += 1;
      }

      let wasmBatchCount = 0;
      let wasmBatchHasVolatile = false;
      let wasmBatchRandCount = 0;
      const spillChangedRoots: number[] = [];
      const spillChangedSeen = new Set<number>();
      const noteSpillChanges = (changedCellIndices: readonly number[]): void => {
        for (let spillIndex = 0; spillIndex < changedCellIndices.length; spillIndex += 1) {
          const changedCellIndex = changedCellIndices[spillIndex]!;
          if (spillChangedSeen.has(changedCellIndex)) {
            continue;
          }
          spillChangedSeen.add(changedCellIndex);
          spillChangedRoots.push(changedCellIndex);
        }
      };
      const queueKernelSync = (cellIndex: number): void => {
        pendingKernelSync[pendingKernelSyncCount] = cellIndex;
        pendingKernelSyncCount += 1;
      };
      const evaluateWasmSpillFormula = (cellIndex: number, formula: RuntimeFormula): number => {
        args.state.wasm.syncFromStore(
          args.state.workbook.cellStore,
          pendingKernelSync.subarray(0, pendingKernelSyncCount),
        );
        pendingKernelSyncCount = 0;
        if (formula.compiled.volatile) {
          args.state.wasm.uploadVolatileNowSerial(volatileState.nowSerial);
          args.state.wasm.uploadVolatileRandomValues(
            consumeVolatileRandomValues(volatileState, formula.compiled.randCallCount, args.random),
          );
        }
        const batchIndices = Uint32Array.of(cellIndex);
        args.state.wasm.evalBatch(batchIndices);
        args.state.wasm.syncToStore(
          args.state.workbook.cellStore,
          batchIndices,
          args.state.strings,
        );
        const spill = args.state.wasm.readSpill(cellIndex, args.state.strings);
        const spillMaterialization = spill
          ? args.materializeSpill(cellIndex, {
              rows: spill.rows,
              cols: spill.cols,
              values: spill.values,
            })
          : {
              changedCellIndices: args.clearOwnedSpill(cellIndex),
              ownerValue: args.state.workbook.cellStore.getValue(cellIndex, (id) =>
                args.state.strings.get(id),
              ),
            };
        args.state.workbook.cellStore.setValue(
          cellIndex,
          spillMaterialization.ownerValue,
          spillMaterialization.ownerValue.tag === ValueTag.String
            ? args.state.strings.intern(spillMaterialization.ownerValue.value)
            : 0,
        );
        queueKernelSync(cellIndex);
        for (
          let spillIndex = 0;
          spillIndex < spillMaterialization.changedCellIndices.length;
          spillIndex += 1
        ) {
          queueKernelSync(spillMaterialization.changedCellIndices[spillIndex]!);
        }
        noteSpillChanges(spillMaterialization.changedCellIndices);
        return 1;
      };

      for (let index = 0; index < orderedCount; index += 1) {
        const cellIndex = ordered[index]!;
        const formula = args.state.formulas.get(cellIndex);
        if (!formula) {
          continue;
        }
        if (((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
          continue;
        }
        if (formula.compiled.mode === FormulaMode.WasmFastPath && args.state.wasm.ready) {
          if (formula.compiled.producesSpill) {
            wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
            wasmBatchCount = 0;
            wasmBatchHasVolatile = false;
            wasmBatchRandCount = 0;
            wasmCount += evaluateWasmSpillFormula(cellIndex, formula);
            continue;
          }
          wasmBatch[wasmBatchCount] = cellIndex;
          wasmBatchCount += 1;
          wasmBatchHasVolatile = wasmBatchHasVolatile || formula.compiled.volatile;
          wasmBatchRandCount += formula.compiled.randCallCount;
          continue;
        }
        wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
        wasmBatchCount = 0;
        wasmBatchHasVolatile = false;
        wasmBatchRandCount = 0;
        jsCount += 1;
        const spillChanges = args.evaluateUnsupportedFormula(cellIndex);
        noteSpillChanges(spillChanges);
        queueKernelSync(cellIndex);
      }

      wasmCount += flushWasmBatch(wasmBatchCount, wasmBatchHasVolatile, wasmBatchRandCount);
      if (pendingKernelSyncCount > 0) {
        args.state.wasm.syncFromStore(
          args.state.workbook.cellStore,
          pendingKernelSync.subarray(0, pendingKernelSyncCount),
        );
      }

      if (spillChangedRoots.length === 0) {
        break;
      }
      if (singlePassOrdered !== null) {
        for (let orderedIndex = 0; orderedIndex < singlePassOrderedCount; orderedIndex += 1) {
          const cellIndex = singlePassOrdered[orderedIndex];
          if (cellIndex !== undefined) {
            allOrdered.push(cellIndex);
          }
        }
        singlePassOrdered = null;
        singlePassOrderedCount = 0;
      }
      allChangedRoots.push(...spillChangedRoots);
      passRoots = spillChangedRoots;
      passKernelRoots = spillChangedRoots;
    }

    const lastMetrics = { ...args.state.getLastMetrics() };
    lastMetrics.dirtyFormulaCount = totalOrderedCount;
    lastMetrics.jsFormulaCount = jsCount;
    lastMetrics.wasmFormulaCount = wasmCount;
    lastMetrics.rangeNodeVisits = totalRangeNodeVisits;
    lastMetrics.recalcMs = args.performanceNow() - started;
    args.state.setLastMetrics(lastMetrics);
    if (singlePassOrdered !== null) {
      return totalOrderedCount === 0 && allChangedRoots.length === 0
        ? args.emptyChangedSet()
        : args.composeChangedRootsAndOrdered(
            allChangedRoots,
            singlePassOrdered,
            singlePassOrderedCount,
          );
    }
    return totalOrderedCount === 0 && allChangedRoots.length === 0
      ? args.emptyChangedSet()
      : args.composeChangedRootsAndOrdered(
          allChangedRoots,
          Uint32Array.from(allOrdered),
          allOrdered.length,
        );
  };

  const reconcilePivotOutputs = (baseChanged: U32, forceAllPivots = false): U32 => {
    let aggregate = baseChanged;
    let pending = baseChanged;
    let forceAll = forceAllPivots;

    for (let iteration = 0; iteration < 4; iteration += 1) {
      const pivotChanged = refreshPivotOutputs(pending, forceAll);
      if (pivotChanged.length === 0) {
        break;
      }
      aggregate =
        aggregate.length === 0 ? pivotChanged : args.unionChangedSets(aggregate, pivotChanged);
      pending = recalculate(pivotChanged, pivotChanged);
      aggregate = pending.length === 0 ? aggregate : args.unionChangedSets(aggregate, pending);
      forceAll = false;
    }

    return aggregate;
  };

  return {
    recalculate(changedRoots, kernelSyncRoots = changedRoots) {
      return Effect.try({
        try: () => recalculate(changedRoots, kernelSyncRoots),
        catch: (cause) =>
          new EngineRecalcError({
            message: "Failed to recalculate workbook state",
            cause,
          }),
      });
    },
    reconcilePivotOutputs(baseChanged, forceAllPivots = false) {
      return Effect.try({
        try: () => reconcilePivotOutputs(baseChanged, forceAllPivots),
        catch: (cause) =>
          new EngineRecalcError({
            message: "Failed to reconcile pivot outputs",
            cause,
          }),
      });
    },
    recalculateNowSync: recalculate,
    reconcilePivotOutputsNow: reconcilePivotOutputs,
    recalculateNow() {
      return Effect.try({
        try: () => {
          args.beginMutationCollection();
          args.state.workbook.setVolatileContext({
            recalcEpoch: args.state.workbook.getVolatileContext().recalcEpoch + 1,
          });
          let formulaChangedCount = 0;
          let explicitChangedCount = 0;
          args.state.formulas.forEach((_formula, cellIndex) => {
            formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount);
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
          });
          const recalculated = reconcilePivotOutputs(
            recalculate(args.composeMutationRoots(0, formulaChangedCount), args.emptyChangedSet()),
            true,
          );
          const changed = args.composeEventChanges(recalculated, explicitChangedCount);
          const lastMetrics = { ...args.state.getLastMetrics() };
          lastMetrics.batchId += 1;
          lastMetrics.changedInputCount = formulaChangedCount;
          args.state.setLastMetrics(lastMetrics);
          args.state.events.emit(
            {
              kind: "batch",
              invalidation: "cells",
              changedCellIndices: changed,
              invalidatedRanges: [],
              invalidatedRows: [],
              invalidatedColumns: [],
              metrics: lastMetrics,
            },
            changed,
            (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex),
          );
          return Array.from(changed);
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: "Failed to recalculate all formulas",
            cause,
          }),
      });
    },
    recalculateDirty(dirtyRegions) {
      return Effect.try({
        try: () => {
          args.beginMutationCollection();
          let changedInputCount = 0;
          let explicitChangedCount = 0;

          for (const region of dirtyRegions) {
            const sheet = args.state.workbook.getSheet(region.sheetName);
            if (!sheet) {
              continue;
            }

            for (let row = region.rowStart; row <= region.rowEnd; row += 1) {
              for (let col = region.colStart; col <= region.colEnd; col += 1) {
                const cellIndex = args.state.workbook.cellKeyToIndex.get(
                  makeCellKey(sheet.id, row, col),
                );
                if (cellIndex !== undefined) {
                  changedInputCount = args.markInputChanged(cellIndex, changedInputCount);
                  explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount);
                }
              }
            }
          }

          const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount);
          const recalculated = reconcilePivotOutputs(
            recalculate(args.composeMutationRoots(changedInputCount, 0), changedInputArray),
            false,
          );
          const changed = args.composeEventChanges(recalculated, explicitChangedCount);
          const lastMetrics = { ...args.state.getLastMetrics() };
          lastMetrics.batchId += 1;
          lastMetrics.changedInputCount = changedInputCount;
          args.state.setLastMetrics(lastMetrics);
          args.state.events.emit(
            {
              kind: "batch",
              invalidation: "cells",
              changedCellIndices: changed,
              invalidatedRanges: [],
              invalidatedRows: [],
              invalidatedColumns: [],
              metrics: lastMetrics,
            },
            changed,
            (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex),
          );
          return Array.from(changed);
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: "Failed to recalculate dirty regions",
            cause,
          }),
      });
    },
    recalculateDifferential() {
      return Effect.try({
        try: () => {
          const originalSnapshot = args.exportSnapshot();
          args.state.formulas.forEach((formula) => {
            formula.compiled.mode = FormulaMode.JsOnly;
          });
          const jsChanged = Effect.runSync(this.recalculateNow());
          const jsResults = jsChanged.map((idx) => args.getCellByIndex(idx));

          args.importSnapshot(originalSnapshot);
          const wasmChanged = Effect.runSync(this.recalculateNow());
          const wasmResults = wasmChanged.map((idx) => args.getCellByIndex(idx));

          const drift: string[] = [];
          const jsMap = new Map(
            jsResults.map((result) => [`${result.sheetName}!${result.address}`, result]),
          );
          const wasmMap = new Map(
            wasmResults.map((result) => [`${result.sheetName}!${result.address}`, result]),
          );

          for (const [addr, jsCell] of jsMap) {
            const wasmCell = wasmMap.get(addr);
            if (!wasmCell) {
              drift.push(`${addr}: Calculated in JS but MISSING in WASM`);
              continue;
            }
            if (JSON.stringify(jsCell.value) !== JSON.stringify(wasmCell.value)) {
              drift.push(
                `${addr}: JS=${JSON.stringify(jsCell.value)} WASM=${JSON.stringify(wasmCell.value)}`,
              );
            }
          }

          for (const addr of wasmMap.keys()) {
            if (!jsMap.has(addr)) {
              drift.push(`${addr}: Calculated in WASM but MISSING in JS`);
            }
          }

          return { js: jsResults, wasm: wasmResults, drift };
        },
        catch: (cause) =>
          new EngineRecalcError({
            message: "Failed to run differential recalculation",
            cause,
          }),
      });
    },
  };
}
