import { Effect } from "effect";
import type { EdgeArena, EdgeSlice } from "../../edge-arena.js";
import { entityPayload, isRangeEntity } from "../../entity-ids.js";
import { growUint32 } from "../../engine-buffer-utils.js";
import type { EngineRuntimeState, U32 } from "../runtime-state.js";
import { EngineTraversalError } from "../errors.js";

export interface EngineTraversalService {
  readonly getEntityDependents: (
    entityId: number,
  ) => Effect.Effect<Uint32Array, EngineTraversalError>;
  readonly collectFormulaDependents: (
    entityId: number,
  ) => Effect.Effect<Uint32Array, EngineTraversalError>;
  readonly forEachFormulaDependencyCell: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>;
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => Effect.Effect<void, EngineTraversalError>;
  readonly getEntityDependentsNow: (entityId: number) => Uint32Array;
  readonly collectFormulaDependentsNow: (entityId: number) => Uint32Array;
  readonly forEachFormulaDependencyCellNow: (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ) => void;
  readonly forEachSheetCellNow: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => void;
}

function traversalErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

export function createEngineTraversalService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "formulas" | "ranges">;
  readonly edgeArena: EdgeArena;
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>;
    reverseRangeEdges: Array<EdgeSlice | undefined>;
  };
}): EngineTraversalService {
  let topoFormulaBuffer: U32 = new Uint32Array(128);
  let topoEntityQueue: U32 = new Uint32Array(128);
  let topoFormulaSeenEpoch = 1;
  let topoRangeSeenEpoch = 1;
  let topoFormulaSeen: U32 = new Uint32Array(128);
  let topoRangeSeen: U32 = new Uint32Array(128);

  const ensureTraversalScratchCapacity = (
    cellSize: number,
    entitySize: number,
    rangeSize: number,
  ): void => {
    if (cellSize > topoFormulaBuffer.length) {
      topoFormulaBuffer = growUint32(topoFormulaBuffer, cellSize);
    }
    if (cellSize > topoFormulaSeen.length) {
      topoFormulaSeen = growUint32(topoFormulaSeen, cellSize);
    }
    if (entitySize > topoEntityQueue.length) {
      topoEntityQueue = growUint32(topoEntityQueue, entitySize);
    }
    if (rangeSize > topoRangeSeen.length) {
      topoRangeSeen = growUint32(topoRangeSeen, rangeSize);
    }
  };

  const getReverseEdgeSlice = (entityId: number): EdgeSlice | undefined => {
    if (isRangeEntity(entityId)) {
      return args.reverseState.reverseRangeEdges[entityPayload(entityId)];
    }
    return args.reverseState.reverseCellEdges[entityPayload(entityId)];
  };

  const getEntityDependentsNow = (entityId: number): Uint32Array =>
    args.edgeArena.readView(getReverseEdgeSlice(entityId) ?? args.edgeArena.empty());

  const forEachFormulaDependencyCellNow = (
    cellIndex: number,
    fn: (dependencyCellIndex: number) => void,
  ): void => {
    const formula = args.state.formulas.get(cellIndex);
    if (!formula) {
      return;
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      fn(formula.dependencyIndices[index]!);
    }
  };

  const forEachSheetCellNow = (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ): void => {
    const sheet = args.state.workbook.getSheetById(sheetId);
    if (!sheet) {
      return;
    }
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      fn(cellIndex, row, col);
    });
  };

  const collectFormulaDependentsNow = (entityId: number): Uint32Array => {
    ensureTraversalScratchCapacity(
      Math.max(args.state.workbook.cellStore.size + 1, 1),
      Math.max(args.state.workbook.cellStore.size + args.state.ranges.size + 1, 1),
      Math.max(args.state.ranges.size + 1, 1),
    );

    topoFormulaSeenEpoch += 1;
    if (topoFormulaSeenEpoch === 0xffff_ffff) {
      topoFormulaSeenEpoch = 1;
      topoFormulaSeen.fill(0);
    }
    topoRangeSeenEpoch += 1;
    if (topoRangeSeenEpoch === 0xffff_ffff) {
      topoRangeSeenEpoch = 1;
      topoRangeSeen.fill(0);
    }

    let entityQueueLength = 1;
    let formulaCount = 0;
    topoEntityQueue[0] = entityId;

    for (let queueIndex = 0; queueIndex < entityQueueLength; queueIndex += 1) {
      const currentEntity = topoEntityQueue[queueIndex]!;
      const dependents = getEntityDependentsNow(currentEntity);
      for (let index = 0; index < dependents.length; index += 1) {
        const dependent = dependents[index]!;
        if (isRangeEntity(dependent)) {
          const rangeIndex = entityPayload(dependent);
          if (topoRangeSeen[rangeIndex] === topoRangeSeenEpoch) {
            continue;
          }
          topoRangeSeen[rangeIndex] = topoRangeSeenEpoch;
          topoEntityQueue[entityQueueLength] = dependent;
          entityQueueLength += 1;
          continue;
        }

        const formulaCellIndex = entityPayload(dependent);
        if (topoFormulaSeen[formulaCellIndex] === topoFormulaSeenEpoch) {
          continue;
        }
        topoFormulaSeen[formulaCellIndex] = topoFormulaSeenEpoch;
        topoFormulaBuffer[formulaCount] = formulaCellIndex;
        formulaCount += 1;
      }
    }

    return topoFormulaBuffer.subarray(0, formulaCount);
  };

  return {
    getEntityDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(getEntityDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage("Failed to read entity dependents", cause),
            cause,
          }),
      });
    },
    collectFormulaDependents(entityId) {
      return Effect.try({
        try: () => Uint32Array.from(collectFormulaDependentsNow(entityId)),
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage("Failed to collect formula dependents", cause),
            cause,
          }),
      });
    },
    forEachFormulaDependencyCell(cellIndex, fn) {
      return Effect.try({
        try: () => {
          forEachFormulaDependencyCellNow(cellIndex, fn);
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage("Failed to iterate formula dependencies", cause),
            cause,
          }),
      });
    },
    forEachSheetCell(sheetId, fn) {
      return Effect.try({
        try: () => {
          forEachSheetCellNow(sheetId, fn);
        },
        catch: (cause) =>
          new EngineTraversalError({
            message: traversalErrorMessage("Failed to iterate sheet cells", cause),
            cause,
          }),
      });
    },
    getEntityDependentsNow,
    collectFormulaDependentsNow,
    forEachFormulaDependencyCellNow,
    forEachSheetCellNow,
  };
}
