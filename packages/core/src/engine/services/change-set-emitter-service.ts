import { formatAddress } from "@bilig/formula";
import type { EngineChangedCell } from "@bilig/protocol";
import type { EngineRuntimeState } from "../runtime-state.js";

export interface EngineChangeSetEmitterService {
  readonly captureChangedCells: (
    changedCellIndices: readonly number[] | Uint32Array,
  ) => readonly EngineChangedCell[];
}

export function createEngineChangeSetEmitterService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings">;
}): EngineChangeSetEmitterService {
  return {
    captureChangedCells(changedCellIndices) {
      if (changedCellIndices.length === 0) {
        return [];
      }
      const sheetNames = new Map<number, string>();
      const changes: EngineChangedCell[] = [];
      for (let index = 0; index < changedCellIndices.length; index += 1) {
        const cellIndex = changedCellIndices[index]!;
        const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex];
        const row = args.state.workbook.cellStore.rows[cellIndex];
        const col = args.state.workbook.cellStore.cols[cellIndex];
        if (sheetId === undefined || row === undefined || col === undefined) {
          continue;
        }
        let sheetName = sheetNames.get(sheetId);
        if (sheetName === undefined) {
          sheetName = args.state.workbook.getSheetNameById(sheetId);
          if (sheetName === undefined) {
            continue;
          }
          sheetNames.set(sheetId, sheetName);
        }
        changes.push({
          kind: "cell",
          cellIndex,
          address: { sheet: sheetId, row, col },
          sheetName,
          a1: formatAddress(row, col),
          newValue: args.state.workbook.cellStore.getValue(cellIndex, (id) =>
            args.state.strings.get(id),
          ),
        });
      }
      return changes;
    },
  };
}
