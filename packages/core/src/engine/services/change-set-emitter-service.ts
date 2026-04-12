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
  const readChangedCell = (
    cellIndex: number,
    fallbackSheetName?: string,
  ): EngineChangedCell | null => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex];
    const row = args.state.workbook.cellStore.rows[cellIndex];
    const col = args.state.workbook.cellStore.cols[cellIndex];
    if (sheetId === undefined || row === undefined || col === undefined) {
      return null;
    }
    const sheetName = fallbackSheetName ?? args.state.workbook.getSheetNameById(sheetId);
    if (sheetName === undefined) {
      return null;
    }
    return {
      kind: "cell",
      cellIndex,
      address: { sheet: sheetId, row, col },
      sheetName,
      a1: formatAddress(row, col),
      newValue: args.state.workbook.cellStore.getValue(cellIndex, (id) =>
        args.state.strings.get(id),
      ),
    };
  };

  return {
    captureChangedCells(changedCellIndices) {
      if (changedCellIndices.length === 0) {
        return [];
      }
      if (changedCellIndices.length <= 2) {
        const first = readChangedCell(changedCellIndices[0]!);
        if (!first) {
          return [];
        }
        if (changedCellIndices.length === 1) {
          return [first];
        }
        const secondCellIndex = changedCellIndices[1]!;
        const secondSheetId = args.state.workbook.cellStore.sheetIds[secondCellIndex];
        const second =
          secondSheetId !== undefined && secondSheetId === first.address.sheet
            ? readChangedCell(secondCellIndex, first.sheetName)
            : readChangedCell(secondCellIndex);
        return second ? [first, second] : [first];
      }
      const sheetNames = new Map<number, string>();
      const changes: EngineChangedCell[] = [];
      for (let index = 0; index < changedCellIndices.length; index += 1) {
        const cellIndex = changedCellIndices[index]!;
        const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex];
        if (sheetId === undefined) {
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
        const changedCell = readChangedCell(cellIndex, sheetName);
        if (changedCell) {
          changes.push(changedCell);
        }
      }
      return changes;
    },
  };
}
