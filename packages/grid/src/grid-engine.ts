import type { CellSnapshot, CellStyleRecord } from "@bilig/protocol";

export interface GridSheetLike {
  grid: {
    forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void;
  };
}

export interface GridWorkbookLike {
  getSheet(sheetName: string): GridSheetLike | undefined;
}

export interface GridEngineLike {
  getCell(sheetName: string, address: string): CellSnapshot;
  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined;
  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void;
  workbook: GridWorkbookLike;
}
