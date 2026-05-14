import type { CellSnapshot, CellStyleRecord, WorkbookMergeRangeSnapshot } from '@bilig/protocol'

export type GridEngineSheetChannel = 'merges'

export interface GridSheetLike {
  grid: {
    forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void
  }
}

export interface GridWorkbookLike {
  getSheet(sheetName: string): GridSheetLike | undefined
}

export interface GridRenderRevisionSnapshot {
  readonly authoritativeRevision: number | null
  readonly projectedRevision: number
  readonly tileSceneCameraSeq: number | null
  readonly tileSceneRevision: number | null
}

export interface GridEngineLike {
  getCell(sheetName: string, address: string): CellSnapshot
  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined
  getMergeRange?(sheetName: string, address: string): WorkbookMergeRangeSnapshot | undefined
  getRenderRevisionSnapshot?(): GridRenderRevisionSnapshot
  listMergeRanges?(sheetName: string): readonly WorkbookMergeRangeSnapshot[]
  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void
  subscribeSheetChannel?(sheetName: string, channel: GridEngineSheetChannel, listener: () => void): () => void
  workbook: GridWorkbookLike
}
