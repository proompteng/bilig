import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import {
  COLUMN_RESIZE_HANDLE_THRESHOLD,
  SCROLLBAR_GUTTER,
  type GridMetrics,
  type GridRect,
  getVisibleColumnBounds,
  getVisibleRowBounds,
  resolveColumnAtClientX,
  resolveRowAtClientY,
} from "./gridMetrics.js";
import type { Item, Rectangle } from "./gridTypes.js";

export interface VisibleRegionState {
  range: Rectangle;
  tx: number;
  ty: number;
}

export interface PointerGeometry {
  hostBounds: GridRect;
  cellWidth: number;
  cellHeight: number;
  dataLeft: number;
  dataTop: number;
  dataRight: number;
  dataBottom: number;
}

export interface SelectedCellBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export type HeaderSelection = { kind: "column"; index: number } | { kind: "row"; index: number };

export interface PointerCellResolutionInput {
  clientX: number;
  clientY: number;
  region: VisibleRegionState;
  geometry: PointerGeometry;
  columnWidths: Readonly<Record<number, number>>;
  rowHeights: Readonly<Record<number, number>>;
  gridMetrics: GridMetrics;
  selectedCell: Item;
  selectedCellBounds?: SelectedCellBounds | null;
  selectionRange?: Rectangle | null;
  hasColumnSelection: boolean;
  hasRowSelection: boolean;
}

export function createPointerGeometry(
  hostBounds: GridRect,
  region: VisibleRegionState,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): PointerGeometry {
  const cellWidth = gridMetrics.columnWidth;
  const cellHeight = gridMetrics.rowHeight;
  const dataLeft = hostBounds.left + gridMetrics.rowMarkerWidth;
  const dataTop = hostBounds.top + gridMetrics.headerHeight;
  const visibleColumnBounds = getVisibleColumnBounds(
    region.range,
    dataLeft,
    MAX_COLS,
    columnWidths,
    gridMetrics.columnWidth,
  );
  const dataWidth =
    visibleColumnBounds.length === 0
      ? region.range.width * cellWidth
      : visibleColumnBounds.at(-1)!.right - dataLeft;
  const visibleRowBounds = getVisibleRowBounds(
    region.range,
    dataTop,
    MAX_ROWS,
    rowHeights,
    gridMetrics.rowHeight,
  );
  const dataHeight =
    visibleRowBounds.length === 0
      ? region.range.height * cellHeight
      : visibleRowBounds.at(-1)!.bottom - dataTop;
  return {
    hostBounds,
    cellWidth,
    cellHeight,
    dataLeft,
    dataTop,
    dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + dataWidth),
    dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + dataHeight),
  };
}

export function resolveColumnResizeTarget(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): number | null {
  if (clientY < geometry.hostBounds.top || clientY >= geometry.dataTop) {
    return null;
  }
  for (const column of getVisibleColumnBounds(
    region.range,
    geometry.dataLeft,
    MAX_COLS,
    columnWidths,
    defaultWidth,
  )) {
    if (column.index >= MAX_COLS - 1) {
      continue;
    }
    if (
      clientX >= column.right - COLUMN_RESIZE_HANDLE_THRESHOLD &&
      clientX <= column.right + COLUMN_RESIZE_HANDLE_THRESHOLD
    ) {
      return column.index;
    }
  }
  return null;
}

export function resolveRowResizeTarget(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
): number | null {
  if (clientX < geometry.hostBounds.left || clientX >= geometry.dataLeft) {
    return null;
  }
  for (const row of getVisibleRowBounds(
    region.range,
    geometry.dataTop,
    MAX_ROWS,
    rowHeights,
    defaultHeight,
  )) {
    if (row.index >= MAX_ROWS - 1) {
      continue;
    }
    if (
      clientY >= row.bottom - COLUMN_RESIZE_HANDLE_THRESHOLD &&
      clientY <= row.bottom + COLUMN_RESIZE_HANDLE_THRESHOLD
    ) {
      return row.index;
    }
  }
  return null;
}

export function resolvePointerCell(input: PointerCellResolutionInput): Item | null {
  const {
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    rowHeights,
    gridMetrics,
    selectedCell,
    selectedCellBounds,
    selectionRange,
    hasColumnSelection,
    hasRowSelection,
  } = input;
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry;

  if (
    clientX >= hostBounds.right - SCROLLBAR_GUTTER ||
    clientY >= hostBounds.bottom - SCROLLBAR_GUTTER
  ) {
    return null;
  }

  if (clientX < dataLeft || clientX >= dataRight || clientY < dataTop || clientY >= dataBottom) {
    return null;
  }

  if (
    selectedCellBounds &&
    !hasColumnSelection &&
    !hasRowSelection &&
    selectionRange?.width === 1 &&
    selectionRange?.height === 1 &&
    clientX >= selectedCellBounds.x - 1 &&
    clientX < selectedCellBounds.x + selectedCellBounds.width &&
    clientY >= selectedCellBounds.y - 1 &&
    clientY < selectedCellBounds.y + selectedCellBounds.height
  ) {
    return selectedCell;
  }

  const col = resolveColumnAtClientX(
    clientX,
    region.range,
    dataLeft,
    MAX_COLS,
    columnWidths,
    gridMetrics.columnWidth,
  );
  const row = resolveRowAtClientY(
    clientY,
    region.range,
    dataTop,
    MAX_ROWS,
    rowHeights,
    gridMetrics.rowHeight,
  );
  if (col === null || col < 0 || col >= MAX_COLS || row === null || row < 0 || row >= MAX_ROWS) {
    return null;
  }

  return [col, row];
}

export function resolveHeaderSelection(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): HeaderSelection | null {
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry;
  const headerBottom = dataTop;
  const rowAreaRight = dataLeft;

  if (
    clientY >= hostBounds.top &&
    clientY < headerBottom &&
    clientX >= dataLeft &&
    clientX < dataRight
  ) {
    const col = resolveColumnAtClientX(
      clientX,
      region.range,
      dataLeft,
      MAX_COLS,
      columnWidths,
      gridMetrics.columnWidth,
    );
    if (col !== null && col >= 0 && col < MAX_COLS) {
      return { kind: "column", index: col };
    }
  }

  if (
    clientX >= hostBounds.left &&
    clientX < rowAreaRight &&
    clientY >= dataTop &&
    clientY < dataBottom
  ) {
    const row = resolveRowAtClientY(
      clientY,
      region.range,
      dataTop,
      MAX_ROWS,
      rowHeights,
      gridMetrics.rowHeight,
    );
    if (row !== null && row >= 0 && row < MAX_ROWS) {
      return { kind: "row", index: row };
    }
  }

  return null;
}

export function resolveHeaderSelectionForDrag(
  kind: HeaderSelection["kind"],
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): HeaderSelection | null {
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry;

  if (kind === "column") {
    if (
      clientX < dataLeft ||
      clientX >= dataRight ||
      clientY < hostBounds.top ||
      clientY >= dataBottom
    ) {
      return null;
    }
    const col = resolveColumnAtClientX(
      clientX,
      region.range,
      dataLeft,
      MAX_COLS,
      columnWidths,
      gridMetrics.columnWidth,
    );
    if (col === null || col < 0 || col >= MAX_COLS) {
      return null;
    }
    return { kind: "column", index: col };
  }

  if (
    clientY < dataTop ||
    clientY >= dataBottom ||
    clientX < hostBounds.left ||
    clientX >= dataRight
  ) {
    return null;
  }
  const row = resolveRowAtClientY(
    clientY,
    region.range,
    dataTop,
    MAX_ROWS,
    rowHeights,
    gridMetrics.rowHeight,
  );
  if (row === null || row < 0 || row >= MAX_ROWS) {
    return null;
  }
  return { kind: "row", index: row };
}
