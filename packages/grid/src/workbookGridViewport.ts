import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import {
  getGridMetrics,
  getResolvedColumnWidth,
  getResolvedRowHeight,
  resolveRowOffset,
} from "./gridMetrics.js";
import type { Item } from "./gridTypes.js";
import type { VisibleRegionState } from "./gridPointer.js";

export function resolveVisibleRegionFromScroll(options: {
  scrollLeft: number;
  scrollTop: number;
  viewportWidth: number;
  viewportHeight: number;
  columnWidths: Readonly<Record<number, number>>;
  rowHeights: Readonly<Record<number, number>>;
  gridMetrics: ReturnType<typeof getGridMetrics>;
}): VisibleRegionState {
  const {
    scrollLeft,
    scrollTop,
    viewportWidth,
    viewportHeight,
    columnWidths,
    rowHeights,
    gridMetrics,
  } = options;
  const bodyWidth = Math.max(0, viewportWidth - gridMetrics.rowMarkerWidth);
  const bodyHeight = Math.max(0, viewportHeight - gridMetrics.headerHeight);
  const horizontalAnchor = resolveColumnAnchor(scrollLeft, columnWidths, gridMetrics.columnWidth);
  const verticalAnchor = resolveRowAnchor(scrollTop, rowHeights, gridMetrics.rowHeight);

  return {
    range: {
      x: horizontalAnchor.index,
      y: verticalAnchor.index,
      width: resolveVisibleColumnCount({
        startCol: horizontalAnchor.index,
        tx: horizontalAnchor.offset,
        bodyWidth,
        columnWidths,
        defaultWidth: gridMetrics.columnWidth,
      }),
      height: resolveVisibleRowCount({
        startRow: verticalAnchor.index,
        ty: verticalAnchor.offset,
        bodyHeight,
        rowHeights,
        defaultHeight: gridMetrics.rowHeight,
      }),
    },
    tx: horizontalAnchor.offset,
    ty: verticalAnchor.offset,
  };
}

export function resolveColumnOffset(
  targetColumn: number,
  sortedColumnWidthOverrides: readonly (readonly [number, number])[],
  defaultWidth: number,
): number {
  let offset = targetColumn * defaultWidth;
  for (const [columnIndex, width] of sortedColumnWidthOverrides) {
    if (columnIndex >= targetColumn) {
      break;
    }
    offset += width - defaultWidth;
  }
  return offset;
}

export function scrollCellIntoView(options: {
  cell: Item;
  columnWidths: Readonly<Record<number, number>>;
  rowHeights: Readonly<Record<number, number>>;
  gridMetrics: ReturnType<typeof getGridMetrics>;
  scrollViewport: HTMLDivElement;
  sortedColumnWidthOverrides: readonly (readonly [number, number])[];
  sortedRowHeightOverrides: readonly (readonly [number, number])[];
}): void {
  const {
    cell,
    columnWidths,
    rowHeights,
    gridMetrics,
    scrollViewport,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  } = options;
  const cellLeft = resolveColumnOffset(
    cell[0],
    sortedColumnWidthOverrides,
    gridMetrics.columnWidth,
  );
  const cellWidth = getResolvedColumnWidth(columnWidths, cell[0], gridMetrics.columnWidth);
  const bodyWidth = Math.max(0, scrollViewport.clientWidth - gridMetrics.rowMarkerWidth);
  if (cellLeft < scrollViewport.scrollLeft) {
    scrollViewport.scrollLeft = cellLeft;
  } else if (cellLeft + cellWidth > scrollViewport.scrollLeft + bodyWidth) {
    scrollViewport.scrollLeft = cellLeft + cellWidth - bodyWidth;
  }

  const cellTop = resolveRowOffset(cell[1], sortedRowHeightOverrides, gridMetrics.rowHeight);
  const cellHeight = getResolvedRowHeight(rowHeights, cell[1], gridMetrics.rowHeight);
  const bodyHeight = Math.max(0, scrollViewport.clientHeight - gridMetrics.headerHeight);
  if (cellTop < scrollViewport.scrollTop) {
    scrollViewport.scrollTop = cellTop;
  } else if (cellTop + cellHeight > scrollViewport.scrollTop + bodyHeight) {
    scrollViewport.scrollTop = cellTop + cellHeight - bodyHeight;
  }
}

export function resolveViewportScrollPosition(options: {
  viewport: Pick<VisibleRegionState["range"], "x" | "y"> | { colStart: number; rowStart: number };
  sortedColumnWidthOverrides: readonly (readonly [number, number])[];
  sortedRowHeightOverrides: readonly (readonly [number, number])[];
  gridMetrics: ReturnType<typeof getGridMetrics>;
}): { scrollLeft: number; scrollTop: number } {
  const colStart = "colStart" in options.viewport ? options.viewport.colStart : options.viewport.x;
  const rowStart = "rowStart" in options.viewport ? options.viewport.rowStart : options.viewport.y;
  return {
    scrollLeft: resolveColumnOffset(
      colStart,
      options.sortedColumnWidthOverrides,
      options.gridMetrics.columnWidth,
    ),
    scrollTop: resolveRowOffset(
      rowStart,
      options.sortedRowHeightOverrides,
      options.gridMetrics.rowHeight,
    ),
  };
}

export function hasSelectionTargetChanged(
  previousSelection: { sheetName: string; col: number; row: number } | null,
  nextSelection: { sheetName: string; col: number; row: number },
): boolean {
  return (
    previousSelection === null ||
    previousSelection.sheetName !== nextSelection.sheetName ||
    previousSelection.col !== nextSelection.col ||
    previousSelection.row !== nextSelection.row
  );
}

function resolveVisibleColumnCount(options: {
  startCol: number;
  tx: number;
  bodyWidth: number;
  columnWidths: Readonly<Record<number, number>>;
  defaultWidth: number;
}): number {
  const { startCol, tx, bodyWidth, columnWidths, defaultWidth } = options;
  let coveredWidth = -tx;
  let count = 0;
  for (let col = startCol; col < MAX_COLS && coveredWidth < bodyWidth; col += 1) {
    coveredWidth += getResolvedColumnWidth(columnWidths, col, defaultWidth);
    count += 1;
  }
  return Math.max(1, count + 1);
}

function resolveVisibleRowCount(options: {
  startRow: number;
  ty: number;
  bodyHeight: number;
  rowHeights: Readonly<Record<number, number>>;
  defaultHeight: number;
}): number {
  const { startRow, ty, bodyHeight, rowHeights, defaultHeight } = options;
  let coveredHeight = -ty;
  let count = 0;
  for (let row = startRow; row < MAX_ROWS && coveredHeight < bodyHeight; row += 1) {
    coveredHeight += getResolvedRowHeight(rowHeights, row, defaultHeight);
    count += 1;
  }
  return Math.max(1, count + 1);
}

function resolveColumnAnchor(
  scrollLeft: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): { index: number; offset: number } {
  let consumed = 0;
  for (let col = 0; col < MAX_COLS; col += 1) {
    const width = getResolvedColumnWidth(columnWidths, col, defaultWidth);
    if (consumed + width > scrollLeft) {
      return { index: col, offset: scrollLeft - consumed };
    }
    consumed += width;
  }
  return { index: MAX_COLS - 1, offset: 0 };
}

function resolveRowAnchor(
  scrollTop: number,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
): { index: number; offset: number } {
  let consumed = 0;
  for (let row = 0; row < MAX_ROWS; row += 1) {
    const height = getResolvedRowHeight(rowHeights, row, defaultHeight);
    if (consumed + height > scrollTop) {
      return { index: row, offset: scrollTop - consumed };
    }
    consumed += height;
  }
  return { index: MAX_ROWS - 1, offset: 0 };
}
