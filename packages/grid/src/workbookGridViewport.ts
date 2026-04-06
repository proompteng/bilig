import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import { getGridMetrics, getResolvedColumnWidth } from "./gridMetrics.js";
import type { Item } from "./gridTypes.js";
import type { VisibleRegionState } from "./gridPointer.js";

export function resolveVisibleRegionFromScroll(options: {
  scrollLeft: number;
  scrollTop: number;
  viewportWidth: number;
  viewportHeight: number;
  columnWidths: Readonly<Record<number, number>>;
  gridMetrics: ReturnType<typeof getGridMetrics>;
}): VisibleRegionState {
  const { scrollLeft, scrollTop, viewportWidth, viewportHeight, columnWidths, gridMetrics } =
    options;
  const bodyWidth = Math.max(0, viewportWidth - gridMetrics.rowMarkerWidth);
  const bodyHeight = Math.max(0, viewportHeight - gridMetrics.headerHeight);
  const horizontalAnchor = resolveColumnAnchor(scrollLeft, columnWidths, gridMetrics.columnWidth);
  const startRow = Math.min(
    MAX_ROWS - 1,
    Math.max(0, Math.floor(scrollTop / gridMetrics.rowHeight)),
  );
  const ty = scrollTop - startRow * gridMetrics.rowHeight;

  return {
    range: {
      x: horizontalAnchor.index,
      y: startRow,
      width: resolveVisibleColumnCount({
        startCol: horizontalAnchor.index,
        tx: horizontalAnchor.offset,
        bodyWidth,
        columnWidths,
        defaultWidth: gridMetrics.columnWidth,
      }),
      height: Math.max(1, Math.ceil((bodyHeight + ty) / gridMetrics.rowHeight) + 1),
    },
    tx: horizontalAnchor.offset,
    ty,
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
  gridMetrics: ReturnType<typeof getGridMetrics>;
  scrollViewport: HTMLDivElement;
  sortedColumnWidthOverrides: readonly (readonly [number, number])[];
}): void {
  const { cell, columnWidths, gridMetrics, scrollViewport, sortedColumnWidthOverrides } = options;
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

  const cellTop = cell[1] * gridMetrics.rowHeight;
  const bodyHeight = Math.max(0, scrollViewport.clientHeight - gridMetrics.headerHeight);
  if (cellTop < scrollViewport.scrollTop) {
    scrollViewport.scrollTop = cellTop;
  } else if (cellTop + gridMetrics.rowHeight > scrollViewport.scrollTop + bodyHeight) {
    scrollViewport.scrollTop = cellTop + gridMetrics.rowHeight - bodyHeight;
  }
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
