export interface GridMetrics {
  columnWidth: number;
  rowHeight: number;
  headerHeight: number;
  rowMarkerWidth: number;
}

export interface GridRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface VisibleColumnBound {
  index: number;
  left: number;
  right: number;
  width: number;
}

export const PRODUCT_COLUMN_WIDTH = 104;
export const PRODUCT_ROW_HEIGHT = 22;
export const PRODUCT_HEADER_HEIGHT = 24;
export const PRODUCT_ROW_MARKER_WIDTH = 46;
export const SCROLLBAR_GUTTER = 18;
export const COLUMN_RESIZE_HANDLE_THRESHOLD = 6;
export const MIN_COLUMN_WIDTH = 44;
export const MAX_COLUMN_WIDTH = 480;
export const EMPTY_COLUMN_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});

export function getGridMetrics(): GridMetrics {
  return {
    columnWidth: PRODUCT_COLUMN_WIDTH,
    rowHeight: PRODUCT_ROW_HEIGHT,
    headerHeight: PRODUCT_HEADER_HEIGHT,
    rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH,
  };
}

export function getResolvedColumnWidth(
  columnWidths: Readonly<Record<number, number>>,
  col: number,
  defaultWidth: number,
): number {
  return columnWidths[col] ?? defaultWidth;
}

export function getVisibleColumnBounds(
  range: { x: number; width: number },
  dataLeft: number,
  maxCols: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): VisibleColumnBound[] {
  const bounds: VisibleColumnBound[] = [];
  const colEnd = Math.min(maxCols - 1, range.x + range.width - 1);
  let cursor = dataLeft;
  for (let col = range.x; col <= colEnd; col += 1) {
    const width = getResolvedColumnWidth(columnWidths, col, defaultWidth);
    bounds.push({ index: col, left: cursor, right: cursor + width, width });
    cursor += width;
  }
  return bounds;
}

export function resolveColumnAtClientX(
  clientX: number,
  range: { x: number; width: number },
  dataLeft: number,
  maxCols: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): number | null {
  for (const column of getVisibleColumnBounds(
    range,
    dataLeft,
    maxCols,
    columnWidths,
    defaultWidth,
  )) {
    if (clientX >= column.left && clientX < column.right) {
      return column.index;
    }
  }
  return null;
}
