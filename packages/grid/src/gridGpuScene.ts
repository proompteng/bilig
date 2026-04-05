import { formatAddress } from "@bilig/formula";
import { ValueTag, type CellStyleRecord } from "@bilig/protocol";
import type { GridEngineLike } from "./grid-engine.js";
import type { GridMetrics } from "./gridMetrics.js";
import { getVisibleColumnBounds } from "./gridMetrics.js";
import type { HeaderSelection } from "./gridPointer.js";
import type { GridSelection, Item, Rectangle } from "./gridTypes.js";

export interface GridGpuColor {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}

export interface GridGpuRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly color: GridGpuColor;
}

export interface GridGpuScene {
  readonly fillRects: readonly GridGpuRect[];
  readonly borderRects: readonly GridGpuRect[];
}

interface BuildGridGpuSceneOptions {
  readonly engine: GridEngineLike;
  readonly sheetName: string;
  readonly visibleItems: readonly Item[];
  readonly visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
  readonly gridMetrics: GridMetrics;
  readonly columnWidths: Readonly<Record<number, number>>;
  readonly hostBounds: Pick<DOMRect, "left" | "top">;
  readonly getCellBounds: (col: number, row: number) => Rectangle | undefined;
  readonly gridSelection: GridSelection;
  readonly selectedCell: Item;
  readonly fillPreviewRange?: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  readonly selectionRange?: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  readonly hoveredCell?: Item | null;
  readonly hoveredHeader?: HeaderSelection | null;
  readonly resizeGuideColumn?: number | null;
  readonly activeHeaderDrag?: HeaderSelection | null;
}

const FALLBACK_COLOR: GridGpuColor = Object.freeze({ r: 0, g: 0, b: 0, a: 1 });
const GRID_LINE_COLOR = parseGpuColor("#e3e9f0");
const HEADER_FILL_COLOR = parseGpuColor("#f8f9fa");
const HEADER_SELECTED_FILL_COLOR = parseGpuColor("#e6f4ea");
const HEADER_HOVER_FILL_COLOR = parseGpuColor("#f1f3f4");
const HEADER_DRAG_ANCHOR_FILL_COLOR = parseGpuColor("#d7eadf");
const SELECTION_FILL_COLOR = parseGpuColor("rgba(31, 122, 67, 0.06)");
const SELECTION_OUTLINE_COLOR = parseGpuColor("#1f7a43");
const HOVER_FILL_COLOR = parseGpuColor("rgba(95, 99, 104, 0.06)");
const HOVER_OUTLINE_COLOR = parseGpuColor("rgba(95, 99, 104, 0.45)");
const RESIZE_GUIDE_COLOR = parseGpuColor("#1f7a43");
const RESIZE_GUIDE_GLOW_COLOR = parseGpuColor("rgba(31, 122, 67, 0.18)");
const CHECKBOX_BORDER_COLOR = parseGpuColor("#5f6368");
const CHECKBOX_SURFACE_COLOR = parseGpuColor("#ffffff");
const CHECKBOX_SELECTED_COLOR = parseGpuColor("#1f7a43");
const CHECKBOX_CHECK_COLOR = parseGpuColor("#ffffff");

export function buildGridGpuScene({
  engine,
  sheetName,
  visibleItems,
  visibleRegion,
  gridMetrics,
  columnWidths,
  hostBounds,
  getCellBounds,
  gridSelection,
  selectedCell,
  fillPreviewRange = null,
  selectionRange = null,
  hoveredCell = null,
  hoveredHeader = null,
  resizeGuideColumn = null,
  activeHeaderDrag = null,
}: BuildGridGpuSceneOptions): GridGpuScene {
  const fillRects: GridGpuRect[] = [];
  const borderRects: GridGpuRect[] = [];
  pushHeaderRects({
    borderRects,
    columnWidths,
    fillRects,
    gridMetrics,
    gridSelection,
    activeHeaderDrag,
    hoveredHeader,
    resizeGuideColumn,
    selectedCell,
    selectionRange,
    visibleRegion,
  });
  if (visibleItems.length === 0) {
    return {
      fillRects,
      borderRects,
    };
  }

  const visibleCols = visibleItems.map(([col]) => col);
  const visibleRows = visibleItems.map(([, row]) => row);
  const visibleMinCol = Math.min(...visibleCols);
  const visibleMaxCol = Math.max(...visibleCols);
  const visibleMinRow = Math.min(...visibleRows);
  const visibleMaxRow = Math.max(...visibleRows);

  for (const [col, row] of visibleItems) {
    const bounds = getCellBounds(col, row);
    if (!bounds) {
      continue;
    }
    const rect = {
      x: bounds.x - hostBounds.left,
      y: bounds.y - hostBounds.top,
      width: bounds.width,
      height: bounds.height,
    };

    const snapshot = engine.getCell(sheetName, formatAddress(row, col));
    const style = engine.getCellStyle(snapshot.styleId);

    if (style?.fill?.backgroundColor) {
      fillRects.push({
        x: rect.x + 1,
        y: rect.y + 1,
        width: Math.max(0, rect.width - 2),
        height: Math.max(0, rect.height - 2),
        color: parseGpuColor(style.fill.backgroundColor),
      });
    }

    pushGridLineRects(borderRects, rect, row, col, visibleMinRow, visibleMinCol);

    if (snapshot.value.tag === ValueTag.Boolean) {
      pushBooleanCellRects(fillRects, borderRects, rect, snapshot.value.value);
    }

    if (!style?.borders) {
      continue;
    }

    const borderEntries = [
      ["top", style.borders.top],
      ["right", style.borders.right],
      ["bottom", style.borders.bottom],
      ["left", style.borders.left],
    ] as const;

    for (const [side, border] of borderEntries) {
      if (!border) {
        continue;
      }
      borderRects.push(...createBorderRects(rect, side, border));
    }
  }

  if (selectionRange) {
    pushSelectionRects({
      allowHandle: gridSelection.columns.length === 0 && gridSelection.rows.length === 0,
      borderRects,
      fillRects,
      getCellBounds,
      hostBounds,
      selectionRange,
      visibleMaxCol,
      visibleMaxRow,
      visibleMinCol,
      visibleMinRow,
    });
  }

  if (hoveredCell) {
    pushHoveredCellRects({
      borderRects,
      fillRects,
      getCellBounds,
      hostBounds,
      hoveredCell,
      selectionRange,
      gridSelection,
    });
  }

  if (fillPreviewRange) {
    pushSelectionRects({
      allowHandle: true,
      borderRects,
      fillColor: parseGpuColor("rgba(31, 122, 67, 0.04)"),
      fillRects,
      getCellBounds,
      hostBounds,
      outlineColor: SELECTION_OUTLINE_COLOR,
      selectionRange: fillPreviewRange,
      visibleMaxCol,
      visibleMaxRow,
      visibleMinCol,
      visibleMinRow,
    });
  }

  return {
    fillRects,
    borderRects,
  };
}

function pushHeaderRects(options: {
  borderRects: GridGpuRect[];
  columnWidths: Readonly<Record<number, number>>;
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  gridSelection: GridSelection;
  activeHeaderDrag: HeaderSelection | null;
  hoveredHeader: HeaderSelection | null;
  resizeGuideColumn: number | null;
  selectedCell: Item;
  selectionRange: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
}) {
  const {
    borderRects,
    columnWidths,
    fillRects,
    gridMetrics,
    gridSelection,
    activeHeaderDrag,
    hoveredHeader,
    resizeGuideColumn,
    selectedCell,
    selectionRange,
    visibleRegion,
  } = options;

  const visibleColumns = getVisibleColumnBounds(
    visibleRegion.range,
    gridMetrics.rowMarkerWidth - visibleRegion.tx,
    Number.MAX_SAFE_INTEGER,
    columnWidths,
    gridMetrics.columnWidth,
  );
  const visibleRowEnd = visibleRegion.range.y + visibleRegion.range.height - 1;
  const selectedColumns = resolveAxisSelectionRange(
    selectionRange?.x ?? selectedCell[0],
    selectionRange ? selectionRange.x + selectionRange.width - 1 : selectedCell[0],
    gridSelection.columns,
  );
  const selectedRows = resolveAxisSelectionRange(
    selectionRange?.y ?? selectedCell[1],
    selectionRange ? selectionRange.y + selectionRange.height - 1 : selectedCell[1],
    gridSelection.rows,
  );

  if (gridSelection.columns.length > 0) {
    pushColumnSelectionBodyRects({
      borderRects,
      fillRects,
      gridMetrics,
      selectedColumns,
      visibleColumns,
      visibleRegion,
    });
  }

  if (gridSelection.rows.length > 0) {
    pushRowSelectionBodyRects({
      borderRects,
      fillRects,
      gridMetrics,
      selectedRows,
      visibleRegion,
      visibleRowEnd,
      visibleWidth:
        visibleColumns.length === 0
          ? 0
          : visibleColumns.at(-1)!.right - gridMetrics.rowMarkerWidth + visibleRegion.tx,
    });
  }

  if (activeHeaderDrag?.kind === "column") {
    pushColumnHeaderDragGuideRects({
      activeHeaderDrag,
      borderRects,
      fillRects,
      gridMetrics,
      selectedColumns,
      visibleColumns,
      visibleRegion,
    });
  }

  if (activeHeaderDrag?.kind === "row") {
    pushRowHeaderDragGuideRects({
      activeHeaderDrag,
      borderRects,
      fillRects,
      gridMetrics,
      selectedRows,
      visibleRegion,
      visibleRowEnd,
      visibleWidth:
        visibleColumns.length === 0
          ? 0
          : visibleColumns.at(-1)!.right - gridMetrics.rowMarkerWidth + visibleRegion.tx,
    });
  }

  fillRects.push({
    x: 0,
    y: 0,
    width: gridMetrics.rowMarkerWidth,
    height: gridMetrics.headerHeight,
    color: HEADER_FILL_COLOR,
  });
  borderRects.push(
    {
      x: 0,
      y: gridMetrics.headerHeight - 1,
      width: gridMetrics.rowMarkerWidth,
      height: 1,
      color: GRID_LINE_COLOR,
    },
    {
      x: gridMetrics.rowMarkerWidth - 1,
      y: 0,
      width: 1,
      height: gridMetrics.headerHeight,
      color: GRID_LINE_COLOR,
    },
  );

  for (const column of visibleColumns) {
    fillRects.push({
      x: column.left,
      y: 0,
      width: column.width,
      height: gridMetrics.headerHeight,
      color:
        column.index >= selectedColumns.start && column.index <= selectedColumns.end
          ? activeHeaderDrag?.kind === "column" && activeHeaderDrag.index === column.index
            ? HEADER_DRAG_ANCHOR_FILL_COLOR
            : HEADER_SELECTED_FILL_COLOR
          : hoveredHeader?.kind === "column" && hoveredHeader.index === column.index
            ? HEADER_HOVER_FILL_COLOR
            : HEADER_FILL_COLOR,
    });
    borderRects.push(
      {
        x: column.left + column.width - 1,
        y: 0,
        width: 1,
        height: gridMetrics.headerHeight,
        color: GRID_LINE_COLOR,
      },
      {
        x: column.left,
        y: gridMetrics.headerHeight - 1,
        width: column.width,
        height: 1,
        color: GRID_LINE_COLOR,
      },
    );
  }

  for (let row = visibleRegion.range.y; row <= visibleRowEnd; row += 1) {
    const top =
      gridMetrics.headerHeight +
      (row - visibleRegion.range.y) * gridMetrics.rowHeight -
      visibleRegion.ty;
    fillRects.push({
      x: 0,
      y: top,
      width: gridMetrics.rowMarkerWidth,
      height: gridMetrics.rowHeight,
      color:
        row >= selectedRows.start && row <= selectedRows.end
          ? activeHeaderDrag?.kind === "row" && activeHeaderDrag.index === row
            ? HEADER_DRAG_ANCHOR_FILL_COLOR
            : HEADER_SELECTED_FILL_COLOR
          : hoveredHeader?.kind === "row" && hoveredHeader.index === row
            ? HEADER_HOVER_FILL_COLOR
            : HEADER_FILL_COLOR,
    });
    borderRects.push(
      {
        x: gridMetrics.rowMarkerWidth - 1,
        y: top,
        width: 1,
        height: gridMetrics.rowHeight,
        color: GRID_LINE_COLOR,
      },
      {
        x: 0,
        y: top + gridMetrics.rowHeight - 1,
        width: gridMetrics.rowMarkerWidth,
        height: 1,
        color: GRID_LINE_COLOR,
      },
    );
  }

  if (resizeGuideColumn !== null) {
    pushResizeGuideRects({
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideColumn,
      visibleColumns,
      visibleRegion,
    });
  }
}

function resolveAxisSelectionRange(
  fallbackStart: number,
  fallbackEnd: number,
  selection: GridSelection["columns"],
): { start: number; end: number } {
  const start = selection.first();
  const end = selection.last();
  if (start === undefined || end === undefined) {
    return { start: fallbackStart, end: fallbackEnd };
  }
  return { start, end };
}

function pushColumnSelectionBodyRects(options: {
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  selectedColumns: { start: number; end: number };
  visibleColumns: ReadonlyArray<{
    index: number;
    left: number;
    right: number;
    width: number;
  }>;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
}) {
  const { borderRects, fillRects, gridMetrics, selectedColumns, visibleColumns, visibleRegion } =
    options;
  for (const column of visibleColumns) {
    if (column.index < selectedColumns.start || column.index > selectedColumns.end) {
      continue;
    }
    const top = gridMetrics.headerHeight;
    const height = visibleRegion.range.height * gridMetrics.rowHeight;
    fillRects.push({
      x: column.left + 1,
      y: top + 1,
      width: Math.max(0, column.width - 2),
      height: Math.max(0, height - 2),
      color: SELECTION_FILL_COLOR,
    });
    borderRects.push(
      {
        x: column.left,
        y: top,
        width: 2,
        height,
        color: SELECTION_OUTLINE_COLOR,
      },
      {
        x: column.right - 2,
        y: top,
        width: 2,
        height,
        color: SELECTION_OUTLINE_COLOR,
      },
    );
  }
}

function pushRowSelectionBodyRects(options: {
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  selectedRows: { start: number; end: number };
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
  visibleRowEnd: number;
  visibleWidth: number;
}) {
  const {
    borderRects,
    fillRects,
    gridMetrics,
    selectedRows,
    visibleRegion,
    visibleRowEnd,
    visibleWidth,
  } = options;
  if (visibleWidth <= 0) {
    return;
  }
  const bodyLeft = gridMetrics.rowMarkerWidth;
  for (
    let row = Math.max(visibleRegion.range.y, selectedRows.start);
    row <= Math.min(visibleRowEnd, selectedRows.end);
    row += 1
  ) {
    const top =
      gridMetrics.headerHeight +
      (row - visibleRegion.range.y) * gridMetrics.rowHeight -
      visibleRegion.ty;
    fillRects.push({
      x: bodyLeft + 1,
      y: top + 1,
      width: Math.max(0, visibleWidth - 2),
      height: Math.max(0, gridMetrics.rowHeight - 2),
      color: SELECTION_FILL_COLOR,
    });
    borderRects.push(
      {
        x: bodyLeft,
        y: top,
        width: visibleWidth,
        height: 2,
        color: SELECTION_OUTLINE_COLOR,
      },
      {
        x: bodyLeft,
        y: top + gridMetrics.rowHeight - 2,
        width: visibleWidth,
        height: 2,
        color: SELECTION_OUTLINE_COLOR,
      },
    );
  }
}

function pushGridLineRects(
  borderRects: GridGpuRect[],
  rect: Pick<Rectangle, "x" | "y" | "width" | "height">,
  row: number,
  col: number,
  visibleMinRow: number,
  visibleMinCol: number,
) {
  if (row === visibleMinRow) {
    borderRects.push({
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: 1,
      color: GRID_LINE_COLOR,
    });
  }
  if (col === visibleMinCol) {
    borderRects.push({
      x: rect.x,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: GRID_LINE_COLOR,
    });
  }
  borderRects.push({
    x: rect.x,
    y: rect.y + rect.height - 1,
    width: rect.width,
    height: 1,
    color: GRID_LINE_COLOR,
  });
  borderRects.push({
    x: rect.x + rect.width - 1,
    y: rect.y,
    width: 1,
    height: rect.height,
    color: GRID_LINE_COLOR,
  });
}

function pushSelectionRects(options: {
  allowHandle: boolean;
  borderRects: GridGpuRect[];
  fillColor?: GridGpuColor;
  fillRects: GridGpuRect[];
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
  hostBounds: Pick<DOMRect, "left" | "top">;
  outlineColor?: GridGpuColor;
  selectionRange: Pick<Rectangle, "x" | "y" | "width" | "height">;
  visibleMaxCol: number;
  visibleMaxRow: number;
  visibleMinCol: number;
  visibleMinRow: number;
}) {
  const {
    allowHandle,
    borderRects,
    fillColor = SELECTION_FILL_COLOR,
    fillRects,
    getCellBounds,
    hostBounds,
    outlineColor = SELECTION_OUTLINE_COLOR,
    selectionRange,
    visibleMaxCol,
    visibleMaxRow,
    visibleMinCol,
    visibleMinRow,
  } = options;
  const startCol = Math.max(selectionRange.x, visibleMinCol);
  const startRow = Math.max(selectionRange.y, visibleMinRow);
  const endCol = Math.min(selectionRange.x + selectionRange.width - 1, visibleMaxCol);
  const endRow = Math.min(selectionRange.y + selectionRange.height - 1, visibleMaxRow);
  if (startCol > endCol || startRow > endRow) {
    return;
  }

  const startBounds = getCellBounds(startCol, startRow);
  const endBounds = getCellBounds(endCol, endRow);
  if (!startBounds || !endBounds) {
    return;
  }

  const selectionRect = {
    x: startBounds.x - hostBounds.left,
    y: startBounds.y - hostBounds.top,
    width: endBounds.x + endBounds.width - startBounds.x,
    height: endBounds.y + endBounds.height - startBounds.y,
  };
  if (selectionRange.width > 1 || selectionRange.height > 1) {
    fillRects.push({
      x: selectionRect.x + 1,
      y: selectionRect.y + 1,
      width: Math.max(0, selectionRect.width - 2),
      height: Math.max(0, selectionRect.height - 2),
      color: fillColor,
    });
  }

  // Sheets-style range outlines read as a single-pixel stroke, with the fill
  // starting just inside the border so underlying content stays legible.
  const outlineThickness = 1;
  borderRects.push(
    {
      x: selectionRect.x,
      y: selectionRect.y,
      width: selectionRect.width,
      height: outlineThickness,
      color: outlineColor,
    },
    {
      x: selectionRect.x,
      y: selectionRect.y + selectionRect.height - outlineThickness,
      width: selectionRect.width,
      height: outlineThickness,
      color: outlineColor,
    },
    {
      x: selectionRect.x,
      y: selectionRect.y,
      width: outlineThickness,
      height: selectionRect.height,
      color: outlineColor,
    },
    {
      x: selectionRect.x + selectionRect.width - outlineThickness,
      y: selectionRect.y,
      width: outlineThickness,
      height: selectionRect.height,
      color: outlineColor,
    },
  );

  if (!allowHandle) {
    return;
  }
}

function pushBooleanCellRects(
  fillRects: GridGpuRect[],
  borderRects: GridGpuRect[],
  rect: Pick<Rectangle, "x" | "y" | "width" | "height">,
  checked: boolean,
): void {
  const size = Math.max(12, Math.min(16, Math.floor(Math.min(rect.width, rect.height) - 8)));
  const left = Math.round(rect.x + (rect.width - size) / 2);
  const top = Math.round(rect.y + (rect.height - size) / 2);
  const outline = 1;
  const surfaceColor = checked ? CHECKBOX_SELECTED_COLOR : CHECKBOX_SURFACE_COLOR;
  const borderColor = checked ? CHECKBOX_SELECTED_COLOR : CHECKBOX_BORDER_COLOR;

  fillRects.push({
    x: left + outline,
    y: top + outline,
    width: Math.max(0, size - outline * 2),
    height: Math.max(0, size - outline * 2),
    color: surfaceColor,
  });
  borderRects.push(
    {
      x: left,
      y: top,
      width: size,
      height: outline,
      color: borderColor,
    },
    {
      x: left,
      y: top + size - outline,
      width: size,
      height: outline,
      color: borderColor,
    },
    {
      x: left,
      y: top,
      width: outline,
      height: size,
      color: borderColor,
    },
    {
      x: left + size - outline,
      y: top,
      width: outline,
      height: size,
      color: borderColor,
    },
  );

  if (!checked) {
    return;
  }

  fillRects.push(
    {
      x: left + 3,
      y: top + 8,
      width: 2,
      height: 3,
      color: CHECKBOX_CHECK_COLOR,
    },
    {
      x: left + 5,
      y: top + 10,
      width: 2,
      height: 2,
      color: CHECKBOX_CHECK_COLOR,
    },
    {
      x: left + 7,
      y: top + 5,
      width: 2,
      height: 7,
      color: CHECKBOX_CHECK_COLOR,
    },
  );
}

function pushHoveredCellRects(options: {
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
  gridSelection: GridSelection;
  hostBounds: Pick<DOMRect, "left" | "top">;
  hoveredCell: Item;
  selectionRange?: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
}) {
  const {
    borderRects,
    fillRects,
    getCellBounds,
    gridSelection,
    hostBounds,
    hoveredCell,
    selectionRange,
  } = options;
  if (
    selectionRange &&
    hoveredCell[0] >= selectionRange.x &&
    hoveredCell[0] < selectionRange.x + selectionRange.width &&
    hoveredCell[1] >= selectionRange.y &&
    hoveredCell[1] < selectionRange.y + selectionRange.height
  ) {
    return;
  }
  if (gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
    return;
  }
  const bounds = getCellBounds(hoveredCell[0], hoveredCell[1]);
  if (!bounds) {
    return;
  }
  const rect = {
    x: bounds.x - hostBounds.left,
    y: bounds.y - hostBounds.top,
    width: bounds.width,
    height: bounds.height,
  };
  fillRects.push({
    x: rect.x + 1,
    y: rect.y + 1,
    width: Math.max(0, rect.width - 2),
    height: Math.max(0, rect.height - 2),
    color: HOVER_FILL_COLOR,
  });
  borderRects.push(
    {
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: 1,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x,
      y: rect.y + rect.height - 1,
      width: rect.width,
      height: 1,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: HOVER_OUTLINE_COLOR,
    },
    {
      x: rect.x + rect.width - 1,
      y: rect.y,
      width: 1,
      height: rect.height,
      color: HOVER_OUTLINE_COLOR,
    },
  );
}

function pushResizeGuideRects(options: {
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  resizeGuideColumn: number;
  visibleColumns: ReadonlyArray<{
    index: number;
    left: number;
    right: number;
    width: number;
  }>;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
}) {
  const { borderRects, fillRects, gridMetrics, resizeGuideColumn, visibleColumns, visibleRegion } =
    options;
  const column = visibleColumns.find((entry) => entry.index === resizeGuideColumn);
  if (!column) {
    return;
  }
  const lineX = column.right - 1;
  const totalHeight = gridMetrics.headerHeight + visibleRegion.range.height * gridMetrics.rowHeight;
  fillRects.push({
    x: lineX - 2,
    y: 0,
    width: 6,
    height: totalHeight,
    color: RESIZE_GUIDE_GLOW_COLOR,
  });
  borderRects.push({
    x: lineX,
    y: 0,
    width: 2,
    height: totalHeight,
    color: RESIZE_GUIDE_COLOR,
  });
}

function pushColumnHeaderDragGuideRects(options: {
  activeHeaderDrag: HeaderSelection;
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  selectedColumns: { start: number; end: number };
  visibleColumns: ReadonlyArray<{
    index: number;
    left: number;
    right: number;
    width: number;
  }>;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
}) {
  const {
    activeHeaderDrag,
    borderRects,
    fillRects,
    gridMetrics,
    selectedColumns,
    visibleColumns,
    visibleRegion,
  } = options;
  const startColumn = visibleColumns.find((entry) => entry.index === selectedColumns.start);
  const endColumn = visibleColumns.find((entry) => entry.index === selectedColumns.end);
  if (!startColumn || !endColumn) {
    return;
  }
  const left = startColumn.left;
  const right = endColumn.right;
  const totalHeight = gridMetrics.headerHeight + visibleRegion.range.height * gridMetrics.rowHeight;
  fillRects.push({
    x: left,
    y: 0,
    width: right - left,
    height: 3,
    color: RESIZE_GUIDE_GLOW_COLOR,
  });
  fillRects.push({
    x: left,
    y: totalHeight - 3,
    width: right - left,
    height: 3,
    color: RESIZE_GUIDE_GLOW_COLOR,
  });
  borderRects.push(
    {
      x: left,
      y: 0,
      width: 2,
      height: totalHeight,
      color: RESIZE_GUIDE_COLOR,
    },
    {
      x: right - 2,
      y: 0,
      width: 2,
      height: totalHeight,
      color: RESIZE_GUIDE_COLOR,
    },
  );
  const anchorColumn = visibleColumns.find((entry) => entry.index === activeHeaderDrag.index);
  if (anchorColumn) {
    fillRects.push({
      x: anchorColumn.left,
      y: gridMetrics.headerHeight - 3,
      width: anchorColumn.width,
      height: 3,
      color: RESIZE_GUIDE_COLOR,
    });
  }
}

function pushRowHeaderDragGuideRects(options: {
  activeHeaderDrag: HeaderSelection;
  borderRects: GridGpuRect[];
  fillRects: GridGpuRect[];
  gridMetrics: GridMetrics;
  selectedRows: { start: number; end: number };
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
  visibleRowEnd: number;
  visibleWidth: number;
}) {
  const {
    activeHeaderDrag,
    borderRects,
    fillRects,
    gridMetrics,
    selectedRows,
    visibleRegion,
    visibleRowEnd,
    visibleWidth,
  } = options;
  if (visibleWidth <= 0) {
    return;
  }
  const topRow = Math.max(visibleRegion.range.y, selectedRows.start);
  const bottomRow = Math.min(visibleRowEnd, selectedRows.end);
  if (topRow > bottomRow) {
    return;
  }
  const top =
    gridMetrics.headerHeight +
    (topRow - visibleRegion.range.y) * gridMetrics.rowHeight -
    visibleRegion.ty;
  const bottom =
    gridMetrics.headerHeight +
    (bottomRow - visibleRegion.range.y + 1) * gridMetrics.rowHeight -
    visibleRegion.ty;
  const totalWidth = gridMetrics.rowMarkerWidth + visibleWidth;
  fillRects.push({
    x: 0,
    y: top,
    width: 3,
    height: bottom - top,
    color: RESIZE_GUIDE_GLOW_COLOR,
  });
  fillRects.push({
    x: totalWidth - 3,
    y: top,
    width: 3,
    height: bottom - top,
    color: RESIZE_GUIDE_GLOW_COLOR,
  });
  borderRects.push(
    {
      x: 0,
      y: top,
      width: totalWidth,
      height: 2,
      color: RESIZE_GUIDE_COLOR,
    },
    {
      x: 0,
      y: bottom - 2,
      width: totalWidth,
      height: 2,
      color: RESIZE_GUIDE_COLOR,
    },
  );
  if (activeHeaderDrag.index >= topRow && activeHeaderDrag.index <= bottomRow) {
    const anchorTop =
      gridMetrics.headerHeight +
      (activeHeaderDrag.index - visibleRegion.range.y) * gridMetrics.rowHeight -
      visibleRegion.ty;
    fillRects.push({
      x: gridMetrics.rowMarkerWidth - 3,
      y: anchorTop,
      width: 3,
      height: gridMetrics.rowHeight,
      color: RESIZE_GUIDE_COLOR,
    });
  }
}

function createBorderRects(
  rect: Pick<Rectangle, "x" | "y" | "width" | "height">,
  side: "top" | "right" | "bottom" | "left",
  border: NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]>,
): GridGpuRect[] {
  const thickness = border.weight === "thick" ? 3 : border.weight === "medium" ? 2 : 1;
  const isHorizontal = side === "top" || side === "bottom";
  const edgeX = side === "left" ? rect.x : side === "right" ? rect.x + rect.width - 1 : rect.x;
  const edgeY = side === "top" ? rect.y : side === "bottom" ? rect.y + rect.height - 1 : rect.y;
  const length = isHorizontal ? rect.width : rect.height;
  const color = parseGpuColor(border.color);

  if (length <= 0) {
    return [];
  }

  switch (border.style) {
    case "dashed":
      return createPatternBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal, 6, 4);
    case "dotted":
      return createPatternBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal, 1, 3);
    case "double":
      return createDoubleBorderRects(edgeX, edgeY, length, thickness, color, isHorizontal);
    case "solid":
    default:
      return [
        {
          x: isHorizontal ? edgeX : edgeX - thickness / 2,
          y: isHorizontal ? edgeY - thickness / 2 : edgeY,
          width: isHorizontal ? length : thickness,
          height: isHorizontal ? thickness : length,
          color,
        },
      ];
  }
}

function createPatternBorderRects(
  edgeX: number,
  edgeY: number,
  length: number,
  thickness: number,
  color: GridGpuColor,
  isHorizontal: boolean,
  segmentLength: number,
  gapLength: number,
): GridGpuRect[] {
  const rects: GridGpuRect[] = [];
  for (let cursor = 0; cursor < length; cursor += segmentLength + gapLength) {
    const currentLength = Math.min(segmentLength, length - cursor);
    rects.push({
      x: isHorizontal ? edgeX + cursor : edgeX - thickness / 2,
      y: isHorizontal ? edgeY - thickness / 2 : edgeY + cursor,
      width: isHorizontal ? currentLength : thickness,
      height: isHorizontal ? thickness : currentLength,
      color,
    });
  }
  return rects;
}

function createDoubleBorderRects(
  edgeX: number,
  edgeY: number,
  length: number,
  thickness: number,
  color: GridGpuColor,
  isHorizontal: boolean,
): GridGpuRect[] {
  const span = Math.max(3, thickness + 2);
  const offset = span / 2;
  if (isHorizontal) {
    return [
      {
        x: edgeX,
        y: edgeY - offset,
        width: length,
        height: 1,
        color,
      },
      {
        x: edgeX,
        y: edgeY - offset + span - 1,
        width: length,
        height: 1,
        color,
      },
    ];
  }
  return [
    {
      x: edgeX - offset,
      y: edgeY,
      width: 1,
      height: length,
      color,
    },
    {
      x: edgeX - offset + span - 1,
      y: edgeY,
      width: 1,
      height: length,
      color,
    },
  ];
}

export function parseGpuColor(input: string | undefined): GridGpuColor {
  if (!input) {
    return FALLBACK_COLOR;
  }

  const color = input.trim();
  if (color === "transparent") {
    return { r: 0, g: 0, b: 0, a: 0 };
  }

  if (color.startsWith("#")) {
    return parseHexGpuColor(color);
  }

  const rgbaMatch = color.match(/^rgba?\(([^)]+)\)$/i);
  if (rgbaMatch) {
    const parts = (rgbaMatch[1] ?? "")
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean);
    const [r = "0", g = "0", b = "0", a = "1"] = parts;
    return {
      r: clampColorChannel(Number.parseFloat(r) / 255),
      g: clampColorChannel(Number.parseFloat(g) / 255),
      b: clampColorChannel(Number.parseFloat(b) / 255),
      a: clampColorChannel(Number.parseFloat(a)),
    };
  }

  return FALLBACK_COLOR;
}

function parseHexGpuColor(input: string): GridGpuColor {
  const hex = input.slice(1);
  switch (hex.length) {
    case 3:
      return {
        r: hexPairToChannel((hex.slice(0, 1) || "0").repeat(2)),
        g: hexPairToChannel((hex.slice(1, 2) || "0").repeat(2)),
        b: hexPairToChannel((hex.slice(2, 3) || "0").repeat(2)),
        a: 1,
      };
    case 4:
      return {
        r: hexPairToChannel((hex.slice(0, 1) || "0").repeat(2)),
        g: hexPairToChannel((hex.slice(1, 2) || "0").repeat(2)),
        b: hexPairToChannel((hex.slice(2, 3) || "0").repeat(2)),
        a: hexPairToChannel((hex.slice(3, 4) || "f").repeat(2)),
      };
    case 6:
      return {
        r: hexPairToChannel(hex.slice(0, 2)),
        g: hexPairToChannel(hex.slice(2, 4)),
        b: hexPairToChannel(hex.slice(4, 6)),
        a: 1,
      };
    case 8:
      return {
        r: hexPairToChannel(hex.slice(0, 2)),
        g: hexPairToChannel(hex.slice(2, 4)),
        b: hexPairToChannel(hex.slice(4, 6)),
        a: hexPairToChannel(hex.slice(6, 8)),
      };
    default:
      return FALLBACK_COLOR;
  }
}

function hexPairToChannel(value: string): number {
  return clampColorChannel(Number.parseInt(value, 16) / 255);
}

function clampColorChannel(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return Math.min(1, Math.max(0, value));
}
