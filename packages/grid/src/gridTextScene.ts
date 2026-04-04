import { ValueTag } from "@bilig/protocol";
import type { GridEngineLike } from "./grid-engine.js";
import { getResolvedCellFontFamily, snapshotToRenderCell } from "./gridCells.js";
import type { GridMetrics } from "./gridMetrics.js";
import { getVisibleColumnBounds } from "./gridMetrics.js";
import { indexToColumn } from "@bilig/formula";
import type { HeaderSelection } from "./gridPointer.js";
import type { Item, Rectangle } from "./gridTypes.js";

export interface GridTextItem {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly text: string;
  readonly align: "left" | "center" | "right";
  readonly wrap: boolean;
  readonly color: string;
  readonly font: string;
  readonly fontSize: number;
  readonly underline: boolean;
  readonly strike: boolean;
}

export interface GridTextScene {
  readonly items: readonly GridTextItem[];
}

interface BuildGridTextSceneOptions {
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
  readonly selectedCell: Item;
  readonly selectionRange?: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  readonly hoveredHeader?: HeaderSelection | null;
  readonly activeHeaderDrag?: HeaderSelection | null;
  readonly resizeGuideColumn?: number | null;
  readonly hostBounds: Pick<DOMRect, "left" | "top">;
  readonly getCellBounds: (col: number, row: number) => Rectangle | undefined;
}

const HEADER_TEXT_COLOR = "#5f6368";
const HEADER_HOVER_TEXT_COLOR = "#3c4043";
const HEADER_SELECTED_TEXT_COLOR = "#1f7a43";
const HEADER_DRAG_ANCHOR_TEXT_COLOR = "#176239";
const HEADER_RESIZE_TEXT_COLOR = "#176239";
const HEADER_FONT = `500 11px ${getResolvedCellFontFamily()}`;

export function buildGridTextScene({
  engine,
  sheetName,
  visibleItems,
  visibleRegion,
  gridMetrics,
  columnWidths,
  selectedCell,
  selectionRange = null,
  hoveredHeader = null,
  activeHeaderDrag = null,
  resizeGuideColumn = null,
  hostBounds,
  getCellBounds,
}: BuildGridTextSceneOptions): GridTextScene {
  const items: GridTextItem[] = [];

  pushHeaderTextItems({
    columnWidths,
    gridMetrics,
    items,
    selectedCell,
    selectionRange,
    hoveredHeader,
    activeHeaderDrag,
    resizeGuideColumn,
    visibleRegion,
  });

  for (const [col, row] of visibleItems) {
    const bounds = getCellBounds(col, row);
    if (!bounds) {
      continue;
    }

    const snapshot = engine.getCell(sheetName, `${indexToColumn(col)}${row + 1}`);
    if (snapshot.value.tag === ValueTag.Boolean) {
      continue;
    }

    const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId));
    if (renderCell.displayText.length === 0) {
      continue;
    }

    items.push({
      x: bounds.x - hostBounds.left,
      y: bounds.y - hostBounds.top,
      width: bounds.width,
      height: bounds.height,
      text: renderCell.displayText,
      align: renderCell.align,
      wrap: renderCell.wrap,
      color: renderCell.color,
      font: renderCell.font,
      fontSize: renderCell.fontSize,
      underline: renderCell.underline,
      strike: false,
    });
  }

  return { items };
}

function pushHeaderTextItems(options: {
  columnWidths: Readonly<Record<number, number>>;
  gridMetrics: GridMetrics;
  items: GridTextItem[];
  selectedCell: Item;
  selectionRange: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  hoveredHeader: HeaderSelection | null;
  activeHeaderDrag: HeaderSelection | null;
  resizeGuideColumn: number | null;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
  };
}) {
  const {
    columnWidths,
    gridMetrics,
    items,
    selectedCell,
    selectionRange,
    hoveredHeader,
    activeHeaderDrag,
    resizeGuideColumn,
    visibleRegion,
  } = options;
  const selectedColumns = resolveSelectionRange(
    selectionRange?.x ?? selectedCell[0],
    selectionRange ? selectionRange.x + selectionRange.width - 1 : selectedCell[0],
  );
  const selectedRows = resolveSelectionRange(
    selectionRange?.y ?? selectedCell[1],
    selectionRange ? selectionRange.y + selectionRange.height - 1 : selectedCell[1],
  );

  const visibleColumns = getVisibleColumnBounds(
    visibleRegion.range,
    gridMetrics.rowMarkerWidth - visibleRegion.tx,
    Number.MAX_SAFE_INTEGER,
    columnWidths,
    gridMetrics.columnWidth,
  );
  for (const column of visibleColumns) {
    items.push({
      x: column.left,
      y: 0,
      width: column.width,
      height: gridMetrics.headerHeight,
      text: indexToColumn(column.index),
      align: "center",
      wrap: false,
      color: resolveHeaderTextColor({
        activeHeaderDrag,
        hoveredHeader,
        index: column.index,
        isSelected: column.index >= selectedColumns.start && column.index <= selectedColumns.end,
        kind: "column",
        resizeGuideColumn,
      }),
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    });
  }

  const visibleRowEnd = visibleRegion.range.y + visibleRegion.range.height - 1;
  for (let row = visibleRegion.range.y; row <= visibleRowEnd; row += 1) {
    const top =
      gridMetrics.headerHeight +
      (row - visibleRegion.range.y) * gridMetrics.rowHeight -
      visibleRegion.ty;
    items.push({
      x: 0,
      y: top,
      width: gridMetrics.rowMarkerWidth,
      height: gridMetrics.rowHeight,
      text: String(row + 1),
      align: "right",
      wrap: false,
      color: resolveHeaderTextColor({
        activeHeaderDrag,
        hoveredHeader,
        index: row,
        isSelected: row >= selectedRows.start && row <= selectedRows.end,
        kind: "row",
        resizeGuideColumn,
      }),
      font: HEADER_FONT,
      fontSize: 11,
      underline: false,
      strike: false,
    });
  }
}

function resolveSelectionRange(start: number, end: number): { start: number; end: number } {
  return { start, end };
}

function resolveHeaderTextColor(options: {
  activeHeaderDrag: HeaderSelection | null;
  hoveredHeader: HeaderSelection | null;
  index: number;
  isSelected: boolean;
  kind: HeaderSelection["kind"];
  resizeGuideColumn: number | null;
}): string {
  const { activeHeaderDrag, hoveredHeader, index, isSelected, kind, resizeGuideColumn } = options;
  if (kind === "column" && resizeGuideColumn === index) {
    return HEADER_RESIZE_TEXT_COLOR;
  }
  if (isSelected) {
    if (activeHeaderDrag?.kind === kind && activeHeaderDrag.index === index) {
      return HEADER_DRAG_ANCHOR_TEXT_COLOR;
    }
    return HEADER_SELECTED_TEXT_COLOR;
  }
  if (hoveredHeader?.kind === kind && hoveredHeader.index === index) {
    return HEADER_HOVER_TEXT_COLOR;
  }
  return HEADER_TEXT_COLOR;
}
