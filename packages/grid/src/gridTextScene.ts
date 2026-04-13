import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import type { GridEngineLike } from "./grid-engine.js";
import { getResolvedCellFontFamily, snapshotToRenderCell } from "./gridCells.js";
import { getVisibleColumnBounds, getVisibleRowBounds, type GridMetrics } from "./gridMetrics.js";
import { indexToColumn } from "@bilig/formula";
import type { HeaderSelection } from "./gridPointer.js";
import type { Item, Rectangle } from "./gridTypes.js";
import { collectVisibleColumnBounds, collectVisibleRowBounds } from "./visibleGridAxes.js";

export interface GridTextItem {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly clipInsetTop: number;
  readonly clipInsetRight: number;
  readonly clipInsetBottom: number;
  readonly clipInsetLeft: number;
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
    readonly freezeRows?: number;
    readonly freezeCols?: number;
  };
  readonly gridMetrics: GridMetrics;
  readonly columnWidths: Readonly<Record<number, number>>;
  readonly rowHeights?: Readonly<Record<number, number>>;
  readonly editingCell?: Item | null;
  readonly selectedCell: Item;
  readonly selectedCellSnapshot?: CellSnapshot | null;
  readonly selectionRange?: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  readonly hoveredHeader?: HeaderSelection | null;
  readonly activeHeaderDrag?: HeaderSelection | null;
  readonly resizeGuideColumn?: number | null;
  readonly hostBounds: Pick<DOMRect, "left" | "top" | "width" | "height">;
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
  rowHeights = {},
  editingCell = null,
  selectedCell,
  selectedCellSnapshot = null,
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
    rowHeights,
    selectedCell,
    selectionRange,
    hoveredHeader,
    activeHeaderDrag,
    resizeGuideColumn,
    visibleRegion,
    visibleItems,
    getCellBounds,
  });

  for (const [col, row] of visibleItems) {
    if (editingCell && editingCell[0] === col && editingCell[1] === row) {
      continue;
    }

    const item = buildCellTextItem({
      engine,
      sheetName,
      col,
      row,
      hostBounds,
      getCellBounds,
      visibleColumnEnd: Math.max(...visibleItems.map(([visibleCol]) => visibleCol)),
      selectedAddress:
        col === selectedCell[0] && row === selectedCell[1]
          ? `${indexToColumn(col)}${row + 1}`
          : null,
      snapshotOverride: selectedCellSnapshot,
      gridMetrics,
    });
    if (item) {
      items.push(item);
    }
  }

  return { items };
}

function buildCellTextItem({
  engine,
  sheetName,
  col,
  row,
  hostBounds,
  getCellBounds,
  visibleColumnEnd,
  selectedAddress = null,
  snapshotOverride = null,
  gridMetrics,
}: {
  engine: GridEngineLike;
  sheetName: string;
  col: number;
  row: number;
  hostBounds: Pick<DOMRect, "left" | "top" | "width" | "height">;
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
  visibleColumnEnd: number;
  selectedAddress?: string | null;
  snapshotOverride?: CellSnapshot | null;
  gridMetrics: GridMetrics;
}): GridTextItem | null {
  const bounds = getCellBounds(col, row);
  if (!bounds) {
    return null;
  }

  const address = `${indexToColumn(col)}${row + 1}`;
  const snapshot = resolveCellTextSnapshot({
    address,
    engine,
    sheetName,
    selectedAddress,
    snapshotOverride,
  });
  if (snapshot.value.tag === ValueTag.Boolean) {
    return null;
  }

  const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId));
  if (renderCell.displayText.length === 0) {
    return null;
  }

  const renderBounds = resolveTextRenderBounds({
    engine,
    sheetName,
    row,
    col,
    bounds,
    visibleColumnEnd,
    getCellBounds,
    renderCell,
  });

  const localBounds = {
    x: renderBounds.x - hostBounds.left,
    y: renderBounds.y - hostBounds.top,
    width: renderBounds.width,
    height: renderBounds.height,
  };

  return {
    ...localBounds,
    ...resolveClipInsets({
      bounds: localBounds,
      clipRect: {
        x: gridMetrics.rowMarkerWidth,
        y: gridMetrics.headerHeight,
        width: hostBounds.width - gridMetrics.rowMarkerWidth,
        height: hostBounds.height - gridMetrics.headerHeight,
      },
    }),
    text: renderCell.displayText,
    align: renderCell.align,
    wrap: renderCell.wrap,
    color: renderCell.color,
    font: renderCell.font,
    fontSize: renderCell.fontSize,
    underline: renderCell.underline,
    strike: false,
  };
}

function resolveClipInsets(options: {
  bounds: Rectangle;
  clipRect: Rectangle;
}): Pick<GridTextItem, "clipInsetTop" | "clipInsetRight" | "clipInsetBottom" | "clipInsetLeft"> {
  const { bounds, clipRect } = options;
  return {
    clipInsetTop: Math.max(0, clipRect.y - bounds.y),
    clipInsetRight: Math.max(0, bounds.x + bounds.width - (clipRect.x + clipRect.width)),
    clipInsetBottom: Math.max(0, bounds.y + bounds.height - (clipRect.y + clipRect.height)),
    clipInsetLeft: Math.max(0, clipRect.x - bounds.x),
  };
}

function resolveCellTextSnapshot({
  address,
  engine,
  sheetName,
  selectedAddress,
  snapshotOverride,
}: {
  address: string;
  engine: GridEngineLike;
  sheetName: string;
  selectedAddress: string | null;
  snapshotOverride: CellSnapshot | null;
}): CellSnapshot {
  const engineSnapshot = engine.getCell(sheetName, address);
  if (!snapshotOverride || selectedAddress !== address || snapshotOverride.address !== address) {
    return engineSnapshot;
  }

  const engineRenderCell = snapshotToRenderCell(
    engineSnapshot,
    engine.getCellStyle(engineSnapshot.styleId),
  );
  const overrideRenderCell = snapshotToRenderCell(
    snapshotOverride,
    engine.getCellStyle(snapshotOverride.styleId),
  );

  if (engineRenderCell.displayText.length > 0 && overrideRenderCell.displayText.length === 0) {
    return engineSnapshot;
  }
  if (overrideRenderCell.displayText.length > 0 && engineRenderCell.displayText.length === 0) {
    return snapshotOverride;
  }

  return snapshotOverride;
}

function resolveTextRenderBounds(options: {
  engine: GridEngineLike;
  sheetName: string;
  row: number;
  col: number;
  bounds: Rectangle;
  visibleColumnEnd: number;
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
  renderCell: ReturnType<typeof snapshotToRenderCell>;
}): Rectangle {
  const { engine, sheetName, row, col, bounds, visibleColumnEnd, getCellBounds, renderCell } =
    options;

  if (
    renderCell.wrap ||
    renderCell.align !== "left" ||
    (renderCell.kind !== "string" && renderCell.kind !== "error")
  ) {
    return bounds;
  }

  let spillWidth = bounds.width;
  for (let spillCol = col + 1; spillCol <= visibleColumnEnd; spillCol += 1) {
    const spillBounds = getCellBounds(spillCol, row);
    if (!spillBounds) {
      break;
    }
    const spillSnapshot = engine.getCell(sheetName, `${indexToColumn(spillCol)}${row + 1}`);
    const spillRenderCell = snapshotToRenderCell(
      spillSnapshot,
      engine.getCellStyle(spillSnapshot.styleId),
    );
    if (spillRenderCell.displayText.length > 0) {
      break;
    }
    spillWidth = spillBounds.x + spillBounds.width - bounds.x;
  }

  if (spillWidth === bounds.width) {
    return bounds;
  }

  return {
    ...bounds,
    width: spillWidth,
  };
}

function pushHeaderTextItems(options: {
  columnWidths: Readonly<Record<number, number>>;
  gridMetrics: GridMetrics;
  items: GridTextItem[];
  rowHeights: Readonly<Record<number, number>>;
  selectedCell: Item;
  selectionRange: Pick<Rectangle, "x" | "y" | "width" | "height"> | null;
  hoveredHeader: HeaderSelection | null;
  activeHeaderDrag: HeaderSelection | null;
  resizeGuideColumn: number | null;
  visibleRegion: {
    readonly range: Pick<Rectangle, "x" | "y" | "width" | "height">;
    readonly tx: number;
    readonly ty: number;
    readonly freezeRows?: number;
    readonly freezeCols?: number;
  };
  visibleItems: readonly Item[];
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
}) {
  const {
    columnWidths,
    gridMetrics,
    items,
    rowHeights,
    selectedCell,
    selectionRange,
    hoveredHeader,
    activeHeaderDrag,
    resizeGuideColumn,
    visibleRegion,
    visibleItems,
    getCellBounds,
  } = options;
  const selectedColumns = resolveSelectionRange(
    selectionRange?.x ?? selectedCell[0],
    selectionRange ? selectionRange.x + selectionRange.width - 1 : selectedCell[0],
  );
  const selectedRows = resolveSelectionRange(
    selectionRange?.y ?? selectedCell[1],
    selectionRange ? selectionRange.y + selectionRange.height - 1 : selectedCell[1],
  );

  const hasFrozenAxes = (visibleRegion.freezeRows ?? 0) > 0 || (visibleRegion.freezeCols ?? 0) > 0;
  const visibleColumns = hasFrozenAxes
    ? collectVisibleColumnBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleColumnBounds(
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
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: Math.max(0, gridMetrics.rowMarkerWidth - column.left),
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

  const visibleRows = hasFrozenAxes
    ? collectVisibleRowBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleRowBounds(
        visibleRegion.range,
        gridMetrics.headerHeight - visibleRegion.ty,
        Number.MAX_SAFE_INTEGER,
        rowHeights,
        gridMetrics.rowHeight,
      );
  for (const row of visibleRows) {
    items.push({
      x: 0,
      y: row.top,
      width: gridMetrics.rowMarkerWidth,
      height: row.height,
      clipInsetTop: Math.max(0, gridMetrics.headerHeight - row.top),
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: String(row.index + 1),
      align: "right",
      wrap: false,
      color: resolveHeaderTextColor({
        activeHeaderDrag,
        hoveredHeader,
        index: row.index,
        isSelected: row.index >= selectedRows.start && row.index <= selectedRows.end,
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
