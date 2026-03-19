import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { MAX_COLS, MAX_ROWS, ValueTag, formatErrorCode } from "@bilig/protocol";
import {
  CompactSelection,
  DataEditor,
  type FillPatternEventArgs,
  GridCellKind,
  type DataEditorRef,
  type GridCell,
  type GridColumn,
  type GridKeyEventArgs,
  type GridSelection,
  type Item,
  type Rectangle
} from "@glideapps/glide-data-grid";
import { CellEditorOverlay } from "./CellEditorOverlay.js";

export type EditMovement = readonly [-1 | 0 | 1, -1 | 0 | 1];

interface SheetGridViewProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  variant?: "playground" | "product";
  selectedAddr: string;
  editorValue: string;
  resolvedValue: string;
  isEditingCell: boolean;
  onSelect(addr: string): void;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
  onBeginEdit(seed?: string): void;
  onEditorChange(next: string): void;
  onCommitEdit(movement?: EditMovement): void;
  onCancelEdit(): void;
  onClearCell(): void;
  onFillRange(sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onCopyRange(sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onPaste(addr: string, values: readonly (readonly string[])[]): void;
}

interface VisibleRegionState {
  range: Rectangle;
  tx: number;
  ty: number;
}

interface PointerGeometry {
  hostBounds: DOMRect;
  cellWidth: number;
  cellHeight: number;
  dataLeft: number;
  dataTop: number;
  dataRight: number;
  dataBottom: number;
}

interface InternalClipboardRange {
  sourceStartAddress: string;
  sourceEndAddress: string;
  signature: string;
  plainText: string;
  rowCount: number;
  colCount: number;
}

type HeaderSelection =
  | { kind: "column"; index: number }
  | { kind: "row"; index: number };

const PLAYGROUND_COLUMN_WIDTH = 120;
const PLAYGROUND_ROW_HEIGHT = 28;
const PLAYGROUND_HEADER_HEIGHT = 30;
const PLAYGROUND_ROW_MARKER_WIDTH = 60;
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;
const PRODUCT_HEADER_HEIGHT = 24;
const PRODUCT_ROW_MARKER_WIDTH = 46;
const SCROLLBAR_GUTTER = 18;
const COLUMN_RESIZE_HANDLE_THRESHOLD = 6;
const MIN_COLUMN_WIDTH = 44;
const MAX_COLUMN_WIDTH = 480;
const EMPTY_COLUMN_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});

function getGridMetrics(variant: "playground" | "product") {
  return variant === "product"
    ? {
        columnWidth: PRODUCT_COLUMN_WIDTH,
        rowHeight: PRODUCT_ROW_HEIGHT,
        headerHeight: PRODUCT_HEADER_HEIGHT,
        rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH
      }
    : {
        columnWidth: PLAYGROUND_COLUMN_WIDTH,
        rowHeight: PLAYGROUND_ROW_HEIGHT,
        headerHeight: PLAYGROUND_HEADER_HEIGHT,
        rowMarkerWidth: PLAYGROUND_ROW_MARKER_WIDTH
      };
}

function createGridSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: []
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty()
  };
}

function clampCell([col, row]: Item): Item {
  return [
    Math.min(MAX_COLS - 1, Math.max(0, col)),
    Math.min(MAX_ROWS - 1, Math.max(0, row))
  ];
}

function clampSelectionRange(range: Rectangle): Rectangle {
  const x = Math.min(MAX_COLS - 1, Math.max(0, range.x));
  const y = Math.min(MAX_ROWS - 1, Math.max(0, range.y));
  const maxWidth = MAX_COLS - x;
  const maxHeight = MAX_ROWS - y;
  return {
    x,
    y,
    width: Math.max(1, Math.min(maxWidth, range.width)),
    height: Math.max(1, Math.min(maxHeight, range.height))
  };
}

function rectangleToAddresses(range: Rectangle): { startAddress: string; endAddress: string } {
  const clamped = clampSelectionRange(range);
  return {
    startAddress: formatAddress(clamped.y, clamped.x),
    endAddress: formatAddress(clamped.y + clamped.height - 1, clamped.x + clamped.width - 1)
  };
}

function getResolvedColumnWidth(
  columnWidths: Readonly<Record<number, number>>,
  col: number,
  defaultWidth: number
): number {
  return columnWidths[col] ?? defaultWidth;
}

function getVisibleColumnBounds(
  region: VisibleRegionState,
  dataLeft: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number
): Array<{ index: number; left: number; right: number; width: number }> {
  const bounds: Array<{ index: number; left: number; right: number; width: number }> = [];
  const colEnd = Math.min(MAX_COLS - 1, region.range.x + region.range.width - 1);
  let cursor = dataLeft;
  for (let col = region.range.x; col <= colEnd; col += 1) {
    const width = getResolvedColumnWidth(columnWidths, col, defaultWidth);
    bounds.push({ index: col, left: cursor, right: cursor + width, width });
    cursor += width;
  }
  return bounds;
}

function resolveColumnAtClientX(
  clientX: number,
  region: VisibleRegionState,
  dataLeft: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number
): number | null {
  for (const column of getVisibleColumnBounds(region, dataLeft, columnWidths, defaultWidth)) {
    if (clientX >= column.left && clientX < column.right) {
      return column.index;
    }
  }
  return null;
}

function resolveColumnResizeTarget(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number
): number | null {
  if (clientY < geometry.hostBounds.top || clientY >= geometry.dataTop) {
    return null;
  }
  for (const column of getVisibleColumnBounds(region, geometry.dataLeft, columnWidths, defaultWidth)) {
    if (column.index >= MAX_COLS - 1) {
      continue;
    }
    if (clientX >= column.right - COLUMN_RESIZE_HANDLE_THRESHOLD && clientX <= column.right + COLUMN_RESIZE_HANDLE_THRESHOLD) {
      return column.index;
    }
  }
  return null;
}

function createRangeSelection(base: GridSelection, anchor: Item, target: Item): GridSelection {
  const startCol = Math.min(anchor[0], target[0]);
  const endCol = Math.max(anchor[0], target[0]);
  const startRow = Math.min(anchor[1], target[1]);
  const endRow = Math.max(anchor[1], target[1]);

  return {
    ...base,
    current: {
      cell: anchor,
      range: {
        x: startCol,
        y: startRow,
        width: endCol - startCol + 1,
        height: endRow - startRow + 1
      },
      rangeStack: []
    }
  };
}

function formatSelectionSummary(selection: GridSelection, fallbackAddress: string): string {
  const selectedColumnStart = selection.columns.first();
  const selectedColumnEnd = selection.columns.last();
  if (selectedColumnStart !== undefined && selectedColumnEnd !== undefined) {
    const start = indexToColumn(selectedColumnStart);
    const end = indexToColumn(selectedColumnEnd);
    return start === end ? `${start}:${start}` : `${start}:${end}`;
  }

  const selectedRowStart = selection.rows.first();
  const selectedRowEnd = selection.rows.last();
  if (selectedRowStart !== undefined && selectedRowEnd !== undefined) {
    const start = String(selectedRowStart + 1);
    const end = String(selectedRowEnd + 1);
    return start === end ? `${start}:${start}` : `${start}:${end}`;
  }

  const range = selection.current?.range;
  if (!range) {
    return fallbackAddress;
  }
  const start = formatAddress(range.y, range.x);
  if (range.width === 1 && range.height === 1) {
    return start;
  }
  const end = formatAddress(range.y + range.height - 1, range.x + range.width - 1);
  return `${start}:${end}`;
}

function isPrintableKey(event: GridKeyEventArgs): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false;
  }
  return event.key.length === 1;
}

function isNavigationKey(key: string): boolean {
  return key === "ArrowUp" || key === "ArrowDown" || key === "ArrowLeft" || key === "ArrowRight";
}

function isClipboardShortcut(event: Pick<GridKeyEventArgs, "altKey" | "ctrlKey" | "key" | "metaKey">): boolean {
  if (!(event.ctrlKey || event.metaKey) || event.altKey) {
    return false;
  }
  const normalizedKey = event.key.toLowerCase();
  return normalizedKey === "c" || normalizedKey === "x" || normalizedKey === "v";
}

function sameBounds(left: Rectangle | undefined, right: Rectangle | undefined): boolean {
  if (left === right) {
    return true;
  }
  if (!left || !right) {
    return false;
  }
  return left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height;
}

function createColumnSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: []
    },
    columns: CompactSelection.fromSingleSelection(col),
    rows: CompactSelection.empty()
  };
}

function createColumnSliceSelection(startCol: number, endCol: number, row: number): GridSelection {
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);
  return {
    current: {
      cell: [startCol, row],
      range: { x: left, y: row, width: right - left + 1, height: 1 },
      rangeStack: []
    },
    columns: CompactSelection.fromSingleSelection([left, right + 1]),
    rows: CompactSelection.empty()
  };
}

function createRowSelection(col: number, row: number): GridSelection {
  return {
    current: {
      cell: [col, row],
      range: { x: col, y: row, width: 1, height: 1 },
      rangeStack: []
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.fromSingleSelection(row)
  };
}

function createRowSliceSelection(col: number, startRow: number, endRow: number): GridSelection {
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  return {
    current: {
      cell: [col, startRow],
      range: { x: col, y: top, width: 1, height: bottom - top + 1 },
      rangeStack: []
    },
    columns: CompactSelection.empty(),
    rows: CompactSelection.fromSingleSelection([top, bottom + 1])
  };
}

function cellToGridCell(engine: SpreadsheetEngine, sheetName: string, addr: string): GridCell {
  const snapshot = selectors.selectCellSnapshot(engine, sheetName, addr);
  const rawValue = cellToEditorSeed(snapshot);

  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return {
        kind: GridCellKind.Number,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: String(snapshot.value.value),
        readonly: false,
        copyData: snapshot.formula ? rawValue : String(snapshot.value.value),
        contentAlign: "right"
      };
    case ValueTag.Boolean:
      return {
        kind: GridCellKind.Boolean,
        allowOverlay: false,
        data: snapshot.value.value,
        readonly: false,
        copyData: snapshot.formula ? rawValue : snapshot.value.value ? "TRUE" : "FALSE"
      };
    case ValueTag.Error:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: formatErrorCode(snapshot.value.code),
        displayData: formatErrorCode(snapshot.value.code),
        readonly: false,
        copyData: snapshot.formula ? rawValue : formatErrorCode(snapshot.value.code)
      };
    case ValueTag.String:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: snapshot.value.value,
        displayData: snapshot.value.value,
        readonly: false,
        copyData: snapshot.formula ? rawValue : snapshot.value.value
      };
    default:
      return {
        kind: GridCellKind.Text,
        allowOverlay: false,
        data: "",
        displayData: "",
        readonly: false,
        copyData: snapshot.formula ? rawValue : ""
      };
  }
}

function cellToEditorSeed(snapshot: ReturnType<typeof selectors.selectCellSnapshot>): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`;
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    switch (snapshot.value.tag) {
      case ValueTag.Number:
        return String(snapshot.value.value);
      case ValueTag.Boolean:
        return snapshot.value.value ? "TRUE" : "FALSE";
      case ValueTag.String:
        return snapshot.value.value;
      case ValueTag.Error:
        return formatErrorCode(snapshot.value.code);
      default:
        return "";
    }
  }
  if (typeof snapshot.input === "boolean") {
    return snapshot.input ? "TRUE" : "FALSE";
  }
  return String(snapshot.input);
}

function serializeClipboardMatrix(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join("\u001f")).join("\u001e");
}

function serializeClipboardPlainText(values: readonly (readonly string[])[]): string {
  return values.map((row) => row.join("\t")).join("\n");
}

function parseClipboardPlainText(rawText: string): readonly (readonly string[])[] {
  if (rawText.length === 0) {
    return [];
  }
  return rawText
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((row) => row.split("\t"));
}

export function SheetGridView({
  engine,
  sheetName,
  variant = "playground",
  selectedAddr,
  editorValue,
  resolvedValue,
  isEditingCell,
  onSelect,
  onSelectionLabelChange,
  onBeginEdit,
  onEditorChange,
  onCommitEdit,
  onCancelEdit,
  onClearCell,
  onFillRange,
  onCopyRange,
  onPaste
}: SheetGridViewProps) {
  const editorRef = useRef<DataEditorRef | null>(null);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const wasEditingOverlayRef = useRef(false);
  const ignoreNextPointerSelectionRef = useRef(false);
  const pendingPointerCellRef = useRef<Item | null>(null);
  const dragAnchorCellRef = useRef<Item | null>(null);
  const dragPointerCellRef = useRef<Item | null>(null);
  const dragHeaderSelectionRef = useRef<HeaderSelection | null>(null);
  const dragViewportRef = useRef<VisibleRegionState | null>(null);
  const dragGeometryRef = useRef<PointerGeometry | null>(null);
  const dragDidMoveRef = useRef(false);
  const postDragSelectionExpiryRef = useRef<number>(0);
  const internalClipboardRef = useRef<InternalClipboardRange | null>(null);
  const activeSheetRef = useRef(sheetName);
  const columnResizeActiveRef = useRef(false);
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0
  });
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined);
  const [columnWidthsBySheet, setColumnWidthsBySheet] = useState<Record<string, Record<number, number>>>({});
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName]);
  const [gridSelection, setGridSelection] = useState<GridSelection>(() => createGridSelection(selectedCell.col, selectedCell.row));
  const gridMetrics = useMemo(() => getGridMetrics(variant), [variant]);
  const columnWidths = columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS;

  const columns = useMemo<readonly GridColumn[]>(
    () =>
      Array.from({ length: MAX_COLS }, (_, index) => ({
        id: indexToColumn(index),
        title: indexToColumn(index),
        width: getResolvedColumnWidth(columnWidths, index, gridMetrics.columnWidth)
      })),
    [columnWidths, gridMetrics.columnWidth]
  );

  const getCellContent = useCallback(
    ([col, row]: Item) => cellToGridCell(engine, sheetName, formatAddress(row, col)),
    [engine, sheetName]
  );

  const visibleItems = useMemo<Item[]>(() => {
    const items: Item[] = [];
    const rowEnd = Math.min(MAX_ROWS - 1, visibleRegion.range.y + visibleRegion.range.height - 1);
    const colEnd = Math.min(MAX_COLS - 1, visibleRegion.range.x + visibleRegion.range.width - 1);
    for (let row = visibleRegion.range.y; row <= rowEnd; row += 1) {
      for (let col = visibleRegion.range.x; col <= colEnd; col += 1) {
        items.push([col, row]);
      }
    }
    return items;
  }, [visibleRegion.range.height, visibleRegion.range.width, visibleRegion.range.x, visibleRegion.range.y]);

  const visibleAddresses = useMemo(
    () => visibleItems.map(([col, row]) => formatAddress(row, col)),
    [visibleItems]
  );
  const visibleDamage = useMemo(() => visibleItems.map((cell) => ({ cell })), [visibleItems]);

  useEffect(() => {
    return engine.subscribeCells(sheetName, visibleAddresses, () => {
      editorRef.current?.updateCells(visibleDamage);
    });
  }, [engine, sheetName, visibleAddresses, visibleDamage]);

  useLayoutEffect(() => {
    editorRef.current?.scrollTo(selectedCell.col, selectedCell.row);
  }, [selectedCell.col, selectedCell.row, sheetName]);

  useEffect(() => {
    const sheetChanged = activeSheetRef.current !== sheetName;
    activeSheetRef.current = sheetName;
    setGridSelection((current) => {
      const currentCell = current.current?.cell;
      if (!sheetChanged && currentCell && currentCell[0] === selectedCell.col && currentCell[1] === selectedCell.row) {
        return current;
      }
      pendingPointerCellRef.current = null;
      dragAnchorCellRef.current = null;
      dragPointerCellRef.current = null;
      dragGeometryRef.current = null;
      return createGridSelection(selectedCell.col, selectedCell.row);
    });
  }, [selectedCell.col, selectedCell.row, sheetName]);

  useEffect(() => {
    if (!isEditingCell) {
      setOverlayBounds(undefined);
      return;
    }

    let frame = 0;
    const tick = () => {
      const next = editorRef.current?.getBounds(selectedCell.col, selectedCell.row);
      setOverlayBounds((current) => {
        if (!next) {
          return current;
        }
        return sameBounds(current, next) ? current : next;
      });
      frame = window.requestAnimationFrame(tick);
    };
    frame = window.requestAnimationFrame(tick);
    return () => {
      window.cancelAnimationFrame(frame);
    };
  }, [isEditingCell, selectedCell.col, selectedCell.row, visibleRegion.tx, visibleRegion.ty]);

  useEffect(() => {
    if (wasEditingOverlayRef.current && !isEditingCell) {
      window.requestAnimationFrame(() => {
        hostRef.current?.focus();
      });
    }
    wasEditingOverlayRef.current = isEditingCell;
  }, [isEditingCell]);

  const beginSelectedEdit = useCallback(
    (seed?: string) => {
      onBeginEdit(seed ?? cellToEditorSeed(selectors.selectCellSnapshot(engine, sheetName, selectedAddr)));
    },
    [engine, onBeginEdit, selectedAddr, sheetName]
  );

  const beginEditAt = useCallback(
    (addr: string, seed?: string) => {
      onBeginEdit(seed ?? cellToEditorSeed(selectors.selectCellSnapshot(engine, sheetName, addr)));
    },
    [engine, onBeginEdit, sheetName]
  );

  const resolvePointerCell = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null
    ): Item | null => {
      const activeGeometry = geometry ?? (() => {
        const hostBounds = hostRef.current?.getBoundingClientRect();
        if (!hostBounds) {
          return null;
        }
        const firstCellBounds = editorRef.current?.getBounds(region.range.x, region.range.y);
        const cellWidth = firstCellBounds?.width ?? gridMetrics.columnWidth;
        const cellHeight = firstCellBounds?.height ?? gridMetrics.rowHeight;
        const dataLeft = firstCellBounds?.x ?? (hostBounds.left + gridMetrics.rowMarkerWidth);
        const dataTop = firstCellBounds?.y ?? (hostBounds.top + gridMetrics.headerHeight);
        const visibleColumnBounds = getVisibleColumnBounds(region, dataLeft, columnWidths, gridMetrics.columnWidth);
        const dataWidth = visibleColumnBounds.length === 0
          ? region.range.width * cellWidth
          : visibleColumnBounds.at(-1)!.right - dataLeft;
        return {
          hostBounds,
          cellWidth,
          cellHeight,
          dataLeft,
          dataTop,
          dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + dataWidth),
          dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + (region.range.height * cellHeight))
        };
      })();
      if (!activeGeometry) {
        return null;
      }

      const { hostBounds, cellHeight, dataLeft, dataTop, dataRight, dataBottom } = activeGeometry;
      if (!hostBounds) {
        return null;
      }

      if (clientX >= hostBounds.right - SCROLLBAR_GUTTER || clientY >= hostBounds.bottom - SCROLLBAR_GUTTER) {
        return null;
      }

      if (clientX < dataLeft || clientX >= dataRight || clientY < dataTop || clientY >= dataBottom) {
        return null;
      }

      const col = resolveColumnAtClientX(clientX, region, dataLeft, columnWidths, gridMetrics.columnWidth);
      const row = region.range.y + Math.floor((clientY - dataTop) / cellHeight);
      if (col === null || col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return null;
      }

      return [col, row];
    },
    [columnWidths, gridMetrics.columnWidth, gridMetrics.headerHeight, gridMetrics.rowHeight, gridMetrics.rowMarkerWidth, visibleRegion]
  );

  const resolvePointerGeometry = useCallback(
    (region: VisibleRegionState = visibleRegion): PointerGeometry | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      if (!hostBounds) {
        return null;
      }

      const firstCellBounds = editorRef.current?.getBounds(region.range.x, region.range.y);
      const cellWidth = firstCellBounds?.width ?? gridMetrics.columnWidth;
      const cellHeight = firstCellBounds?.height ?? gridMetrics.rowHeight;
      const dataLeft = firstCellBounds?.x ?? (hostBounds.left + gridMetrics.rowMarkerWidth);
      const dataTop = firstCellBounds?.y ?? (hostBounds.top + gridMetrics.headerHeight);
      const visibleColumnBounds = getVisibleColumnBounds(region, dataLeft, columnWidths, gridMetrics.columnWidth);
      const dataWidth = visibleColumnBounds.length === 0
        ? region.range.width * cellWidth
        : visibleColumnBounds.at(-1)!.right - dataLeft;
      return {
        hostBounds,
        cellWidth,
        cellHeight,
        dataLeft,
        dataTop,
        dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + dataWidth),
        dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + (region.range.height * cellHeight))
      };
    },
    [columnWidths, gridMetrics.columnWidth, gridMetrics.headerHeight, gridMetrics.rowHeight, gridMetrics.rowMarkerWidth, visibleRegion]
  );

  const resolveHeaderSelection = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region);
      if (!activeGeometry) {
        return null;
      }

      const { hostBounds, cellHeight, dataLeft, dataTop, dataRight, dataBottom } = activeGeometry;
      const headerBottom = dataTop;
      const rowAreaLeft = hostBounds.left;
      const rowAreaRight = dataLeft;

      if (clientY >= hostBounds.top && clientY < headerBottom && clientX >= dataLeft && clientX < dataRight) {
        const col = resolveColumnAtClientX(clientX, region, dataLeft, columnWidths, gridMetrics.columnWidth);
        if (col !== null && col >= 0 && col < MAX_COLS) {
          return { kind: "column", index: col };
        }
      }

      if (clientX >= rowAreaLeft && clientX < rowAreaRight && clientY >= dataTop && clientY < dataBottom) {
        const row = region.range.y + Math.floor((clientY - dataTop) / cellHeight);
        if (row >= 0 && row < MAX_ROWS) {
          return { kind: "row", index: row };
        }
      }

      return null;
    },
    [columnWidths, gridMetrics.columnWidth, resolvePointerGeometry, visibleRegion]
  );

  const resolveHeaderSelectionForDrag = useCallback(
    (
      kind: HeaderSelection["kind"],
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region);
      if (!activeGeometry) {
        return null;
      }

      const { hostBounds, cellHeight, dataLeft, dataTop, dataRight, dataBottom } = activeGeometry;

      if (kind === "column") {
        if (clientX < dataLeft || clientX >= dataRight || clientY < hostBounds.top || clientY >= dataBottom) {
          return null;
        }
        const col = resolveColumnAtClientX(clientX, region, dataLeft, columnWidths, gridMetrics.columnWidth);
        if (col === null || col < 0 || col >= MAX_COLS) {
          return null;
        }
        return { kind: "column", index: col };
      }

      if (clientY < dataTop || clientY >= dataBottom || clientX < hostBounds.left || clientX >= dataRight) {
        return null;
      }
      const row = region.range.y + Math.floor((clientY - dataTop) / cellHeight);
      if (row < 0 || row >= MAX_ROWS) {
        return null;
      }
      return { kind: "row", index: row };
    },
    [columnWidths, gridMetrics.columnWidth, resolvePointerGeometry, visibleRegion]
  );

  const selectionSummary = useMemo(
    () => formatSelectionSummary(gridSelection, selectedAddr),
    [gridSelection, selectedAddr]
  );

  const applyClipboardValues = useCallback(
    (target: Item, values: readonly (readonly string[])[]) => {
      if (values.length === 0 || values[0]?.length === 0) {
        return;
      }

      const internalClipboard = internalClipboardRef.current;
      const signature = serializeClipboardMatrix(values);
      if (
        internalClipboard
        && internalClipboard.signature === signature
        && internalClipboard.rowCount === values.length
        && internalClipboard.colCount === (values[0]?.length ?? 0)
      ) {
        onCopyRange(
          internalClipboard.sourceStartAddress,
          internalClipboard.sourceEndAddress,
          formatAddress(target[1], target[0]),
          formatAddress(
            target[1] + internalClipboard.rowCount - 1,
            target[0] + internalClipboard.colCount - 1
          )
        );
        return;
      }

      onPaste(formatAddress(target[1], target[0]), values);
    },
    [onCopyRange, onPaste]
  );

  const captureInternalClipboardSelection = useCallback(() => {
    const range = gridSelection.current?.range;
    if (!range || gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
      internalClipboardRef.current = null;
      return;
    }

    const values = Array.from({ length: range.height }, (_, rowOffset) =>
      Array.from({ length: range.width }, (_, colOffset) =>
        cellToEditorSeed(selectors.selectCellSnapshot(engine, sheetName, formatAddress(range.y + rowOffset, range.x + colOffset)))
      )
    );

    internalClipboardRef.current = {
      sourceStartAddress: formatAddress(range.y, range.x),
      sourceEndAddress: formatAddress(range.y + range.height - 1, range.x + range.width - 1),
      signature: serializeClipboardMatrix(values),
      plainText: serializeClipboardPlainText(values),
      rowCount: range.height,
      colCount: range.width
    };
  }, [engine, gridSelection, sheetName]);

  useEffect(() => {
    onSelectionLabelChange?.(selectionSummary);
  }, [onSelectionLabelChange, selectionSummary]);

  const handleGridKey = useCallback(
    (event: {
      key: string;
      ctrlKey: boolean;
      metaKey: boolean;
      altKey: boolean;
      shiftKey?: boolean;
      preventDefault(): void;
      cancel?: () => void;
    }) => {
      if (isEditingCell) {
        return;
      }

      if (event.key === "Enter" || event.key === "F2") {
        event.preventDefault();
        event.cancel?.();
        beginSelectedEdit();
        return;
      }

      if (event.key === "Tab") {
        event.preventDefault();
        event.cancel?.();
        onSelect(formatAddress(selectedCell.row, Math.min(MAX_COLS - 1, Math.max(0, selectedCell.col + (event.shiftKey ? -1 : 1)))));
        return;
      }

      if (isNavigationKey(event.key)) {
        event.preventDefault();
        event.cancel?.();

        const delta: Item =
          event.key === "ArrowUp"
            ? [0, -1]
            : event.key === "ArrowDown"
              ? [0, 1]
              : event.key === "ArrowLeft"
                ? [-1, 0]
                : [1, 0];
        const nextCell = clampCell([selectedCell.col + delta[0], selectedCell.row + delta[1]]);

        if (event.shiftKey) {
          setGridSelection((current) => {
            const anchor = current.current?.cell ?? [selectedCell.col, selectedCell.row];
            return createRangeSelection(createGridSelection(anchor[0], anchor[1]), anchor, nextCell);
          });
          return;
        }

        setGridSelection(createGridSelection(nextCell[0], nextCell[1]));
        onSelect(formatAddress(nextCell[1], nextCell[0]));
        return;
      }

      if (event.key === "Backspace" || event.key === "Delete") {
        event.preventDefault();
        event.cancel?.();
        onClearCell();
        return;
      }

      if ((event.ctrlKey || event.metaKey) && !event.altKey) {
        const normalizedKey = event.key.toLowerCase();
        if (normalizedKey === "c") {
          captureInternalClipboardSelection();
          const clipboard = internalClipboardRef.current;
          if (clipboard && typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
            event.preventDefault();
            event.cancel?.();
            void navigator.clipboard.writeText(clipboard.plainText).catch(() => {});
          }
          return;
        }

        if (normalizedKey === "x") {
          captureInternalClipboardSelection();
          return;
        }

        if (normalizedKey === "v" && typeof navigator !== "undefined" && navigator.clipboard?.readText) {
          event.preventDefault();
          event.cancel?.();
          const target: Item = gridSelection.current?.cell
            ? [...gridSelection.current.cell] as Item
            : [selectedCell.col, selectedCell.row];
          void navigator.clipboard.readText().then((rawText) => {
            const values = parseClipboardPlainText(rawText);
            applyClipboardValues(target, values);
          }).catch(() => {});
          return;
        }
      }

      if (
        event.key.length === 1
        && !event.ctrlKey
        && !event.metaKey
        && !event.altKey
      ) {
        event.preventDefault();
        event.cancel?.();
        beginSelectedEdit(event.key);
      }
    },
    [applyClipboardValues, beginSelectedEdit, captureInternalClipboardSelection, gridSelection.current?.cell, isEditingCell, onClearCell, onSelect, selectedCell.col, selectedCell.row]
  );

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      const activeElement = document.activeElement;
      if (
        activeElement instanceof HTMLInputElement
        || activeElement instanceof HTMLTextAreaElement
        || activeElement instanceof HTMLSelectElement
        || Boolean(activeElement && (activeElement as HTMLElement).isContentEditable)
      ) {
        return;
      }

      const withinGridHost = Boolean(activeElement && hostRef.current?.contains(activeElement));
      const onDocumentBody =
        activeElement === document.body || activeElement === document.documentElement || activeElement === null;
      if (withinGridHost) {
        return;
      }
      if (!onDocumentBody) {
        return;
      }

      if (
        !isPrintableKey({
          altKey: event.altKey,
          ctrlKey: event.ctrlKey,
          key: event.key,
          metaKey: event.metaKey
        } as GridKeyEventArgs)
        && !isClipboardShortcut({
          altKey: event.altKey,
          ctrlKey: event.ctrlKey,
          key: event.key,
          metaKey: event.metaKey
        })
        && !isNavigationKey(event.key)
        && event.key !== "Enter"
        && event.key !== "Tab"
        && event.key !== "F2"
        && event.key !== "Backspace"
        && event.key !== "Delete"
      ) {
        return;
      }

      handleGridKey({
        altKey: event.altKey,
        cancel: () => {
          event.stopPropagation();
        },
        ctrlKey: event.ctrlKey,
        key: event.key,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
        preventDefault: () => event.preventDefault()
      });
    };

    window.addEventListener("keydown", handleWindowKeyDown, true);
    return () => {
      window.removeEventListener("keydown", handleWindowKeyDown, true);
    };
  }, [handleGridKey]);

  const overlayStyle = useMemo(() => {
    if (!isEditingCell || !overlayBounds) {
      return undefined;
    }
    return {
      height: overlayBounds.height + 2,
      left: overlayBounds.x - 1,
      position: "fixed" as const,
      top: overlayBounds.y - 1,
      width: overlayBounds.width + 2
    };
  }, [isEditingCell, overlayBounds]);

  const gridTheme = useMemo(
    () => ({
      accentColor: "#1f7a43",
      accentFg: "#ffffff",
      bgCell: "#ffffff",
      bgCellMedium: "#f3f5f7",
      bgHeader: "#f6f7f8",
      borderColor: "#d5d9de",
      cellHorizontalPadding: variant === "product" ? 8 : 10,
      cellVerticalPadding: variant === "product" ? 4 : 6,
      drilldownBorder: "#d5d9de",
      editorFontSize: variant === "product" ? "12px" : "13px",
      fontFamily: '"Aptos","Segoe UI","IBM Plex Sans",sans-serif',
      headerFontStyle: variant === "product"
        ? "600 11px Aptos, Segoe UI, IBM Plex Sans, sans-serif"
        : "600 12px Aptos, Segoe UI, IBM Plex Sans, sans-serif",
      horizontalBorderColor: "#e5e7eb",
      lineHeight: variant === "product" ? 1.2 : 1.3,
      textDark: "#101828",
      textHeader: "#344054",
      textLight: "#667085"
    }),
    [variant]
  );

  const handleFillPattern = useCallback(
    (event: FillPatternEventArgs) => {
      const source = rectangleToAddresses(event.patternSource);
      const target = rectangleToAddresses(event.fillDestination);
      if (
        source.startAddress === target.startAddress
        && source.endAddress === target.endAddress
      ) {
        return;
      }
      event.preventDefault();
      onFillRange(source.startAddress, source.endAddress, target.startAddress, target.endAddress);
    },
    [onFillRange]
  );

  const applyColumnWidth = useCallback(
    (columnIndex: number, newSize: number) => {
      const clampedSize = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(newSize)));
      setColumnWidthsBySheet((current) => {
        const nextSheetWidths = current[sheetName] ?? EMPTY_COLUMN_WIDTHS;
        if (nextSheetWidths[columnIndex] === clampedSize) {
          return current;
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetWidths,
            [columnIndex]: clampedSize
          }
        };
      });
    },
    [sheetName]
  );

  const computeAutofitColumnWidth = useCallback(
    (columnIndex: number): number => {
      const canvas = textMeasureCanvasRef.current ?? document.createElement("canvas");
      textMeasureCanvasRef.current = canvas;
      const context = canvas.getContext("2d");
      if (!context) {
        return gridMetrics.columnWidth;
      }

      const cellFont = `400 ${gridTheme.editorFontSize} ${gridTheme.fontFamily}`;
      const headerFont = gridTheme.headerFontStyle;
      let measuredWidth = 0;

      context.font = headerFont;
      measuredWidth = Math.max(measuredWidth, context.measureText(indexToColumn(columnIndex)).width);

      const sheet = engine.workbook.getSheet(sheetName);
      sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (col !== columnIndex) {
          return;
        }
        const cell = cellToGridCell(engine, sheetName, formatAddress(row, col));
        const displayText = "displayData" in cell ? String(cell.displayData ?? "") : "copyData" in cell ? String(cell.copyData ?? "") : "";
        context.font = cellFont;
        measuredWidth = Math.max(measuredWidth, context.measureText(displayText).width);
      });

      return Math.max(
        MIN_COLUMN_WIDTH,
        Math.min(MAX_COLUMN_WIDTH, Math.ceil(measuredWidth + (variant === "product" ? 28 : 32)))
      );
    },
    [engine, gridMetrics.columnWidth, gridTheme.editorFontSize, gridTheme.fontFamily, gridTheme.headerFontStyle, sheetName, variant]
  );

  return (
    <div className="sheet-grid-shell" data-testid="sheet-grid-shell">
      {variant === "playground" ? (
        <div className="sheet-grid-banner">
          <div>
            <p className="panel-eyebrow">Surface</p>
            <strong>
              {MAX_ROWS.toLocaleString()} rows x {MAX_COLS.toLocaleString()} columns
            </strong>
          </div>
          <div className="viewport-meta">
            <span data-testid="selection-chip">
              {sheetName}!{selectionSummary}
            </span>
            <span>{resolvedValue || "∅"}</span>
          </div>
        </div>
      ) : null}
      <div
        className="sheet-grid-host"
        data-column-width-overrides={JSON.stringify(columnWidths)}
        data-default-column-width={String(gridMetrics.columnWidth)}
        data-testid="sheet-grid"
        onKeyDownCapture={() => {
          ignoreNextPointerSelectionRef.current = false;
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragHeaderSelectionRef.current = null;
          dragGeometryRef.current = null;
          dragDidMoveRef.current = false;
          dragViewportRef.current = null;
          postDragSelectionExpiryRef.current = 0;
        }}
        onCopyCapture={(event) => {
          captureInternalClipboardSelection();
          if (!event.clipboardData) {
            return;
          }
          const clipboard = internalClipboardRef.current;
          if (!clipboard) {
            return;
          }
          event.clipboardData.setData("text/plain", clipboard.plainText);
          event.preventDefault();
        }}
        onPasteCapture={(event) => {
          const rawText = event.clipboardData?.getData("text/plain") ?? "";
          const values = parseClipboardPlainText(rawText);
          if (values.length === 0 || values[0]?.length === 0) {
            return;
          }

          const target = gridSelection.current?.cell ?? [selectedCell.col, selectedCell.row];
          applyClipboardValues(target, values);
          event.preventDefault();
          event.stopPropagation();
        }}
        onKeyDown={(event) => handleGridKey(event)}
        onDoubleClickCapture={(event) => {
          if (variant !== "product") {
            return;
          }
          const activeGeometry = resolvePointerGeometry(visibleRegion);
          if (!activeGeometry) {
            return;
          }
          const resizeTarget = resolveColumnResizeTarget(
            event.clientX,
            event.clientY,
            visibleRegion,
            activeGeometry,
            columnWidths,
            gridMetrics.columnWidth
          );
          if (resizeTarget === null) {
            return;
          }

          event.preventDefault();
          event.stopPropagation();
          columnResizeActiveRef.current = false;
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragHeaderSelectionRef.current = null;
          dragGeometryRef.current = null;
          dragDidMoveRef.current = false;
          dragViewportRef.current = null;
          postDragSelectionExpiryRef.current = 0;
          applyColumnWidth(resizeTarget, computeAutofitColumnWidth(resizeTarget));
        }}
        onPointerMoveCapture={(event) => {
          if (columnResizeActiveRef.current) {
            return;
          }
          if ((event.buttons & 1) !== 1) {
            return;
          }
          const headerAnchor = dragHeaderSelectionRef.current;
          if (headerAnchor) {
            const nextHeader = resolveHeaderSelectionForDrag(
              headerAnchor.kind,
              event.clientX,
              event.clientY,
              dragViewportRef.current ?? visibleRegion,
              dragGeometryRef.current
            );
            if (!nextHeader || nextHeader.index === headerAnchor.index) {
              return;
            }
            dragPointerCellRef.current = null;
            dragDidMoveRef.current = true;
            setGridSelection(
              headerAnchor.kind === "column"
                ? createColumnSliceSelection(headerAnchor.index, nextHeader.index, selectedCell.row)
                : createRowSliceSelection(selectedCell.col, headerAnchor.index, nextHeader.index)
            );
            return;
          }
          if (dragAnchorCellRef.current === null) {
            return;
          }
          const pointerCell = resolvePointerCell(
            event.clientX,
            event.clientY,
            dragViewportRef.current ?? visibleRegion,
            dragGeometryRef.current
          );
          if (!pointerCell) {
            return;
          }
          const currentPointer = dragPointerCellRef.current;
          if (currentPointer && currentPointer[0] === pointerCell[0] && currentPointer[1] === pointerCell[1]) {
            return;
          }
          dragPointerCellRef.current = pointerCell;
          if (pointerCell[0] !== dragAnchorCellRef.current[0] || pointerCell[1] !== dragAnchorCellRef.current[1]) {
            dragDidMoveRef.current = true;
            ignoreNextPointerSelectionRef.current = false;
            setGridSelection(
              createRangeSelection(
                createGridSelection(dragAnchorCellRef.current[0], dragAnchorCellRef.current[1]),
                dragAnchorCellRef.current,
                pointerCell
              )
            );
          }
        }}
        onPointerDownCapture={(event) => {
          if (event.button !== 0) {
            return;
          }
          const activeGeometry = resolvePointerGeometry(visibleRegion);
          if (
            variant === "product"
            && activeGeometry
            && resolveColumnResizeTarget(
              event.clientX,
              event.clientY,
              visibleRegion,
              activeGeometry,
              columnWidths,
              gridMetrics.columnWidth
            ) !== null
          ) {
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragHeaderSelectionRef.current = null;
            dragGeometryRef.current = null;
            dragDidMoveRef.current = false;
            dragViewportRef.current = null;
            postDragSelectionExpiryRef.current = 0;
            return;
          }
          const headerSelection = resolveHeaderSelection(event.clientX, event.clientY);
          if (headerSelection) {
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragHeaderSelectionRef.current = headerSelection;
            dragGeometryRef.current = null;
            dragDidMoveRef.current = false;
            dragViewportRef.current = null;
            postDragSelectionExpiryRef.current = 0;
            if (isEditingCell) {
              onCommitEdit();
            }
            dragGeometryRef.current = activeGeometry;
            dragViewportRef.current = visibleRegion;
            if (headerSelection.kind === "row") {
              ignoreNextPointerSelectionRef.current = true;
              setGridSelection(createRowSliceSelection(selectedCell.col, headerSelection.index, headerSelection.index));
              onSelect(formatAddress(headerSelection.index, selectedCell.col));
              hostRef.current?.focus();
              window.requestAnimationFrame(() => {
                ignoreNextPointerSelectionRef.current = false;
              });
              return;
            }
            ignoreNextPointerSelectionRef.current = true;
            setGridSelection(createColumnSliceSelection(headerSelection.index, headerSelection.index, selectedCell.row));
            onSelect(formatAddress(selectedCell.row, headerSelection.index));
            hostRef.current?.focus();
            window.requestAnimationFrame(() => {
              ignoreNextPointerSelectionRef.current = false;
            });
            return;
          }
          const pointerCell = resolvePointerCell(event.clientX, event.clientY);
          ignoreNextPointerSelectionRef.current = pointerCell === null;
          pendingPointerCellRef.current = pointerCell;
          dragAnchorCellRef.current = pointerCell;
          dragPointerCellRef.current = pointerCell;
          dragGeometryRef.current = activeGeometry;
          dragDidMoveRef.current = false;
          dragViewportRef.current = visibleRegion;
          if (pointerCell) {
            ignoreNextPointerSelectionRef.current = true;
            setGridSelection(createGridSelection(pointerCell[0], pointerCell[1]));
            if (isEditingCell) {
              onCommitEdit();
            }
            onSelect(formatAddress(pointerCell[1], pointerCell[0]));
          }
          hostRef.current?.focus();
        }}
        onPointerUpCapture={(event) => {
          if (columnResizeActiveRef.current) {
            return;
          }
          const headerAnchor = dragHeaderSelectionRef.current;
          if (headerAnchor) {
            const finalHeader = resolveHeaderSelectionForDrag(
              headerAnchor.kind,
              event.clientX,
              event.clientY,
              dragViewportRef.current ?? visibleRegion,
              dragGeometryRef.current
            ) ?? headerAnchor;
            setGridSelection(
              headerAnchor.kind === "column"
                ? createColumnSliceSelection(headerAnchor.index, finalHeader.index, selectedCell.row)
                : createRowSliceSelection(selectedCell.col, headerAnchor.index, finalHeader.index)
            );
            onSelect(
              headerAnchor.kind === "column"
                ? formatAddress(selectedCell.row, headerAnchor.index)
                : formatAddress(headerAnchor.index, selectedCell.col)
            );
            window.requestAnimationFrame(() => {
              pendingPointerCellRef.current = null;
              dragAnchorCellRef.current = null;
              dragPointerCellRef.current = null;
              dragHeaderSelectionRef.current = null;
              dragGeometryRef.current = null;
              dragDidMoveRef.current = false;
              dragViewportRef.current = null;
            });
            return;
          }
          const anchorCell = dragAnchorCellRef.current;
          if (!anchorCell) {
            return;
          }

          if (dragDidMoveRef.current) {
            const pointerCell = resolvePointerCell(
                event.clientX,
                event.clientY,
                dragViewportRef.current ?? visibleRegion,
                dragGeometryRef.current
              )
              ?? dragPointerCellRef.current
              ?? anchorCell;

            const finalSelection = createRangeSelection(
                createGridSelection(anchorCell[0], anchorCell[1]),
                anchorCell,
                pointerCell
              );
            postDragSelectionExpiryRef.current = window.performance.now() + 200;
            setGridSelection(finalSelection);
            onSelect(formatAddress(anchorCell[1], anchorCell[0]));
          }

          window.requestAnimationFrame(() => {
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragHeaderSelectionRef.current = null;
            dragGeometryRef.current = null;
            dragDidMoveRef.current = false;
            dragViewportRef.current = null;
          });
        }}
        ref={hostRef}
        tabIndex={0}
      >
        <DataEditor
          ref={editorRef}
          cellActivationBehavior="double-click"
          className="glide-sheet-grid"
          columns={columns}
          drawFocusRing={false}
          fillHandle={variant === "product"}
          freezeColumns={0}
          getCellContent={getCellContent}
          getCellsForSelection={true}
          gridSelection={gridSelection}
          headerHeight={gridMetrics.headerHeight}
          height="100%"
          {...(variant === "product"
            ? {
                onColumnResizeStart: () => {
                  columnResizeActiveRef.current = true;
                  pendingPointerCellRef.current = null;
                  dragAnchorCellRef.current = null;
                  dragPointerCellRef.current = null;
                  dragHeaderSelectionRef.current = null;
                  dragGeometryRef.current = null;
                  dragDidMoveRef.current = false;
                  dragViewportRef.current = null;
                  postDragSelectionExpiryRef.current = 0;
                },
                onColumnResize: (_column: GridColumn, newSize: number, columnIndex: number) => {
                  applyColumnWidth(columnIndex, newSize);
                },
                onColumnResizeEnd: (_column: GridColumn, newSize: number, columnIndex: number) => {
                  applyColumnWidth(columnIndex, newSize);
                  window.requestAnimationFrame(() => {
                    columnResizeActiveRef.current = false;
                  });
                },
                maxColumnWidth: MAX_COLUMN_WIDTH,
                minColumnWidth: MIN_COLUMN_WIDTH
              }
            : {})}
          onCellActivated={([col, row]) => {
            const cell = dragAnchorCellRef.current ?? pendingPointerCellRef.current ?? clampCell([col, row]);
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            setGridSelection(createGridSelection(cell[0], cell[1]));
            const addr = formatAddress(cell[1], cell[0]);
            onSelect(addr);
            beginEditAt(addr);
          }}
          onDelete={() => {
            onClearCell();
            return false;
          }}
          onFillPattern={handleFillPattern}
          onHeaderClicked={(col, event) => {
            if (variant === "product" && event.isEdge && event.isDoubleClick) {
              applyColumnWidth(col, computeAutofitColumnWidth(col));
              columnResizeActiveRef.current = false;
              return;
            }
            if (columnResizeActiveRef.current) {
              return;
            }
            ignoreNextPointerSelectionRef.current = true;
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragHeaderSelectionRef.current = null;
            dragGeometryRef.current = null;
            dragViewportRef.current = null;
            postDragSelectionExpiryRef.current = 0;
            if (isEditingCell) {
              onCommitEdit();
            }
            setGridSelection(createColumnSelection(col, selectedCell.row));
            onSelect(formatAddress(selectedCell.row, col));
            hostRef.current?.focus();
            window.requestAnimationFrame(() => {
              ignoreNextPointerSelectionRef.current = false;
            });
          }}
          onGridSelectionChange={(nextSelection) => {
            if (columnResizeActiveRef.current) {
              return;
            }
            if (postDragSelectionExpiryRef.current > 0) {
              if (window.performance.now() <= postDragSelectionExpiryRef.current) {
                postDragSelectionExpiryRef.current = 0;
                return;
              }
              postDragSelectionExpiryRef.current = 0;
            }
            if (ignoreNextPointerSelectionRef.current) {
              ignoreNextPointerSelectionRef.current = false;
              return;
            }
            if (dragViewportRef.current) {
              return;
            }
            if (nextSelection.columns.length > 0 || nextSelection.rows.length > 0) {
              setGridSelection(nextSelection);
              const nextColumn = nextSelection.columns.first();
              if (nextColumn !== undefined) {
                onSelect(formatAddress(selectedCell.row, nextColumn));
                return;
              }
              const nextRow = nextSelection.rows.first();
              if (nextRow !== undefined) {
                onSelect(formatAddress(nextRow, selectedCell.col));
                return;
              }
            }
            const nextCell = nextSelection.current?.cell;
            if (!nextCell) {
              return;
            }
            const anchorCell = dragAnchorCellRef.current ?? pendingPointerCellRef.current;
            const pointerCell = dragPointerCellRef.current ?? pendingPointerCellRef.current;
            const correctedSelection = anchorCell && pointerCell
              ? createRangeSelection(nextSelection, anchorCell, pointerCell)
              : {
                  ...nextSelection,
                  current: nextSelection.current
                    ? {
                        ...nextSelection.current,
                        cell: clampCell(nextSelection.current.cell),
                        range: clampSelectionRange(nextSelection.current.range)
                      }
                    : nextSelection.current
                };
            const cell = correctedSelection.current?.cell ?? nextCell;
            setGridSelection(correctedSelection);
            if (isEditingCell) {
              onCommitEdit();
            }
            onSelect(formatAddress(cell[1], cell[0]));
          }}
          onKeyDown={(event) => {
            if (!isPrintableKey(event) && !isClipboardShortcut(event) && !isNavigationKey(event.key) && event.key !== "Enter" && event.key !== "Tab" && event.key !== "F2" && event.key !== "Backspace" && event.key !== "Delete") {
              return;
            }
            handleGridKey(event);
          }}
          onPaste={(target, values) => {
            applyClipboardValues(target, values);
            return false;
          }}
          onVisibleRegionChanged={(range, tx, ty) => {
            setVisibleRegion({ range, tx, ty });
          }}
          rowHeight={gridMetrics.rowHeight}
          columnSelect="multi"
          columnSelectionBlending="additive"
          columnSelectionMode="multi"
          rowMarkers={{ kind: "clickable-number", width: gridMetrics.rowMarkerWidth }}
          rowSelect="multi"
          rowSelectionBlending="additive"
          rowSelectionMode="multi"
          rows={MAX_ROWS}
          smoothScrollX={true}
          smoothScrollY={false}
          theme={gridTheme}
          trapFocus={false}
          verticalBorder={true}
          width="100%"
        />
      </div>
      {isEditingCell && overlayStyle ? (
        <CellEditorOverlay
          label={`${sheetName}!${selectedAddr}`}
          onCancel={onCancelEdit}
          onChange={onEditorChange}
          onCommit={onCommitEdit}
          resolvedValue={resolvedValue}
          value={editorValue}
          style={overlayStyle}
        />
      ) : null}
    </div>
  );
}
