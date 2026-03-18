import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { MAX_COLS, MAX_ROWS, ValueTag, formatErrorCode } from "@bilig/protocol";
import {
  CompactSelection,
  DataEditor,
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

type HeaderSelection =
  | { kind: "column"; index: number }
  | { kind: "row"; index: number };

const DEFAULT_COLUMN_WIDTH = 120;
const DEFAULT_ROW_HEIGHT = 28;
const HEADER_HEIGHT = 30;
const ROW_MARKER_WIDTH = 60;
const SCROLLBAR_GUTTER = 18;

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
  onPaste
}: SheetGridViewProps) {
  const editorRef = useRef<DataEditorRef | null>(null);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const wasEditingOverlayRef = useRef(false);
  const ignoreNextPointerSelectionRef = useRef(false);
  const pendingPointerCellRef = useRef<Item | null>(null);
  const dragAnchorCellRef = useRef<Item | null>(null);
  const dragPointerCellRef = useRef<Item | null>(null);
  const dragViewportRef = useRef<VisibleRegionState | null>(null);
  const dragGeometryRef = useRef<PointerGeometry | null>(null);
  const dragDidMoveRef = useRef(false);
  const postDragSelectionExpiryRef = useRef<number>(0);
  const activeSheetRef = useRef(sheetName);
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0
  });
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined);
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName]);
  const [gridSelection, setGridSelection] = useState<GridSelection>(() => createGridSelection(selectedCell.col, selectedCell.row));

  const columns = useMemo<readonly GridColumn[]>(
    () =>
      Array.from({ length: MAX_COLS }, (_, index) => ({
        id: indexToColumn(index),
        title: indexToColumn(index),
        width: DEFAULT_COLUMN_WIDTH
      })),
    []
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
        const cellWidth = firstCellBounds?.width ?? DEFAULT_COLUMN_WIDTH;
        const cellHeight = firstCellBounds?.height ?? DEFAULT_ROW_HEIGHT;
        const dataLeft = firstCellBounds?.x ?? (hostBounds.left + ROW_MARKER_WIDTH);
        const dataTop = firstCellBounds?.y ?? (hostBounds.top + HEADER_HEIGHT);
        return {
          hostBounds,
          cellWidth,
          cellHeight,
          dataLeft,
          dataTop,
          dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + (region.range.width * cellWidth)),
          dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + (region.range.height * cellHeight))
        };
      })();
      if (!activeGeometry) {
        return null;
      }

      const { hostBounds, cellWidth, cellHeight, dataLeft, dataTop, dataRight, dataBottom } = activeGeometry;
      if (!hostBounds) {
        return null;
      }

      if (clientX >= hostBounds.right - SCROLLBAR_GUTTER || clientY >= hostBounds.bottom - SCROLLBAR_GUTTER) {
        return null;
      }

      if (clientX < dataLeft || clientX >= dataRight || clientY < dataTop || clientY >= dataBottom) {
        return null;
      }

      const col = region.range.x + Math.floor((clientX - dataLeft) / cellWidth);
      const row = region.range.y + Math.floor((clientY - dataTop) / cellHeight);
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return null;
      }

      return [col, row];
    },
    [visibleRegion]
  );

  const resolvePointerGeometry = useCallback(
    (region: VisibleRegionState = visibleRegion): PointerGeometry | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      if (!hostBounds) {
        return null;
      }

      const firstCellBounds = editorRef.current?.getBounds(region.range.x, region.range.y);
      const cellWidth = firstCellBounds?.width ?? DEFAULT_COLUMN_WIDTH;
      const cellHeight = firstCellBounds?.height ?? DEFAULT_ROW_HEIGHT;
      const dataLeft = firstCellBounds?.x ?? (hostBounds.left + ROW_MARKER_WIDTH);
      const dataTop = firstCellBounds?.y ?? (hostBounds.top + HEADER_HEIGHT);
      return {
        hostBounds,
        cellWidth,
        cellHeight,
        dataLeft,
        dataTop,
        dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + (region.range.width * cellWidth)),
        dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + (region.range.height * cellHeight))
      };
    },
    [visibleRegion]
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

      const { hostBounds, cellWidth, cellHeight, dataLeft, dataTop, dataRight, dataBottom } = activeGeometry;
      const headerBottom = dataTop;
      const rowAreaLeft = hostBounds.left;
      const rowAreaRight = dataLeft;

      if (clientY >= hostBounds.top && clientY < headerBottom && clientX >= dataLeft && clientX < dataRight) {
        const col = region.range.x + Math.floor((clientX - dataLeft) / cellWidth);
        if (col >= 0 && col < MAX_COLS) {
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
    [resolvePointerGeometry, visibleRegion]
  );

  const selectionSummary = useMemo(
    () => formatSelectionSummary(gridSelection, selectedAddr),
    [gridSelection, selectedAddr]
  );

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
    [beginSelectedEdit, isEditingCell, onClearCell, onSelect, selectedCell.col, selectedCell.row]
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
        data-testid="sheet-grid"
        onKeyDownCapture={() => {
          ignoreNextPointerSelectionRef.current = false;
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragGeometryRef.current = null;
          dragDidMoveRef.current = false;
          dragViewportRef.current = null;
          postDragSelectionExpiryRef.current = 0;
        }}
        onKeyDown={(event) => handleGridKey(event)}
        onPointerMoveCapture={(event) => {
          if ((event.buttons & 1) !== 1 || dragAnchorCellRef.current === null) {
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
          const headerSelection = resolveHeaderSelection(event.clientX, event.clientY);
          if (headerSelection) {
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragGeometryRef.current = null;
            dragDidMoveRef.current = false;
            dragViewportRef.current = null;
            postDragSelectionExpiryRef.current = 0;
            if (isEditingCell) {
              onCommitEdit();
            }
            if (headerSelection.kind === "row") {
              ignoreNextPointerSelectionRef.current = true;
              setGridSelection(createRowSelection(selectedCell.col, headerSelection.index));
              onSelect(formatAddress(headerSelection.index, selectedCell.col));
              hostRef.current?.focus();
              window.requestAnimationFrame(() => {
                ignoreNextPointerSelectionRef.current = false;
              });
            }
            return;
          }
          const pointerCell = resolvePointerCell(event.clientX, event.clientY);
          ignoreNextPointerSelectionRef.current = pointerCell === null;
          pendingPointerCellRef.current = pointerCell;
          dragAnchorCellRef.current = pointerCell;
          dragPointerCellRef.current = pointerCell;
          dragGeometryRef.current = resolvePointerGeometry(visibleRegion);
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
          freezeColumns={0}
          getCellContent={getCellContent}
          getCellsForSelection={true}
          gridSelection={gridSelection}
          headerHeight={HEADER_HEIGHT}
          height="100%"
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
          onHeaderClicked={(col) => {
            ignoreNextPointerSelectionRef.current = true;
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
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
            if (!isPrintableKey(event) && !isNavigationKey(event.key) && event.key !== "Enter" && event.key !== "Tab" && event.key !== "F2" && event.key !== "Backspace" && event.key !== "Delete") {
              return;
            }
            handleGridKey(event);
          }}
          onPaste={(target, values) => {
            onPaste(formatAddress(target[1], target[0]), values);
            return false;
          }}
          onVisibleRegionChanged={(range, tx, ty) => {
            setVisibleRegion({ range, tx, ty });
          }}
          rowHeight={DEFAULT_ROW_HEIGHT}
          columnSelect="multi"
          columnSelectionBlending="additive"
          columnSelectionMode="multi"
          rowMarkers={{ kind: "clickable-number", width: ROW_MARKER_WIDTH }}
          rowSelect="multi"
          rowSelectionBlending="additive"
          rowSelectionMode="multi"
          rows={MAX_ROWS}
          smoothScrollX={true}
          smoothScrollY={false}
          theme={{
            accentColor: "#1f7a43",
            accentFg: "#ffffff",
            bgCell: "#ffffff",
            bgCellMedium: "#f3f5f7",
            bgHeader: "#f6f7f8",
            borderColor: "#d5d9de",
            cellHorizontalPadding: 10,
            cellVerticalPadding: 6,
            drilldownBorder: "#d5d9de",
            editorFontSize: "13px",
            fontFamily: '"Aptos","Segoe UI","IBM Plex Sans",sans-serif',
            headerFontStyle: "600 12px Aptos, Segoe UI, IBM Plex Sans, sans-serif",
            horizontalBorderColor: "#e5e7eb",
            lineHeight: 1.3,
            textDark: "#101828",
            textHeader: "#344054",
            textLight: "#667085"
          }}
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
