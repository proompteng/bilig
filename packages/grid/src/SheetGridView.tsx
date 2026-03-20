import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { MAX_COLS, MAX_ROWS, ValueTag, formatErrorCode } from "@bilig/protocol";
import {
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
import {
  COLUMN_RESIZE_HANDLE_THRESHOLD,
  EMPTY_COLUMN_WIDTHS,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  getGridMetrics,
  getResolvedColumnWidth,
  type GridRect
} from "./gridMetrics.js";
import {
  createColumnSelection,
  createColumnSliceSelection,
  createGridSelection,
  createRangeSelection,
  createRowSliceSelection,
  clampCell,
  clampSelectionRange,
  formatSelectionSummary,
  rectangleToAddresses,
  sameItem
} from "./gridSelection.js";
import {
  createPointerGeometry,
  resolveColumnResizeTarget,
  resolveHeaderSelection as resolveHeaderSelectionFromGeometry,
  resolveHeaderSelectionForDrag as resolveHeaderSelectionForDragFromGeometry,
  resolvePointerCell as resolvePointerCellFromGeometry,
  type HeaderSelection,
  type PointerGeometry,
  type VisibleRegionState
} from "./gridPointer.js";

export type EditMovement = readonly [-1 | 0 | 1, -1 | 0 | 1];
export type EditSelectionBehavior = "select-all" | "caret-end";

interface SheetGridViewProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  variant?: "playground" | "product";
  selectedAddr: string;
  editorValue: string;
  editorSelectionBehavior: EditSelectionBehavior;
  resolvedValue: string;
  isEditingCell: boolean;
  onSelect(addr: string): void;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
  onBeginEdit(seed?: string, selectionBehavior?: EditSelectionBehavior): void;
  onEditorChange(next: string): void;
  onCommitEdit(movement?: EditMovement): void;
  onCancelEdit(): void;
  onClearCell(): void;
  onFillRange(sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onCopyRange(sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onPaste(addr: string, values: readonly (readonly string[])[]): void;
}

interface InternalClipboardRange {
  sourceStartAddress: string;
  sourceEndAddress: string;
  signature: string;
  plainText: string;
  rowCount: number;
  colCount: number;
}

function isPrintableKey(event: GridKeyEventArgs): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return false;
  }
  return event.key.length === 1;
}

function normalizeKeyboardKey(key: string, code?: string): string {
  if (code?.startsWith("Numpad")) {
    const suffix = code.slice("Numpad".length);
    if (/^\d$/.test(suffix)) {
      return suffix;
    }
    if (suffix === "Decimal") {
      return ".";
    }
    if (suffix === "Add") {
      return "+";
    }
    if (suffix === "Subtract") {
      return "-";
    }
    if (suffix === "Multiply") {
      return "*";
    }
    if (suffix === "Divide") {
      return "/";
    }
  }
  return key;
}

function isNumericEditorSeed(value: string): boolean {
  const normalized = value.trim();
  if (normalized.length === 0 || normalized.startsWith("=")) {
    return false;
  }
  return /^-?\d+(\.\d+)?$/.test(normalized);
}

function isCellEditorInputFocused(): boolean {
  if (typeof document === "undefined") {
    return false;
  }
  const activeElement = document.activeElement;
  return activeElement instanceof HTMLInputElement && activeElement.dataset.testid === "cell-editor-input";
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
  editorSelectionBehavior,
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
  const lastBodyClickCellRef = useRef<Item | null>(null);
  const internalClipboardRef = useRef<InternalClipboardRange | null>(null);
  const activeSheetRef = useRef(sheetName);
  const columnResizeActiveRef = useRef(false);
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const pendingTypeSeedRef = useRef<string | null>(null);
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

  const focusGrid = useCallback(() => {
    editorRef.current?.focus();
  }, []);

  useEffect(() => {
    if (wasEditingOverlayRef.current && !isEditingCell) {
      window.requestAnimationFrame(() => {
        focusGrid();
      });
    }
    if (isEditingCell) {
      pendingTypeSeedRef.current = null;
    }
    wasEditingOverlayRef.current = isEditingCell;
  }, [focusGrid, isEditingCell]);

  const beginSelectedEdit = useCallback(
    (seed?: string, selectionBehavior: EditSelectionBehavior = "caret-end") => {
      onBeginEdit(
        seed ?? cellToEditorSeed(selectors.selectCellSnapshot(engine, sheetName, selectedAddr)),
        selectionBehavior
      );
    },
    [engine, onBeginEdit, selectedAddr, sheetName]
  );

  const beginEditAt = useCallback(
    (addr: string, seed?: string, selectionBehavior: EditSelectionBehavior = "caret-end") => {
      onBeginEdit(
        seed ?? cellToEditorSeed(selectors.selectCellSnapshot(engine, sheetName, addr)),
        selectionBehavior
      );
    },
    [engine, onBeginEdit, sheetName]
  );

  const resolvePointerGeometry = useCallback(
    (region: VisibleRegionState = visibleRegion): PointerGeometry | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      if (!hostBounds) {
        return null;
      }

      const rect: GridRect = {
        left: hostBounds.left,
        top: hostBounds.top,
        right: hostBounds.right,
        bottom: hostBounds.bottom
      };
      return createPointerGeometry(rect, region, columnWidths, gridMetrics);
    },
    [columnWidths, gridMetrics, visibleRegion]
  );

  const resolvePointerCell = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null
    ): Item | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region);
      if (!activeGeometry) {
        return null;
      }
      return resolvePointerCellFromGeometry({
        clientX,
        clientY,
        region,
        geometry: activeGeometry,
        columnWidths,
        gridMetrics,
        selectedCell: [selectedCell.col, selectedCell.row],
        selectedCellBounds: editorRef.current?.getBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange: gridSelection.current?.range ?? null,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0
      });
    },
    [columnWidths, gridMetrics, gridSelection, resolvePointerGeometry, selectedCell.col, selectedCell.row, visibleRegion]
  );

  const resolveHeaderSelectionAtPointer = useCallback(
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
      return resolveHeaderSelectionFromGeometry(clientX, clientY, region, activeGeometry, columnWidths, gridMetrics);
    },
    [columnWidths, gridMetrics, resolvePointerGeometry, visibleRegion]
  );

  const resolveHeaderSelectionForPointerDrag = useCallback(
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
      return resolveHeaderSelectionForDragFromGeometry(kind, clientX, clientY, region, activeGeometry, columnWidths, gridMetrics);
    },
    [columnWidths, gridMetrics, resolvePointerGeometry, visibleRegion]
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
        if (
          event.key.length === 1
          && !event.ctrlKey
          && !event.metaKey
          && !event.altKey
          && !isCellEditorInputFocused()
        ) {
          const nextValue = `${editorValue}${event.key}`;
          event.preventDefault();
          event.cancel?.();
          onEditorChange(nextValue);
        }
        return;
      }

      if (event.key === "F2") {
        pendingTypeSeedRef.current = null;
        event.preventDefault();
        event.cancel?.();
        beginSelectedEdit(undefined, "caret-end");
        return;
      }

      if (event.key === "Enter") {
        event.preventDefault();
        event.cancel?.();
        onSelect(
          formatAddress(
            Math.min(MAX_ROWS - 1, Math.max(0, selectedCell.row + (event.shiftKey ? -1 : 1))),
            selectedCell.col
          )
        );
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
        pendingTypeSeedRef.current = null;
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
        const seededValue = `${pendingTypeSeedRef.current ?? ""}${event.key}`;
        pendingTypeSeedRef.current = seededValue;
        event.preventDefault();
        event.cancel?.();
        beginSelectedEdit(seededValue, "caret-end");
      }
    },
    [applyClipboardValues, beginSelectedEdit, captureInternalClipboardSelection, editorValue, gridSelection.current?.cell, isEditingCell, onClearCell, onEditorChange, onSelect, selectedCell.col, selectedCell.row]
  );

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      const normalizedKey = normalizeKeyboardKey(event.key, event.code);
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
          key: normalizedKey,
          metaKey: event.metaKey
        } as GridKeyEventArgs)
        && !isClipboardShortcut({
          altKey: event.altKey,
          ctrlKey: event.ctrlKey,
          key: normalizedKey,
          metaKey: event.metaKey
        })
        && !isNavigationKey(normalizedKey)
        && normalizedKey !== "Enter"
        && normalizedKey !== "Tab"
        && normalizedKey !== "F2"
        && normalizedKey !== "Backspace"
        && normalizedKey !== "Delete"
      ) {
        return;
      }

      handleGridKey({
        altKey: event.altKey,
        cancel: () => {
          event.stopPropagation();
        },
        ctrlKey: event.ctrlKey,
        key: normalizedKey,
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

  const editorTextAlign = useMemo<"left" | "right">(
    () => (isNumericEditorSeed(editorValue) ? "right" : "left"),
    [editorValue]
  );

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
        onKeyDownCapture={(event) => {
          const normalizedKey = normalizeKeyboardKey(event.key, event.code);
          ignoreNextPointerSelectionRef.current = false;
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragHeaderSelectionRef.current = null;
          dragGeometryRef.current = null;
          dragDidMoveRef.current = false;
          dragViewportRef.current = null;
          postDragSelectionExpiryRef.current = 0;

          if (
            !isPrintableKey({
              altKey: event.altKey,
              ctrlKey: event.ctrlKey,
              key: normalizedKey,
              metaKey: event.metaKey
            } as GridKeyEventArgs)
            && !isClipboardShortcut({
              altKey: event.altKey,
              ctrlKey: event.ctrlKey,
              key: normalizedKey,
              metaKey: event.metaKey
            })
            && !isNavigationKey(normalizedKey)
            && normalizedKey !== "Enter"
            && normalizedKey !== "Tab"
            && normalizedKey !== "F2"
            && normalizedKey !== "Backspace"
            && normalizedKey !== "Delete"
          ) {
            return;
          }

          handleGridKey({
            altKey: event.altKey,
            cancel: () => event.stopPropagation(),
            ctrlKey: event.ctrlKey,
            key: normalizedKey,
            metaKey: event.metaKey,
            shiftKey: event.shiftKey,
            preventDefault: () => event.preventDefault()
          });
          if (event.defaultPrevented) {
            event.stopPropagation();
          }
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
            const bodyCell = resolvePointerCell(
              event.clientX,
              event.clientY,
              visibleRegion,
              activeGeometry
            );
            if (!sameItem(bodyCell, lastBodyClickCellRef.current)) {
              return;
            }
            if (bodyCell === null) {
              return;
            }
            event.preventDefault();
            event.stopPropagation();
            const editAddress = formatAddress(bodyCell[1], bodyCell[0]);
            setGridSelection(createGridSelection(bodyCell[0], bodyCell[1]));
            onSelect(editAddress);
            beginEditAt(editAddress);
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
            const nextHeader = resolveHeaderSelectionForPointerDrag(
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
          const headerSelection = resolveHeaderSelectionAtPointer(event.clientX, event.clientY);
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
              focusGrid();
              window.requestAnimationFrame(() => {
                ignoreNextPointerSelectionRef.current = false;
              });
              return;
            }
            ignoreNextPointerSelectionRef.current = true;
            setGridSelection(createColumnSliceSelection(headerSelection.index, headerSelection.index, selectedCell.row));
            onSelect(formatAddress(selectedCell.row, headerSelection.index));
            focusGrid();
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
          focusGrid();
        }}
        onPointerUpCapture={(event) => {
          if (columnResizeActiveRef.current) {
            return;
          }
          const headerAnchor = dragHeaderSelectionRef.current;
          if (headerAnchor) {
            const finalHeader = resolveHeaderSelectionForPointerDrag(
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
            lastBodyClickCellRef.current = null;
          } else {
            lastBodyClickCellRef.current = anchorCell;
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
          drawFocusRing={variant === "product"}
          editOnType={false}
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
            focusGrid();
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
          selectionBehavior={editorSelectionBehavior}
          textAlign={editorTextAlign}
          value={editorValue}
          style={overlayStyle}
        />
      ) : null}
    </div>
  );
}
