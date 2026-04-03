import {
  startTransition,
  useDeferredValue,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import type { CellStyleRecord, Viewport } from "@bilig/protocol";
import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import {
  DataEditor,
  type FillPatternEventArgs,
  type DataEditorRef,
  type DrawCellCallback,
  type GridColumn,
  type GridSelection,
  type Item,
  type Rectangle,
} from "@glideapps/glide-data-grid";
import { CellEditorOverlay } from "./CellEditorOverlay.js";
import {
  buildBorderOverlayState,
  shouldRefreshBorderOverlay,
  type BorderOverlaySegment,
} from "./gridBorderOverlay.js";
import {
  EMPTY_COLUMN_WIDTHS,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  getGridMetrics,
  getResolvedColumnWidth,
  type GridRect,
} from "./gridMetrics.js";
import {
  createGridSelection,
  formatSelectionSummary,
  rectangleToAddresses,
} from "./gridSelection.js";
import {
  createPointerGeometry,
  resolveHeaderSelection as resolveHeaderSelectionFromGeometry,
  resolveHeaderSelectionForDrag as resolveHeaderSelectionForDragFromGeometry,
  resolvePointerCell as resolvePointerCellFromGeometry,
  resolveColumnResizeTarget,
  type HeaderSelection,
  type PointerGeometry,
  type VisibleRegionState,
} from "./gridPointer.js";
import { cellToEditorSeed, cellToGridCell, getResolvedCellFontFamily } from "./gridCells.js";
import { isHandledGridKey } from "./gridKeyboard.js";
import { getEditorTextAlign, getGridTheme, getOverlayStyle } from "./gridPresentation.js";
import type { InternalClipboardRange } from "./gridInternalClipboard.js";
import {
  finishGridResize,
  handleGridBodyDoubleClick,
  handleGridCellActivated,
  handleGridHeaderClick,
  handleGridPointerDown,
  handleGridPointerMove,
  handleGridPointerUp,
  handleGridSelectionChange,
  startGridResize,
} from "./gridInteractionController.js";
import {
  clearGridPendingPointerActivation,
  resetGridPointerInteraction,
} from "./gridInteractionState.js";
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  getNormalizedGridKeyboardKey,
  handleGridCopyCapture,
  handleGridKey as dispatchGridKey,
  handleGridPasteCapture,
  shouldHandleDataEditorGridKey,
  shouldHandleGridWindowKey,
  type GridKeyboardEventLike,
} from "./gridClipboardKeyboardController.js";
import type { GridEngineLike } from "./grid-engine.js";

export type EditMovement = readonly [-1 | 0 | 1, -1 | 0 | 1];
export type EditSelectionBehavior = "select-all" | "caret-end";
export type SheetGridViewportSubscription = (
  sheetName: string,
  viewport: Viewport,
  listener: (damage?: readonly { cell: Item }[]) => void,
) => () => void;

interface SheetGridViewProps {
  engine: GridEngineLike;
  sheetName: string;
  selectedAddr: string;
  editorValue: string;
  editorSelectionBehavior: EditSelectionBehavior;
  resolvedValue: string;
  isEditingCell: boolean;
  onSelect(this: void, addr: string): void;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
  onBeginEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void;
  onEditorChange(this: void, next: string): void;
  onCommitEdit(this: void, movement?: EditMovement): void;
  onCancelEdit(this: void): void;
  onClearCell(this: void): void;
  onFillRange(
    this: void,
    sourceStartAddr: string,
    sourceEndAddr: string,
    targetStartAddr: string,
    targetEndAddr: string,
  ): void;
  onCopyRange(
    this: void,
    sourceStartAddr: string,
    sourceEndAddr: string,
    targetStartAddr: string,
    targetEndAddr: string,
  ): void;
  onPaste(
    this: void,
    sheetName: string,
    addr: string,
    values: readonly (readonly string[])[],
  ): void;
  subscribeViewport?: SheetGridViewportSubscription | undefined;
  columnWidths?: Readonly<Record<number, number>> | undefined;
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined;
  onAutofitColumn?:
    | ((columnIndex: number, fallbackWidth: number) => void | Promise<void>)
    | undefined;
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined;
}

function sameBounds(left: Rectangle | undefined, right: Rectangle | undefined): boolean {
  if (left === right) {
    return true;
  }
  if (!left || !right) {
    return false;
  }
  return (
    left.x === right.x &&
    left.y === right.y &&
    left.width === right.width &&
    left.height === right.height
  );
}

function sameVisibleRegion(left: VisibleRegionState, right: VisibleRegionState): boolean {
  return (
    left.tx === right.tx &&
    left.ty === right.ty &&
    left.range.x === right.range.x &&
    left.range.y === right.range.y &&
    left.range.width === right.range.width &&
    left.range.height === right.range.height
  );
}

export function SheetGridView({
  engine,
  sheetName,
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
  onPaste,
  subscribeViewport,
  columnWidths: controlledColumnWidths,
  onColumnWidthChange,
  onAutofitColumn,
  onVisibleViewportChange,
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
  const pendingKeyboardPasteSequenceRef = useRef(0);
  const suppressNextNativePasteRef = useRef(false);
  const activeSheetRef = useRef(sheetName);
  const columnResizeActiveRef = useRef(false);
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const pendingTypeSeedRef = useRef<string | null>(null);
  const visibleBorderSignaturesRef = useRef(new Map<string, string>());
  const interactionState = useMemo(
    () => ({
      ignoreNextPointerSelectionRef,
      pendingPointerCellRef,
      dragAnchorCellRef,
      dragPointerCellRef,
      dragHeaderSelectionRef,
      dragViewportRef,
      dragGeometryRef,
      dragDidMoveRef,
      postDragSelectionExpiryRef,
      columnResizeActiveRef,
    }),
    [],
  );
  const [borderOverlayRevision, setBorderOverlayRevision] = useState(0);
  const [borderSegments, setBorderSegments] = useState<readonly BorderOverlaySegment[]>([]);
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0,
  });
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined);
  const [columnWidthsBySheet, setColumnWidthsBySheet] = useState<
    Record<string, Record<number, number>>
  >({});
  const selectedCell = useMemo(
    () => parseCellAddress(selectedAddr, sheetName),
    [selectedAddr, sheetName],
  );
  const [gridSelection, setGridSelection] = useState<GridSelection>(() =>
    createGridSelection(selectedCell.col, selectedCell.row),
  );
  const gridMetrics = useMemo(() => getGridMetrics(), []);
  const columnWidths =
    controlledColumnWidths ?? columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS;

  const columns = useMemo<readonly GridColumn[]>(
    () =>
      Array.from({ length: MAX_COLS }, (_, index) => ({
        id: indexToColumn(index),
        title: indexToColumn(index),
        width: getResolvedColumnWidth(columnWidths, index, gridMetrics.columnWidth),
      })),
    [columnWidths, gridMetrics.columnWidth],
  );
  const columnWidthOverridesAttr = useMemo(() => {
    const entries = Object.entries(columnWidths).toSorted(
      ([left], [right]) => Number(left) - Number(right),
    );
    return entries.length === 0 ? "{}" : JSON.stringify(Object.fromEntries(entries));
  }, [columnWidths]);

  const getCellContent = useCallback(
    ([col, row]: Item) => cellToGridCell(engine, sheetName, formatAddress(row, col)),
    [engine, sheetName],
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
  }, [
    visibleRegion.range.height,
    visibleRegion.range.width,
    visibleRegion.range.x,
    visibleRegion.range.y,
  ]);

  const visibleAddresses = useMemo(
    () => visibleItems.map(([col, row]) => formatAddress(row, col)),
    [visibleItems],
  );
  const visibleDamage = useMemo(() => visibleItems.map((cell) => ({ cell })), [visibleItems]);
  const deferredVisibleRegion = useDeferredValue(visibleRegion);
  const deferredVisibleItems = useDeferredValue(visibleItems);
  const viewport = useMemo<Viewport>(
    () => ({
      rowStart: visibleRegion.range.y,
      rowEnd: Math.min(MAX_ROWS - 1, visibleRegion.range.y + visibleRegion.range.height - 1),
      colStart: visibleRegion.range.x,
      colEnd: Math.min(MAX_COLS - 1, visibleRegion.range.x + visibleRegion.range.width - 1),
    }),
    [
      visibleRegion.range.height,
      visibleRegion.range.width,
      visibleRegion.range.x,
      visibleRegion.range.y,
    ],
  );

  useEffect(() => {
    onVisibleViewportChange?.(viewport);
  }, [onVisibleViewportChange, viewport]);

  useEffect(() => {
    if (subscribeViewport) {
      return subscribeViewport(sheetName, viewport, (damage) => {
        editorRef.current?.updateCells(damage ? [...damage] : visibleDamage);
        if (
          !damage ||
          shouldRefreshBorderOverlay(visibleBorderSignaturesRef.current, engine, sheetName, damage)
        ) {
          startTransition(() => {
            setBorderOverlayRevision((current) => current + 1);
          });
        }
      });
    }
    return engine.subscribeCells(sheetName, visibleAddresses, () => {
      editorRef.current?.updateCells(visibleDamage);
      startTransition(() => {
        setBorderOverlayRevision((current) => current + 1);
      });
    });
  }, [engine, sheetName, subscribeViewport, viewport, visibleAddresses, visibleDamage]);

  useEffect(() => {
    const hostBounds = hostRef.current?.getBoundingClientRect();
    const editor = editorRef.current;
    if (!hostBounds || !editor) {
      setBorderSegments([]);
      return;
    }

    const nextOverlayState = buildBorderOverlayState(
      engine,
      sheetName,
      deferredVisibleItems,
      hostBounds,
      (col, row) => editor.getBounds(col, row),
    );
    visibleBorderSignaturesRef.current = nextOverlayState.signatures;
    setBorderSegments(nextOverlayState.segments);
  }, [borderOverlayRevision, deferredVisibleItems, deferredVisibleRegion, engine, sheetName]);

  useLayoutEffect(() => {
    editorRef.current?.scrollTo(selectedCell.col, selectedCell.row);
  }, [selectedCell.col, selectedCell.row, sheetName]);

  useLayoutEffect(() => {
    const sheetChanged = activeSheetRef.current !== sheetName;
    activeSheetRef.current = sheetName;
    setGridSelection((current) => {
      const currentCell = current.current?.cell;
      if (
        !sheetChanged &&
        currentCell &&
        currentCell[0] === selectedCell.col &&
        currentCell[1] === selectedCell.row
      ) {
        return current;
      }
      clearGridPendingPointerActivation(interactionState);
      dragGeometryRef.current = null;
      return createGridSelection(selectedCell.col, selectedCell.row);
    });
  }, [interactionState, selectedCell.col, selectedCell.row, sheetName]);

  const refreshOverlayBounds = useCallback(() => {
    const next = editorRef.current?.getBounds(selectedCell.col, selectedCell.row);
    setOverlayBounds((current) => {
      if (!next) {
        return current;
      }
      return sameBounds(current, next) ? current : next;
    });
  }, [selectedCell.col, selectedCell.row]);

  useLayoutEffect(() => {
    if (!isEditingCell) {
      setOverlayBounds(undefined);
      return;
    }

    const frame = window.requestAnimationFrame(refreshOverlayBounds);
    return () => {
      window.cancelAnimationFrame(frame);
    };
  }, [isEditingCell, refreshOverlayBounds, visibleRegion.tx, visibleRegion.ty]);

  useEffect(() => {
    if (!isEditingCell) {
      return;
    }
    const handleWindowResize = () => refreshOverlayBounds();
    window.addEventListener("resize", handleWindowResize);
    return () => {
      window.removeEventListener("resize", handleWindowResize);
    };
  }, [isEditingCell, refreshOverlayBounds]);

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
        seed ?? cellToEditorSeed(engine.getCell(sheetName, selectedAddr)),
        selectionBehavior,
      );
    },
    [engine, onBeginEdit, selectedAddr, sheetName],
  );

  const beginEditAt = useCallback(
    (addr: string, seed?: string, selectionBehavior: EditSelectionBehavior = "caret-end") => {
      onBeginEdit(seed ?? cellToEditorSeed(engine.getCell(sheetName, addr)), selectionBehavior);
    },
    [engine, onBeginEdit, sheetName],
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
        bottom: hostBounds.bottom,
      };
      return createPointerGeometry(rect, region, columnWidths, gridMetrics);
    },
    [columnWidths, gridMetrics, visibleRegion],
  );

  const resolvePointerCell = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null,
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
        selectedCellBounds:
          editorRef.current?.getBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange: gridSelection.current?.range ?? null,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      });
    },
    [
      columnWidths,
      gridMetrics,
      gridSelection,
      resolvePointerGeometry,
      selectedCell.col,
      selectedCell.row,
      visibleRegion,
    ],
  );

  const resolveHeaderSelectionAtPointer = useCallback(
    (
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region);
      if (!activeGeometry) {
        return null;
      }
      return resolveHeaderSelectionFromGeometry(
        clientX,
        clientY,
        region,
        activeGeometry,
        columnWidths,
        gridMetrics,
      );
    },
    [columnWidths, gridMetrics, resolvePointerGeometry, visibleRegion],
  );

  const resolveHeaderSelectionForPointerDrag = useCallback(
    (
      kind: HeaderSelection["kind"],
      clientX: number,
      clientY: number,
      region: VisibleRegionState = visibleRegion,
      geometry?: PointerGeometry | null,
    ): HeaderSelection | null => {
      const activeGeometry = geometry ?? resolvePointerGeometry(region);
      if (!activeGeometry) {
        return null;
      }
      return resolveHeaderSelectionForDragFromGeometry(
        kind,
        clientX,
        clientY,
        region,
        activeGeometry,
        columnWidths,
        gridMetrics,
      );
    },
    [columnWidths, gridMetrics, resolvePointerGeometry, visibleRegion],
  );

  const selectionSummary = useMemo(
    () => formatSelectionSummary(gridSelection, selectedAddr),
    [gridSelection, selectedAddr],
  );

  const applyClipboardValues = useCallback(
    (target: Item, values: readonly (readonly string[])[]) => {
      applyGridClipboardValues({
        internalClipboardRef,
        onCopyRange,
        onPaste,
        sheetName,
        target,
        values,
      });
    },
    [onCopyRange, onPaste, sheetName],
  );

  const captureInternalClipboardSelection = useCallback(() => {
    return captureGridClipboardSelection({
      engine,
      gridSelection,
      internalClipboardRef,
      sheetName,
    });
  }, [engine, gridSelection, sheetName]);

  useEffect(() => {
    onSelectionLabelChange?.(selectionSummary);
  }, [onSelectionLabelChange, selectionSummary]);

  const handleGridKey = useCallback(
    (event: GridKeyboardEventLike) => {
      dispatchGridKey({
        applyClipboardValues,
        beginSelectedEdit,
        captureInternalClipboardSelection,
        editorValue,
        event,
        gridSelection,
        isEditingCell,
        onCancelEdit,
        onClearCell,
        onCommitEdit,
        onEditorChange,
        onSelect,
        pendingKeyboardPasteSequenceRef,
        pendingTypeSeedRef,
        selectedCell,
        setGridSelection,
        suppressNextNativePasteRef,
      });
    },
    [
      applyClipboardValues,
      beginSelectedEdit,
      captureInternalClipboardSelection,
      editorValue,
      gridSelection,
      isEditingCell,
      onCancelEdit,
      onClearCell,
      onCommitEdit,
      onEditorChange,
      onSelect,
      selectedCell,
    ],
  );

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code);
      const activeElement = document.activeElement;
      if (
        !shouldHandleGridWindowKey(
          {
            altKey: event.altKey,
            ctrlKey: event.ctrlKey,
            key: normalizedKey,
            metaKey: event.metaKey,
            shiftKey: event.shiftKey,
          },
          activeElement,
          hostRef.current,
        )
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
        preventDefault: () => event.preventDefault(),
      });
    };

    window.addEventListener("keydown", handleWindowKeyDown, true);
    return () => {
      window.removeEventListener("keydown", handleWindowKeyDown, true);
    };
  }, [handleGridKey]);

  const overlayStyle = useMemo(
    () => getOverlayStyle(isEditingCell, overlayBounds),
    [isEditingCell, overlayBounds],
  );

  const editorTextAlign = useMemo<"left" | "right">(
    () => getEditorTextAlign(editorValue),
    [editorValue],
  );

  const gridTheme = useMemo(() => getGridTheme(), []);

  const handleFillPattern = useCallback(
    (event: FillPatternEventArgs) => {
      const source = rectangleToAddresses(event.patternSource);
      const target = rectangleToAddresses(event.fillDestination);
      if (source.startAddress === target.startAddress && source.endAddress === target.endAddress) {
        return;
      }
      event.preventDefault();
      onFillRange(source.startAddress, source.endAddress, target.startAddress, target.endAddress);
    },
    [onFillRange],
  );

  const applyColumnWidth = useCallback(
    (columnIndex: number, newSize: number) => {
      const clampedSize = Math.max(
        MIN_COLUMN_WIDTH,
        Math.min(MAX_COLUMN_WIDTH, Math.round(newSize)),
      );
      if (onColumnWidthChange) {
        onColumnWidthChange(columnIndex, clampedSize);
        return;
      }
      setColumnWidthsBySheet((current) => {
        const nextSheetWidths = current[sheetName] ?? EMPTY_COLUMN_WIDTHS;
        if (nextSheetWidths[columnIndex] === clampedSize) {
          return current;
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetWidths,
            [columnIndex]: clampedSize,
          },
        };
      });
    },
    [onColumnWidthChange, sheetName],
  );

  const computeAutofitColumnWidth = useCallback(
    (columnIndex: number): number => {
      const canvas = textMeasureCanvasRef.current ?? document.createElement("canvas");
      textMeasureCanvasRef.current = canvas;
      const context = canvas.getContext("2d");
      if (!context) {
        return gridMetrics.columnWidth;
      }

      const headerFont = gridTheme.headerFontStyle;
      let measuredWidth = 0;

      context.font = headerFont;
      measuredWidth = Math.max(
        measuredWidth,
        context.measureText(indexToColumn(columnIndex)).width,
      );

      const sheet = engine.workbook.getSheet(sheetName);
      sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (col !== columnIndex) {
          return;
        }
        const cell = cellToGridCell(engine, sheetName, formatAddress(row, col));
        const displayText =
          "displayData" in cell
            ? String(cell.displayData ?? "")
            : "copyData" in cell
              ? String(cell.copyData ?? "")
              : "";
        context.font = `400 ${gridTheme.editorFontSize} ${getResolvedCellFontFamily()}`;
        measuredWidth = Math.max(measuredWidth, context.measureText(displayText).width);
      });

      return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(measuredWidth + 28)));
    },
    [
      engine,
      gridMetrics.columnWidth,
      gridTheme.editorFontSize,
      gridTheme.headerFontStyle,
      sheetName,
    ],
  );

  const drawCell = useCallback<DrawCellCallback>(
    (args, drawContent) => {
      const snapshot = engine.getCell(sheetName, formatAddress(args.row, args.col));
      const style = engine.getCellStyle(snapshot.styleId);
      if (style?.fill?.backgroundColor) {
        args.ctx.save();
        args.ctx.fillStyle = style.fill.backgroundColor;
        args.ctx.fillRect(
          args.rect.x + 1,
          args.rect.y + 1,
          Math.max(0, args.rect.width - 2),
          Math.max(0, args.rect.height - 2),
        );
        args.ctx.restore();
      }
      drawContent();
      if (!style) {
        return;
      }
      if (style.font?.underline && "displayData" in args.cell) {
        drawCellUnderline(
          args,
          style,
          String(args.cell.displayData ?? ""),
          gridTheme.editorFontSize,
        );
      }
    },
    [engine, gridTheme.editorFontSize, sheetName],
  );

  return (
    <div
      className="relative flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]"
      data-testid="sheet-grid-shell"
    >
      <div
        className="sheet-grid-host min-h-0 flex-1 bg-[var(--wb-surface)] pr-2 pb-2"
        data-column-width-overrides={columnWidthOverridesAttr}
        data-default-column-width={gridMetrics.columnWidth}
        data-testid="sheet-grid"
        role="grid"
        onKeyDownCapture={(event) => {
          const normalizedKey = getNormalizedGridKeyboardKey(event.key, event.code);
          resetGridPointerInteraction(interactionState, {
            clearIgnoreNextPointerSelection: true,
          });

          if (
            !isHandledGridKey({
              altKey: event.altKey,
              ctrlKey: event.ctrlKey,
              key: normalizedKey,
              metaKey: event.metaKey,
            })
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
            preventDefault: () => event.preventDefault(),
          });
          if (event.defaultPrevented) {
            event.stopPropagation();
          }
        }}
        onCopyCapture={(event) => {
          handleGridCopyCapture({
            captureInternalClipboardSelection,
            event,
            internalClipboardRef,
          });
        }}
        onPasteCapture={(event) => {
          handleGridPasteCapture({
            applyClipboardValues,
            event,
            gridSelection,
            pendingKeyboardPasteSequenceRef,
            selectedCell,
            suppressNextNativePasteRef,
          });
        }}
        onKeyDown={(event) => handleGridKey(event)}
        onDoubleClickCapture={(event) => {
          handleGridBodyDoubleClick({
            event,
            applyColumnWidth,
            beginEditAt,
            columnWidths,
            computeAutofitColumnWidth,
            defaultColumnWidth: gridMetrics.columnWidth,
            interactionState,
            isEditingCell,
            lastBodyClickCell: lastBodyClickCellRef.current,
            onAutofitColumn,
            onCommitEdit: () => onCommitEdit(),
            onSelect,
            resolvePointerCell,
            resolvePointerGeometry,
            selectedCell: [selectedCell.col, selectedCell.row],
            setGridSelection,
            visibleRegion,
          });
        }}
        onPointerMoveCapture={(event) => {
          handleGridPointerMove({
            dragAnchorCell: dragAnchorCellRef.current,
            dragGeometry: dragGeometryRef.current,
            dragHeaderSelection: dragHeaderSelectionRef.current,
            dragPointerCell: dragPointerCellRef.current,
            dragViewport: dragViewportRef.current,
            event,
            interactionState,
            isEditingCell,
            onCommitEdit: () => onCommitEdit(),
            onSelect,
            resolveHeaderSelectionForPointerDrag,
            resolvePointerCell,
            selectedCell: [selectedCell.col, selectedCell.row],
            setGridSelection,
            visibleRegion,
          });
        }}
        onPointerDownCapture={(event) => {
          handleGridPointerDown({
            columnWidths,
            defaultColumnWidth: gridMetrics.columnWidth,
            event,
            focusGrid,
            interactionState,
            isEditingCell,
            onCommitEdit: () => onCommitEdit(),
            onSelect,
            resolveColumnResizeTargetAtPointer: resolveColumnResizeTarget,
            resolveHeaderSelectionAtPointer,
            resolvePointerCell,
            resolvePointerGeometry,
            selectedCell: [selectedCell.col, selectedCell.row],
            setGridSelection,
            visibleRegion,
          });
        }}
        onPointerUpCapture={(event) => {
          handleGridPointerUp({
            dragAnchorCell: dragAnchorCellRef.current,
            dragDidMove: dragDidMoveRef.current,
            dragGeometry: dragGeometryRef.current,
            dragHeaderSelection: dragHeaderSelectionRef.current,
            dragPointerCell: dragPointerCellRef.current,
            dragViewport: dragViewportRef.current,
            event,
            interactionState,
            isEditingCell,
            lastBodyClickCellRef,
            onCommitEdit: () => onCommitEdit(),
            onSelect,
            postDragSelectionExpiryRef,
            resolveHeaderSelectionForPointerDrag,
            resolvePointerCell,
            selectedCell: [selectedCell.col, selectedCell.row],
            setGridSelection,
            visibleRegion,
          });
        }}
        ref={hostRef}
        // oxlint-disable-next-line jsx-a11y/no-noninteractive-tabindex
        tabIndex={0}
      >
        <DataEditor
          ref={editorRef}
          cellActivationBehavior="double-click"
          className="glide-sheet-grid"
          columns={columns}
          drawCell={drawCell}
          drawFocusRing={true}
          editOnType={false}
          fillHandle={true}
          freezeColumns={0}
          getCellContent={getCellContent}
          getCellsForSelection={true}
          gridSelection={gridSelection}
          headerHeight={gridMetrics.headerHeight}
          height="100%"
          maxColumnWidth={MAX_COLUMN_WIDTH}
          minColumnWidth={MIN_COLUMN_WIDTH}
          onColumnResizeStart={() => {
            startGridResize(interactionState);
          }}
          onColumnResize={(_column: GridColumn, newSize: number, columnIndex: number) => {
            applyColumnWidth(columnIndex, newSize);
          }}
          onColumnResizeEnd={(_column: GridColumn, newSize: number, columnIndex: number) => {
            applyColumnWidth(columnIndex, newSize);
            finishGridResize(interactionState);
          }}
          onCellActivated={([col, row]) => {
            handleGridCellActivated({
              activatedCell: [col, row],
              dragAnchorCell: dragAnchorCellRef.current,
              interactionState,
              onSelect,
              pendingPointerCell: pendingPointerCellRef.current,
              setGridSelection,
            });
          }}
          onDelete={() => {
            onClearCell();
            return false;
          }}
          onFillPattern={handleFillPattern}
          onHeaderClicked={(col, event) => {
            handleGridHeaderClick({
              applyColumnWidth,
              columnIndex: col,
              computeAutofitColumnWidth,
              event,
              focusGrid,
              interactionState,
              isEditingCell,
              onCommitEdit: () => onCommitEdit(),
              onSelect,
              selectedRow: selectedCell.row,
              setGridSelection,
            });
          }}
          onGridSelectionChange={(nextSelection) => {
            handleGridSelectionChange({
              dragAnchorCell: dragAnchorCellRef.current,
              dragPointerCell: dragPointerCellRef.current,
              interactionState,
              isEditingCell,
              nextSelection,
              now: window.performance.now(),
              onCommitEdit: () => onCommitEdit(),
              onSelect,
              pendingPointerCell: pendingPointerCellRef.current,
              selectedCell: [selectedCell.col, selectedCell.row],
              setGridSelection,
            });
          }}
          onKeyDown={(event) => {
            if (!shouldHandleDataEditorGridKey(event)) {
              return;
            }
            handleGridKey(event);
          }}
          onPaste={(target, values) => {
            applyClipboardValues(target, values);
            return false;
          }}
          onVisibleRegionChanged={(range, tx, ty) => {
            startTransition(() => {
              setVisibleRegion((current) => {
                const next = { range, tx, ty };
                return sameVisibleRegion(current, next) ? current : next;
              });
            });
          }}
          rowHeight={gridMetrics.rowHeight}
          columnSelect="multi"
          columnSelectionBlending="additive"
          columnSelectionMode="multi"
          rowMarkers={{
            kind: "clickable-number",
            width: gridMetrics.rowMarkerWidth,
            headerDisabled: true,
          }}
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
        {borderSegments.length > 0 ? (
          <div
            className="pointer-events-none absolute inset-0 z-10"
            aria-hidden="true"
            data-testid="grid-border-overlay"
          >
            {borderSegments.map((segment) => (
              <div
                key={segment.key}
                className="absolute"
                data-testid="grid-border-overlay-segment"
                style={segment.style}
              />
            ))}
          </div>
        ) : null}
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

function drawCellUnderline(
  args: Parameters<DrawCellCallback>[0],
  style: CellStyleRecord,
  text: string,
  editorFontSize: string,
): void {
  if (!text) {
    return;
  }
  const size = Number.parseInt(editorFontSize, 10) || style.font?.size || 13;
  const fontParts = [
    style.font?.italic ? "italic" : "",
    style.font?.bold ? "700" : "400",
    `${style.font?.size ?? size}px`,
    getResolvedCellFontFamily(),
  ].filter(Boolean);
  const padding = 8;
  args.ctx.save();
  args.ctx.font = fontParts.join(" ");
  const textWidth = args.ctx.measureText(text).width;
  const align = "contentAlign" in args.cell ? args.cell.contentAlign : "left";
  const startX =
    align === "right"
      ? args.rect.x + args.rect.width - padding - textWidth
      : align === "center"
        ? args.rect.x + (args.rect.width - textWidth) / 2
        : args.rect.x + padding;
  const endX = startX + textWidth;
  const y = args.rect.y + args.rect.height - 5;
  args.ctx.strokeStyle = style.font?.color ?? args.theme.textDark;
  args.ctx.lineWidth = 1;
  args.ctx.beginPath();
  args.ctx.moveTo(startX, y);
  args.ctx.lineTo(endX, y);
  args.ctx.stroke();
  args.ctx.restore();
}
