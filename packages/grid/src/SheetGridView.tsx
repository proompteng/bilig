import {
  startTransition,
  useDeferredValue,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  type PointerEvent as ReactPointerEvent,
  useRef,
  useState,
} from "react";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import type { CellStyleRecord, Viewport } from "@bilig/protocol";
import { MAX_COLS, MAX_ROWS, ValueTag } from "@bilig/protocol";
import {
  DataEditor,
  type DataEditorRef,
  type DrawCellCallback,
  type GridColumn,
  type GridSelection,
  type Item,
  type Rectangle,
} from "@glideapps/glide-data-grid";
import { CellEditorOverlay } from "./CellEditorOverlay.js";
import { GridGpuSurface } from "./GridGpuSurface.js";
import { GridTextOverlay } from "./GridTextOverlay.js";
import {
  buildBorderOverlayState,
  shouldRefreshBorderOverlay,
  type BorderOverlaySegment,
} from "./gridBorderOverlay.js";
import { buildGridGpuScene, type GridGpuScene } from "./gridGpuScene.js";
import { buildGridTextScene, type GridTextScene } from "./gridTextScene.js";
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
  createSheetSelection,
  formatSelectionSummary,
  isSheetSelection,
  rectangleToAddresses,
} from "./gridSelection.js";
import {
  resolveFillHandleOverlayBounds,
  resolveFillHandlePreviewRange,
  type FillHandleOverlayBounds,
} from "./gridFillHandle.js";
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
import { resolveGridHoverState, sameGridHoverState, type GridHoverState } from "./gridHover.js";
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
  onToggleBooleanCell?:
    | ((sheetName: string, address: string, nextValue: boolean) => void)
    | undefined;
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
  onToggleBooleanCell,
  onPaste,
  subscribeViewport,
  columnWidths: controlledColumnWidths,
  onColumnWidthChange,
  onAutofitColumn,
  onVisibleViewportChange,
}: SheetGridViewProps) {
  const emptyGpuScene = useMemo<GridGpuScene>(() => ({ fillRects: [], borderRects: [] }), []);
  const emptyTextScene = useMemo<GridTextScene>(() => ({ items: [] }), []);
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
  const fillPreviewRangeRef = useRef<Rectangle | null>(null);
  const fillHandleCleanupRef = useRef<(() => void) | null>(null);
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
  const [gpuScene, setGpuScene] = useState<GridGpuScene>(emptyGpuScene);
  const [textScene, setTextScene] = useState<GridTextScene>(emptyTextScene);
  const [fillHandleBounds, setFillHandleBounds] = useState<FillHandleOverlayBounds | undefined>(
    undefined,
  );
  const [fillPreviewRange, setFillPreviewRange] = useState<Rectangle | null>(null);
  const [hoverState, setHoverState] = useState<GridHoverState>({
    cell: null,
    header: null,
    cursor: "default",
  });
  const [activeResizeColumn, setActiveResizeColumn] = useState<number | null>(null);
  const [activeHeaderDrag, setActiveHeaderDrag] = useState<HeaderSelection | null>(null);
  const [isWebGpuActive, setIsWebGpuActive] = useState(false);
  const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null);
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
    ([col, row]: Item) =>
      cellToGridCell(engine, sheetName, formatAddress(row, col), {
        booleanSurfaceEnabled: isWebGpuActive,
        textSurfaceEnabled: isWebGpuActive,
      }),
    [engine, isWebGpuActive, sheetName],
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
  const selectionRange = gridSelection.current?.range ?? null;
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
    fillPreviewRangeRef.current = fillPreviewRange;
  }, [fillPreviewRange]);

  useEffect(() => {
    return () => {
      fillHandleCleanupRef.current?.();
    };
  }, []);

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
      setGpuScene(emptyGpuScene);
      setFillHandleBounds(undefined);
      setFillPreviewRange(null);
      setTextScene(emptyTextScene);
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
    setGpuScene(
      buildGridGpuScene({
        engine,
        columnWidths,
        fillPreviewRange,
        gridMetrics,
        gridSelection,
        activeHeaderDrag,
        hoveredCell: hoverState.cell,
        hoveredHeader: hoverState.header,
        resizeGuideColumn:
          activeResizeColumn ??
          (hoverState.cursor === "col-resize" && hoverState.header?.kind === "column"
            ? hoverState.header.index
            : null),
        selectedCell: [selectedCell.col, selectedCell.row],
        selectionRange,
        sheetName,
        visibleItems: deferredVisibleItems,
        visibleRegion: deferredVisibleRegion,
        hostBounds,
        getCellBounds: (col, row) => editor.getBounds(col, row),
      }),
    );
    setTextScene(
      buildGridTextScene({
        engine,
        columnWidths,
        gridMetrics,
        activeHeaderDrag,
        hoveredHeader: hoverState.header,
        resizeGuideColumn:
          activeResizeColumn ??
          (hoverState.cursor === "col-resize" && hoverState.header?.kind === "column"
            ? hoverState.header.index
            : null),
        selectedCell: [selectedCell.col, selectedCell.row],
        selectionRange,
        sheetName,
        visibleItems: deferredVisibleItems,
        visibleRegion: deferredVisibleRegion,
        hostBounds,
        getCellBounds: (col, row) => editor.getBounds(col, row),
      }),
    );
    if (
      selectionRange &&
      gridSelection.columns.length === 0 &&
      gridSelection.rows.length === 0 &&
      !fillPreviewRange
    ) {
      setFillHandleBounds(
        resolveFillHandleOverlayBounds({
          sourceRange: selectionRange,
          hostBounds,
          getCellBounds: (col, row) => editor.getBounds(col, row),
        }),
      );
    } else {
      setFillHandleBounds(undefined);
    }
  }, [
    borderOverlayRevision,
    columnWidths,
    deferredVisibleItems,
    deferredVisibleRegion,
    emptyGpuScene,
    emptyTextScene,
    engine,
    fillPreviewRange,
    gridMetrics,
    gridSelection,
    activeHeaderDrag,
    hoverState.cell,
    hoverState.header,
    hoverState.cursor,
    activeResizeColumn,
    sheetName,
    selectedCell.col,
    selectedCell.row,
    selectionRange,
  ]);

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

  const toggleBooleanCellAt = useCallback(
    (col: number, row: number): boolean => {
      if (!onToggleBooleanCell) {
        return false;
      }
      const address = formatAddress(row, col);
      const snapshot = engine.getCell(sheetName, address);
      if (snapshot.value.tag !== ValueTag.Boolean) {
        return false;
      }
      onToggleBooleanCell(sheetName, address, !snapshot.value.value);
      return true;
    },
    [engine, onToggleBooleanCell, sheetName],
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
  const isEntireSheetSelected = useMemo(() => isSheetSelection(gridSelection), [gridSelection]);

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
        isSelectedCellBoolean: () =>
          engine.getCell(sheetName, formatAddress(selectedCell.row, selectedCell.col)).value.tag ===
          ValueTag.Boolean,
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
        toggleSelectedBooleanCell: () => {
          toggleBooleanCellAt(selectedCell.col, selectedCell.row);
        },
      });
    },
    [
      applyClipboardValues,
      beginSelectedEdit,
      captureInternalClipboardSelection,
      engine,
      editorValue,
      gridSelection,
      isEditingCell,
      onCancelEdit,
      onClearCell,
      onCommitEdit,
      onEditorChange,
      onSelect,
      sheetName,
      selectedCell,
      toggleBooleanCellAt,
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

  const gridTheme = useMemo(
    () =>
      getGridTheme({
        gpuSurfaceEnabled: isWebGpuActive,
        textSurfaceEnabled: isWebGpuActive,
      }),
    [isWebGpuActive],
  );

  const handleFillHandlePointerDown = useCallback(
    (event: ReactPointerEvent<HTMLButtonElement>) => {
      if (!selectionRange) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      focusGrid();

      const move = (nativeEvent: PointerEvent) => {
        const pointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY);
        const nextPreviewRange = pointerCell
          ? resolveFillHandlePreviewRange(selectionRange, pointerCell)
          : null;
        fillPreviewRangeRef.current = nextPreviewRange;
        setFillPreviewRange(nextPreviewRange);
      };

      const cleanup = () => {
        window.removeEventListener("pointermove", move);
        window.removeEventListener("pointerup", up, true);
        fillHandleCleanupRef.current = null;
      };

      const finish = () => {
        const previewRange = fillPreviewRangeRef.current;
        if (previewRange) {
          const source = rectangleToAddresses(selectionRange);
          const target = rectangleToAddresses(previewRange);
          if (
            source.startAddress !== target.startAddress ||
            source.endAddress !== target.endAddress
          ) {
            onFillRange(
              source.startAddress,
              source.endAddress,
              target.startAddress,
              target.endAddress,
            );
          }
        }
        fillPreviewRangeRef.current = null;
        setFillPreviewRange(null);
        cleanup();
      };

      const up = () => {
        finish();
      };

      fillHandleCleanupRef.current?.();
      fillHandleCleanupRef.current = cleanup;
      window.addEventListener("pointermove", move);
      window.addEventListener("pointerup", up, true);
    },
    [focusGrid, onFillRange, resolvePointerCell, selectionRange],
  );

  const refreshHoverState = useCallback(
    (clientX: number, clientY: number, buttons: number) => {
      if (buttons !== 0 || fillPreviewRangeRef.current) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
            ? current
            : { cell: null, header: null, cursor: "default" },
        );
        return;
      }
      const geometry = resolvePointerGeometry(visibleRegion);
      if (!geometry) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
            ? current
            : { cell: null, header: null, cursor: "default" },
        );
        return;
      }
      const next = resolveGridHoverState({
        clientX,
        clientY,
        region: visibleRegion,
        geometry,
        columnWidths,
        defaultColumnWidth: gridMetrics.columnWidth,
        gridMetrics,
        selectedCell: [selectedCell.col, selectedCell.row],
        selectedCellBounds:
          editorRef.current?.getBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      });
      setHoverState((current) => (sameGridHoverState(current, next) ? current : next));
    },
    [
      columnWidths,
      gridMetrics,
      gridSelection.columns.length,
      gridSelection.rows.length,
      resolvePointerGeometry,
      selectedCell.col,
      selectedCell.row,
      selectionRange,
      visibleRegion,
    ],
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
      if (!isWebGpuActive && style?.fill?.backgroundColor) {
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
      if (!style || isWebGpuActive) {
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
    [engine, gridTheme.editorFontSize, isWebGpuActive, sheetName],
  );

  const handleHostRef = useCallback((node: HTMLDivElement | null) => {
    hostRef.current = node;
    setHostElement(node);
  }, []);

  const handleSelectEntireSheet = useCallback(() => {
    if (isEditingCell) {
      onCommitEdit();
    }
    setGridSelection(createSheetSelection());
    onSelect(formatAddress(0, 0));
    focusGrid();
  }, [focusGrid, isEditingCell, onCommitEdit, onSelect]);

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
        style={{ cursor: hoverState.cursor }}
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
          refreshHoverState(event.clientX, event.clientY, event.buttons);
        }}
        onPointerLeave={() => {
          if (activeResizeColumn !== null) {
            return;
          }
          setHoverState((current) =>
            sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
              ? current
              : { cell: null, header: null, cursor: "default" },
          );
        }}
        onPointerDownCapture={(event) => {
          const pointerGeometry = resolvePointerGeometry(visibleRegion);
          const resizeTarget =
            pointerGeometry === null
              ? null
              : resolveColumnResizeTarget(
                  event.clientX,
                  event.clientY,
                  visibleRegion,
                  pointerGeometry,
                  columnWidths,
                  gridMetrics.columnWidth,
                );
          const headerSelection =
            pointerGeometry === null
              ? null
              : resolveHeaderSelectionAtPointer(
                  event.clientX,
                  event.clientY,
                  visibleRegion,
                  pointerGeometry,
                );
          setActiveResizeColumn(resizeTarget);
          setActiveHeaderDrag(resizeTarget === null ? headerSelection : null);
          setHoverState((current) =>
            sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
              ? current
              : { cell: null, header: null, cursor: "default" },
          );
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
          setActiveHeaderDrag(null);
          refreshHoverState(event.clientX, event.clientY, 0);
        }}
        ref={handleHostRef}
        // oxlint-disable-next-line jsx-a11y/no-noninteractive-tabindex
        tabIndex={0}
      >
        <GridGpuSurface host={hostElement} scene={gpuScene} onActiveChange={setIsWebGpuActive} />
        <GridTextOverlay active={isWebGpuActive} host={hostElement} scene={textScene} />
        <button
          aria-label="Select entire sheet"
          className="absolute z-20 flex items-center justify-center border-r border-b border-[var(--wb-border-subtle)] bg-[var(--wb-muted)] text-[var(--wb-text-muted)] outline-none transition-colors hover:bg-[var(--wb-muted-strong)] hover:text-[var(--wb-text)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent)] focus-visible:ring-offset-0"
          data-testid="grid-select-entire-sheet"
          onClick={handleSelectEntireSheet}
          style={{
            height: gridMetrics.headerHeight,
            left: 0,
            top: 0,
            width: gridMetrics.rowMarkerWidth,
          }}
          type="button"
        >
          <span
            aria-hidden="true"
            className="block h-0 w-0 border-t-[11px] border-r-[11px] border-t-transparent border-r-current opacity-80"
            style={{
              color: isEntireSheetSelected ? "var(--wb-accent)" : "currentColor",
              transform: "translate(2px, 1px)",
            }}
          />
        </button>
        {fillHandleBounds ? (
          <button
            aria-label="Fill handle"
            className="absolute z-30 cursor-crosshair rounded-[2px] border border-white bg-[#1f7a43] shadow-[0_0_0_1px_rgba(31,122,67,0.45)]"
            onPointerDown={handleFillHandlePointerDown}
            style={{
              height: fillHandleBounds.height,
              left: fillHandleBounds.x,
              top: fillHandleBounds.y,
              width: fillHandleBounds.width,
            }}
            type="button"
          />
        ) : null}
        <div className="relative z-[1] h-full">
          <DataEditor
            ref={editorRef}
            cellActivationBehavior="double-click"
            className="glide-sheet-grid"
            columns={columns}
            drawCell={drawCell}
            drawFocusRing={!isWebGpuActive}
            editOnType={false}
            experimental={{
              disableAccessibilityTree: true,
              kineticScrollPerfHack: true,
              renderStrategy: "direct",
            }}
            fillHandle={false}
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
              setActiveResizeColumn(columnIndex);
              applyColumnWidth(columnIndex, newSize);
            }}
            onColumnResizeEnd={(_column: GridColumn, newSize: number, columnIndex: number) => {
              applyColumnWidth(columnIndex, newSize);
              setActiveResizeColumn(null);
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
              if (isWebGpuActive) {
                toggleBooleanCellAt(col, row);
              }
            }}
            onDelete={() => {
              onClearCell();
              return false;
            }}
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
        </div>
        {!isWebGpuActive && borderSegments.length > 0 ? (
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
