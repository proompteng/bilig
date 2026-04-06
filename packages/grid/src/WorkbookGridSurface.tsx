import {
  useDeferredValue,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
} from "react";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, Viewport } from "@bilig/protocol";
import { MAX_COLS, MAX_ROWS, ValueTag } from "@bilig/protocol";
import { CellEditorOverlay } from "./CellEditorOverlay.js";
import { GridGpuSurface } from "./GridGpuSurface.js";
import { GridTextOverlay } from "./GridTextOverlay.js";
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
  createRectangleSelectionFromRange,
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
  resolveMovedRange,
  resolveSelectionMoveAnchorCell,
  sameRectangle,
} from "./gridRangeMove.js";
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
import { cellToEditorSeed, getResolvedCellFontFamily, snapshotToRenderCell } from "./gridCells.js";
import { isHandledGridKey } from "./gridKeyboard.js";
import {
  getEditorPresentation,
  getEditorTextAlign,
  getGridTheme,
  getOverlayStyle,
} from "./gridPresentation.js";
import type { InternalClipboardRange } from "./gridInternalClipboard.js";
import {
  finishGridResize,
  handleGridBodyDoubleClick,
  handleGridPointerDown,
  handleGridPointerMove,
  handleGridPointerUp,
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
  shouldHandleGridWindowKey,
  type GridKeyboardEventLike,
} from "./gridClipboardKeyboardController.js";
import type { GridEngineLike } from "./grid-engine.js";
import type { GridSelection, Item, Rectangle } from "./gridTypes.js";

export type EditMovement = readonly [-1 | 0 | 1, -1 | 0 | 1];
export type EditSelectionBehavior = "select-all" | "caret-end";
export type SheetGridViewportSubscription = (
  sheetName: string,
  viewport: Viewport,
  listener: (damage?: readonly { cell: Item }[]) => void,
) => () => void;

export interface WorkbookGridSurfaceProps {
  engine: GridEngineLike;
  sheetName: string;
  selectedAddr: string;
  selectedCellSnapshot: CellSnapshot;
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
  onMoveRange(
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

export function WorkbookGridSurface({
  engine,
  sheetName,
  selectedAddr,
  selectedCellSnapshot,
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
  onMoveRange,
  onToggleBooleanCell,
  onPaste,
  subscribeViewport,
  columnWidths: controlledColumnWidths,
  onColumnWidthChange,
  onAutofitColumn,
  onVisibleViewportChange,
}: WorkbookGridSurfaceProps) {
  const emptyGpuScene = useMemo<GridGpuScene>(() => ({ fillRects: [], borderRects: [] }), []);
  const emptyTextScene = useMemo<GridTextScene>(() => ({ items: [] }), []);
  const hostRef = useRef<HTMLDivElement | null>(null);
  const focusTargetRef = useRef<HTMLDivElement | null>(null);
  const scrollViewportRef = useRef<HTMLDivElement | null>(null);
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
  const autoScrollSelectionRef = useRef<{ sheetName: string; col: number; row: number } | null>(
    null,
  );
  const columnResizeActiveRef = useRef(false);
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const pendingTypeSeedRef = useRef<string | null>(null);
  const fillPreviewRangeRef = useRef<Rectangle | null>(null);
  const fillHandleCleanupRef = useRef<(() => void) | null>(null);
  const fillHandlePointerIdRef = useRef<number | null>(null);
  const rangeMoveCleanupRef = useRef<(() => void) | null>(null);
  const rangeMoveSourceRangeRef = useRef<Rectangle | null>(null);
  const rangeMovePreviewRangeRef = useRef<Rectangle | null>(null);
  const rangeMoveAnchorOffsetRef = useRef<Item | null>(null);
  const resizeCleanupRef = useRef<(() => void) | null>(null);
  const scrollSyncFrameRef = useRef<number | null>(null);
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
  const [sceneRevision, setSceneRevision] = useState(0);
  const [fillPreviewRange, setFillPreviewRange] = useState<Rectangle | null>(null);
  const [isFillHandleDragging, setIsFillHandleDragging] = useState(false);
  const [isRangeMoveDragging, setIsRangeMoveDragging] = useState(false);
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
  const gridTheme = useMemo(() => getGridTheme(), []);
  const columnWidths =
    controlledColumnWidths ?? columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS;
  const sortedColumnWidthOverrides = useMemo(
    () =>
      Object.entries(columnWidths)
        .map(([index, width]) => [Number(index), width] as const)
        .toSorted((left, right) => left[0] - right[0]),
    [columnWidths],
  );
  const columnWidthOverridesAttr = useMemo(() => {
    const entries = Object.entries(columnWidths).toSorted(
      ([left], [right]) => Number(left) - Number(right),
    );
    return entries.length === 0 ? "{}" : JSON.stringify(Object.fromEntries(entries));
  }, [columnWidths]);
  const totalGridWidth = useMemo(
    () =>
      gridMetrics.rowMarkerWidth +
      resolveColumnOffset(MAX_COLS, sortedColumnWidthOverrides, gridMetrics.columnWidth),
    [gridMetrics.columnWidth, gridMetrics.rowMarkerWidth, sortedColumnWidthOverrides],
  );
  const totalGridHeight = useMemo(
    () => gridMetrics.headerHeight + MAX_ROWS * gridMetrics.rowHeight,
    [gridMetrics.headerHeight, gridMetrics.rowHeight],
  );
  const scrollLeft = useMemo(
    () =>
      resolveColumnOffset(
        visibleRegion.range.x,
        sortedColumnWidthOverrides,
        gridMetrics.columnWidth,
      ) + visibleRegion.tx,
    [gridMetrics.columnWidth, sortedColumnWidthOverrides, visibleRegion.range.x, visibleRegion.tx],
  );
  const scrollTop = useMemo(
    () => visibleRegion.range.y * gridMetrics.rowHeight + visibleRegion.ty,
    [gridMetrics.rowHeight, visibleRegion.range.y, visibleRegion.ty],
  );

  const getCellLocalBounds = useCallback(
    (col: number, row: number): Rectangle | undefined => {
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return undefined;
      }
      return {
        x:
          gridMetrics.rowMarkerWidth +
          resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
          scrollLeft,
        y: gridMetrics.headerHeight + row * gridMetrics.rowHeight - scrollTop,
        width: getResolvedColumnWidth(columnWidths, col, gridMetrics.columnWidth),
        height: gridMetrics.rowHeight,
      };
    },
    [
      columnWidths,
      gridMetrics.columnWidth,
      gridMetrics.headerHeight,
      gridMetrics.rowHeight,
      gridMetrics.rowMarkerWidth,
      scrollLeft,
      scrollTop,
      sortedColumnWidthOverrides,
    ],
  );

  const getCellScreenBounds = useCallback(
    (col: number, row: number): Rectangle | undefined => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      const localBounds = getCellLocalBounds(col, row);
      if (!hostBounds || !localBounds) {
        return undefined;
      }
      return {
        x: hostBounds.left + localBounds.x,
        y: hostBounds.top + localBounds.y,
        width: localBounds.width,
        height: localBounds.height,
      };
    },
    [getCellLocalBounds],
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
  const invalidateScene = useCallback(() => {
    setSceneRevision((current) => current + 1);
  }, []);

  const syncVisibleRegion = useCallback(() => {
    const scrollViewport = scrollViewportRef.current;
    if (!scrollViewport) {
      return;
    }
    const next = resolveVisibleRegionFromScroll({
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      columnWidths,
      gridMetrics,
    });
    setVisibleRegion((current) => (sameVisibleRegion(current, next) ? current : next));
  }, [columnWidths, gridMetrics]);

  useEffect(() => {
    fillPreviewRangeRef.current = fillPreviewRange;
  }, [fillPreviewRange]);

  useEffect(() => {
    return () => {
      fillHandleCleanupRef.current?.();
      rangeMoveCleanupRef.current?.();
      resizeCleanupRef.current?.();
    };
  }, []);

  useEffect(() => {
    onVisibleViewportChange?.(viewport);
  }, [onVisibleViewportChange, viewport]);

  useEffect(() => {
    const scrollViewport = scrollViewportRef.current;
    if (!scrollViewport) {
      return;
    }

    syncVisibleRegion();
    const scheduleVisibleRegionSync = () => {
      if (scrollSyncFrameRef.current !== null) {
        return;
      }
      scrollSyncFrameRef.current = window.requestAnimationFrame(() => {
        scrollSyncFrameRef.current = null;
        syncVisibleRegion();
      });
    };
    const handleScroll = () => {
      scheduleVisibleRegionSync();
    };
    scrollViewport.addEventListener("scroll", handleScroll, { passive: true });
    const observer = new ResizeObserver(() => {
      scheduleVisibleRegionSync();
    });
    observer.observe(scrollViewport);
    return () => {
      if (scrollSyncFrameRef.current !== null) {
        window.cancelAnimationFrame(scrollSyncFrameRef.current);
        scrollSyncFrameRef.current = null;
      }
      observer.disconnect();
      scrollViewport.removeEventListener("scroll", handleScroll);
    };
  }, [hostElement, syncVisibleRegion]);

  useEffect(() => {
    if (subscribeViewport) {
      return subscribeViewport(sheetName, viewport, invalidateScene);
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene);
  }, [engine, invalidateScene, sheetName, subscribeViewport, viewport, visibleAddresses]);

  const resizeGuideColumn = useMemo(
    () =>
      activeResizeColumn ??
      (hoverState.cursor === "col-resize" && hoverState.header?.kind === "column"
        ? hoverState.header.index
        : null),
    [activeResizeColumn, hoverState.cursor, hoverState.header],
  );

  const gpuScene = useMemo<GridGpuScene>(() => {
    if (!hostElement) {
      return emptyGpuScene;
    }
    // The engine mutates behind a stable object identity, so explicit revision
    // reads are required to force scene recomputation on viewport invalidations.
    void sceneRevision;
    return buildGridGpuScene({
      engine,
      columnWidths,
      fillPreviewRange,
      gridMetrics,
      gridSelection,
      activeHeaderDrag,
      hoveredCell: hoverState.cell,
      hoveredHeader: hoverState.header,
      resizeGuideColumn,
      selectedCell: [selectedCell.col, selectedCell.row],
      selectionRange,
      sheetName,
      visibleItems: deferredVisibleItems,
      visibleRegion: deferredVisibleRegion,
      hostBounds: { left: 0, top: 0 },
      getCellBounds: getCellLocalBounds,
    });
  }, [
    activeHeaderDrag,
    columnWidths,
    deferredVisibleItems,
    deferredVisibleRegion,
    emptyGpuScene,
    engine,
    fillPreviewRange,
    getCellLocalBounds,
    gridMetrics,
    gridSelection,
    hostElement,
    hoverState.cell,
    hoverState.header,
    resizeGuideColumn,
    sceneRevision,
    selectedCell.col,
    selectedCell.row,
    selectionRange,
    sheetName,
  ]);

  const textScene = useMemo<GridTextScene>(() => {
    if (!hostElement) {
      return emptyTextScene;
    }
    // The engine mutates behind a stable object identity, so explicit revision
    // reads are required to force scene recomputation on viewport invalidations.
    void sceneRevision;
    return buildGridTextScene({
      engine,
      columnWidths,
      editingCell: isEditingCell ? ([selectedCell.col, selectedCell.row] as const) : null,
      gridMetrics,
      activeHeaderDrag,
      hoveredHeader: hoverState.header,
      resizeGuideColumn,
      selectedCell: [selectedCell.col, selectedCell.row],
      selectedCellSnapshot,
      selectionRange,
      sheetName,
      visibleItems: deferredVisibleItems,
      visibleRegion: deferredVisibleRegion,
      hostBounds: { left: 0, top: 0 },
      getCellBounds: getCellLocalBounds,
    });
  }, [
    activeHeaderDrag,
    columnWidths,
    deferredVisibleItems,
    deferredVisibleRegion,
    emptyTextScene,
    engine,
    getCellLocalBounds,
    gridMetrics,
    hostElement,
    hoverState.header,
    isEditingCell,
    resizeGuideColumn,
    sceneRevision,
    selectedCellSnapshot,
    selectedCell.col,
    selectedCell.row,
    selectionRange,
    sheetName,
  ]);

  const fillHandleBounds = useMemo<FillHandleOverlayBounds | undefined>(() => {
    if (
      !hostElement ||
      !selectionRange ||
      gridSelection.columns.length > 0 ||
      gridSelection.rows.length > 0 ||
      fillPreviewRange ||
      isRangeMoveDragging
    ) {
      return undefined;
    }
    return resolveFillHandleOverlayBounds({
      sourceRange: selectionRange,
      hostBounds: { left: 0, top: 0 },
      getCellBounds: getCellLocalBounds,
    });
  }, [
    fillPreviewRange,
    getCellLocalBounds,
    gridSelection.columns.length,
    gridSelection.rows.length,
    hostElement,
    isRangeMoveDragging,
    selectionRange,
  ]);

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current;
    if (!scrollViewport) {
      return;
    }
    const previousAutoScrollSelection = autoScrollSelectionRef.current;
    const nextAutoScrollSelection = {
      sheetName,
      col: selectedCell.col,
      row: selectedCell.row,
    };
    const selectionChanged = hasSelectionTargetChanged(
      previousAutoScrollSelection,
      nextAutoScrollSelection,
    );
    if (!selectionChanged) {
      return;
    }
    autoScrollSelectionRef.current = nextAutoScrollSelection;
    scrollCellIntoView({
      cell: [selectedCell.col, selectedCell.row],
      columnWidths,
      gridMetrics,
      scrollViewport,
      sortedColumnWidthOverrides,
    });
  }, [
    columnWidths,
    gridMetrics,
    selectedCell.col,
    selectedCell.row,
    sheetName,
    sortedColumnWidthOverrides,
  ]);

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
    const next = getCellScreenBounds(selectedCell.col, selectedCell.row);
    setOverlayBounds((current) => {
      if (!next) {
        return current;
      }
      return sameBounds(current, next) ? current : next;
    });
  }, [getCellScreenBounds, selectedCell.col, selectedCell.row]);

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
    const focusTarget = focusTargetRef.current;
    if (focusTarget) {
      focusTarget.focus({ preventScroll: true });
      return;
    }
    hostRef.current?.focus({ preventScroll: true });
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
        selectedCellBounds: getCellScreenBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange: gridSelection.current?.range ?? null,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      });
    },
    [
      columnWidths,
      getCellScreenBounds,
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
  const allowsRangeMove = Boolean(
    selectionRange &&
    gridSelection.columns.length === 0 &&
    gridSelection.rows.length === 0 &&
    !fillPreviewRange &&
    !isFillHandleDragging,
  );

  const isFillHandleTarget = useCallback((target: EventTarget | null): boolean => {
    return target instanceof Element && target.closest("[data-grid-fill-handle='true']") !== null;
  }, []);

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

  const editorPresentation = useMemo(() => {
    const selectedCellStyle = engine.getCellStyle(selectedCellSnapshot.styleId);
    const renderCell = snapshotToRenderCell(selectedCellSnapshot, selectedCellStyle);
    return getEditorPresentation({
      renderCell,
      fillColor: selectedCellStyle?.fill?.backgroundColor,
    });
  }, [engine, selectedCellSnapshot]);

  const editorTextAlign = useMemo<"left" | "right">(
    () => getEditorTextAlign(editorValue),
    [editorValue],
  );

  const handleFillHandlePointerDown = useCallback(
    (event: ReactPointerEvent<HTMLButtonElement>) => {
      if (!selectionRange || event.button !== 0) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      focusGrid();
      const handleElement = event.currentTarget;
      fillHandleCleanupRef.current?.();
      fillPreviewRangeRef.current = null;
      setFillPreviewRange(null);
      fillHandlePointerIdRef.current = event.pointerId;
      setIsFillHandleDragging(true);
      setHoverState((current) =>
        sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
          ? current
          : { cell: null, header: null, cursor: "default" },
      );
      handleElement.setPointerCapture(event.pointerId);

      const move = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return;
        }
        const pointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY);
        const nextPreviewRange = pointerCell
          ? resolveFillHandlePreviewRange(selectionRange, pointerCell)
          : null;
        fillPreviewRangeRef.current = nextPreviewRange;
        setFillPreviewRange(nextPreviewRange);
      };

      const cleanup = () => {
        if (fillHandleCleanupRef.current !== cleanup) {
          return;
        }
        fillHandleCleanupRef.current = null;
        handleElement.removeEventListener("pointermove", move);
        handleElement.removeEventListener("pointerup", up);
        handleElement.removeEventListener("pointercancel", cancel);
        handleElement.removeEventListener("lostpointercapture", lostPointerCapture);
        const pointerId = fillHandlePointerIdRef.current;
        fillHandlePointerIdRef.current = null;
        setIsFillHandleDragging(false);
        if (pointerId !== null && handleElement.hasPointerCapture(pointerId)) {
          handleElement.releasePointerCapture(pointerId);
        }
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
            ? current
            : { cell: null, header: null, cursor: "default" },
        );
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

      const up = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return;
        }
        finish();
      };

      const cancel = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return;
        }
        fillPreviewRangeRef.current = null;
        setFillPreviewRange(null);
        cleanup();
      };

      const lostPointerCapture = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return;
        }
        fillPreviewRangeRef.current = null;
        setFillPreviewRange(null);
        cleanup();
      };

      fillHandleCleanupRef.current = cleanup;
      handleElement.addEventListener("pointermove", move);
      handleElement.addEventListener("pointerup", up);
      handleElement.addEventListener("pointercancel", cancel);
      handleElement.addEventListener("lostpointercapture", lostPointerCapture);
    },
    [focusGrid, onFillRange, resolvePointerCell, selectionRange],
  );

  const refreshHoverState = useCallback(
    (clientX: number, clientY: number, buttons: number) => {
      if (isFillHandleDragging) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
            ? current
            : { cell: null, header: null, cursor: "default" },
        );
        return;
      }
      if (isRangeMoveDragging) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "grabbing" })
            ? current
            : { cell: null, header: null, cursor: "grabbing" },
        );
        return;
      }
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
      const rangeMoveAnchorCell = allowsRangeMove
        ? resolveSelectionMoveAnchorCell(clientX, clientY, selectionRange, getCellScreenBounds)
        : null;
      if (rangeMoveAnchorCell) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: "grab" })
            ? current
            : { cell: null, header: null, cursor: "grab" },
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
        selectedCellBounds: getCellScreenBounds(selectedCell.col, selectedCell.row) ?? null,
        selectionRange,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      });
      setHoverState((current) => (sameGridHoverState(current, next) ? current : next));
    },
    [
      columnWidths,
      getCellScreenBounds,
      gridMetrics,
      gridSelection.columns.length,
      gridSelection.rows.length,
      isFillHandleDragging,
      isRangeMoveDragging,
      resolvePointerGeometry,
      allowsRangeMove,
      selectedCell.col,
      selectedCell.row,
      selectionRange,
      visibleRegion,
    ],
  );

  const beginRangeMove = useCallback(
    (pointerCell: Item) => {
      if (!selectionRange) {
        return;
      }
      const sourceRange = selectionRange;
      const anchorOffset: Item = [pointerCell[0] - sourceRange.x, pointerCell[1] - sourceRange.y];
      rangeMoveCleanupRef.current?.();
      rangeMoveSourceRangeRef.current = sourceRange;
      rangeMovePreviewRangeRef.current = sourceRange;
      rangeMoveAnchorOffsetRef.current = anchorOffset;
      setIsRangeMoveDragging(true);
      setHoverState({ cell: null, header: null, cursor: "grabbing" });
      if (isEditingCell) {
        onCommitEdit();
      }
      focusGrid();

      const move = (nativeEvent: PointerEvent) => {
        const nextPointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY);
        if (!nextPointerCell) {
          return;
        }
        const nextRange = resolveMovedRange(sourceRange, nextPointerCell, anchorOffset);
        if (sameRectangle(rangeMovePreviewRangeRef.current, nextRange)) {
          return;
        }
        rangeMovePreviewRangeRef.current = nextRange;
        setGridSelection(createRectangleSelectionFromRange(nextRange));
      };

      const cleanup = (clientX?: number, clientY?: number) => {
        window.removeEventListener("pointermove", move, true);
        window.removeEventListener("pointerup", up, true);
        rangeMoveCleanupRef.current = null;
        rangeMoveSourceRangeRef.current = null;
        rangeMovePreviewRangeRef.current = null;
        rangeMoveAnchorOffsetRef.current = null;
        setIsRangeMoveDragging(false);
        if (clientX !== undefined && clientY !== undefined) {
          refreshHoverState(clientX, clientY, 0);
        }
      };

      const up = (nativeEvent: PointerEvent) => {
        const resolvedSourceRange = rangeMoveSourceRangeRef.current ?? sourceRange;
        const resolvedTargetRange = rangeMovePreviewRangeRef.current ?? resolvedSourceRange;
        cleanup(nativeEvent.clientX, nativeEvent.clientY);
        setGridSelection(createRectangleSelectionFromRange(resolvedTargetRange));
        onSelect(formatAddress(resolvedTargetRange.y, resolvedTargetRange.x));
        if (sameRectangle(resolvedSourceRange, resolvedTargetRange)) {
          return;
        }
        const sourceAddresses = rectangleToAddresses(resolvedSourceRange);
        const targetAddresses = rectangleToAddresses(resolvedTargetRange);
        onMoveRange(
          sourceAddresses.startAddress,
          sourceAddresses.endAddress,
          targetAddresses.startAddress,
          targetAddresses.endAddress,
        );
      };

      rangeMoveCleanupRef.current = cleanup;
      window.addEventListener("pointermove", move, true);
      window.addEventListener("pointerup", up, true);
    },
    [
      focusGrid,
      isEditingCell,
      onCommitEdit,
      onMoveRange,
      onSelect,
      refreshHoverState,
      resolvePointerCell,
      selectionRange,
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

      let measuredWidth = 0;

      context.font = gridTheme.headerFontStyle;
      measuredWidth = Math.max(
        measuredWidth,
        context.measureText(indexToColumn(columnIndex)).width,
      );

      const sheet = engine.workbook.getSheet(sheetName);
      sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (col !== columnIndex) {
          return;
        }
        const snapshot = engine.getCell(sheetName, formatAddress(row, col));
        const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId));
        const displayText = renderCell.displayText || renderCell.copyText;
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

  const beginColumnResize = useCallback(
    (columnIndex: number, startClientX: number) => {
      resizeCleanupRef.current?.();
      startGridResize(interactionState);
      setActiveResizeColumn(columnIndex);
      const startWidth = getResolvedColumnWidth(columnWidths, columnIndex, gridMetrics.columnWidth);

      const handlePointerMove = (nativeEvent: PointerEvent) => {
        applyColumnWidth(columnIndex, startWidth + (nativeEvent.clientX - startClientX));
      };

      const cleanup = (nativeEvent?: PointerEvent) => {
        window.removeEventListener("pointermove", handlePointerMove, true);
        window.removeEventListener("pointerup", handlePointerUp, true);
        resizeCleanupRef.current = null;
        setActiveResizeColumn(null);
        finishGridResize(interactionState);
        if (nativeEvent) {
          refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0);
        }
      };

      const handlePointerUp = (nativeEvent: PointerEvent) => {
        cleanup(nativeEvent);
      };

      resizeCleanupRef.current = cleanup;
      window.addEventListener("pointermove", handlePointerMove, true);
      window.addEventListener("pointerup", handlePointerUp, true);
    },
    [applyColumnWidth, columnWidths, gridMetrics.columnWidth, interactionState, refreshHoverState],
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
    <div className="relative flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]">
      <div
        className="sheet-grid-host min-h-0 flex-1 bg-[var(--wb-surface)] pr-2 pb-2"
        data-column-width-overrides={columnWidthOverridesAttr}
        data-default-column-width={gridMetrics.columnWidth}
        data-testid="sheet-grid"
        role="grid"
        style={{ cursor: hoverState.cursor }}
        onFocus={(event) => {
          if (event.target === event.currentTarget) {
            focusGrid();
          }
        }}
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
          if (isFillHandleDragging || isFillHandleTarget(event.target)) {
            return;
          }
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
          if (activeResizeColumn !== null || isFillHandleDragging || isRangeMoveDragging) {
            return;
          }
          setHoverState((current) =>
            sameGridHoverState(current, { cell: null, header: null, cursor: "default" })
              ? current
              : { cell: null, header: null, cursor: "default" },
          );
        }}
        onPointerDownCapture={(event) => {
          if (isFillHandleTarget(event.target)) {
            return;
          }
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
          if (resizeTarget !== null) {
            event.preventDefault();
            event.stopPropagation();
            if (isEditingCell) {
              onCommitEdit();
            }
            focusGrid();
            setActiveHeaderDrag(null);
            setHoverState((current) =>
              sameGridHoverState(current, {
                cell: null,
                header: { kind: "column", index: resizeTarget },
                cursor: "col-resize",
              })
                ? current
                : {
                    cell: null,
                    header: { kind: "column", index: resizeTarget },
                    cursor: "col-resize",
                  },
            );
            beginColumnResize(resizeTarget, event.clientX);
            return;
          }

          if (allowsRangeMove) {
            const rangeMoveAnchorCell = resolveSelectionMoveAnchorCell(
              event.clientX,
              event.clientY,
              selectionRange,
              getCellScreenBounds,
            );
            if (rangeMoveAnchorCell) {
              event.preventDefault();
              event.stopPropagation();
              resetGridPointerInteraction(interactionState, {
                clearIgnoreNextPointerSelection: true,
              });
              setActiveResizeColumn(null);
              setActiveHeaderDrag(null);
              beginRangeMove(rangeMoveAnchorCell);
              return;
            }
          }

          const headerSelection =
            pointerGeometry === null
              ? null
              : resolveHeaderSelectionAtPointer(
                  event.clientX,
                  event.clientY,
                  visibleRegion,
                  pointerGeometry,
                );
          setActiveResizeColumn(null);
          setActiveHeaderDrag(headerSelection);
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
          if (isRangeMoveDragging) {
            return;
          }
          const clickedCell =
            dragDidMoveRef.current || dragHeaderSelectionRef.current
              ? null
              : resolvePointerCell(event.clientX, event.clientY);
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
          if (clickedCell) {
            toggleBooleanCellAt(clickedCell[0], clickedCell[1]);
          }
          setActiveHeaderDrag(null);
          refreshHoverState(event.clientX, event.clientY, 0);
        }}
        ref={handleHostRef}
        // oxlint-disable-next-line jsx-a11y/no-noninteractive-tabindex
        tabIndex={0}
      >
        <div
          aria-label={`${sheetName} grid focus target`}
          className="pointer-events-none absolute h-px w-px overflow-hidden opacity-0"
          data-testid="sheet-grid-focus-target"
          ref={focusTargetRef}
          tabIndex={-1}
        />
        <div ref={scrollViewportRef} aria-hidden="true" className="absolute inset-0 overflow-auto">
          <div style={{ height: totalGridHeight, width: totalGridWidth }} />
        </div>
        <GridGpuSurface host={hostElement} scene={gpuScene} onActiveChange={setIsWebGpuActive} />
        <GridTextOverlay active={hostElement !== null && isWebGpuActive} scene={textScene} />
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
            className="absolute z-30 cursor-crosshair rounded-full border-0 bg-[#1f7a43] shadow-[0_0_0_1px_rgba(31,122,67,0.45)] outline-none"
            data-grid-fill-handle="true"
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onPointerDown={handleFillHandlePointerDown}
            style={{
              height: fillHandleBounds.height,
              left: fillHandleBounds.x,
              touchAction: "none",
              top: fillHandleBounds.y,
              width: fillHandleBounds.width,
            }}
            type="button"
          />
        ) : null}
        <div className="pointer-events-none absolute inset-0 z-[1]" />
      </div>
      {isEditingCell && overlayStyle ? (
        <CellEditorOverlay
          label={`${sheetName}!${selectedAddr}`}
          onCancel={onCancelEdit}
          onChange={onEditorChange}
          onCommit={onCommitEdit}
          backgroundColor={editorPresentation.backgroundColor}
          color={editorPresentation.color}
          font={editorPresentation.font}
          fontSize={editorPresentation.fontSize}
          resolvedValue={resolvedValue}
          selectionBehavior={editorSelectionBehavior}
          textAlign={editorTextAlign}
          underline={editorPresentation.underline}
          value={editorValue}
          style={overlayStyle}
        />
      ) : null}
    </div>
  );
}

function resolveVisibleRegionFromScroll(options: {
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

function resolveColumnOffset(
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

function scrollCellIntoView(options: {
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
