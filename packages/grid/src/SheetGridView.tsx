import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
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
  EMPTY_COLUMN_WIDTHS,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  getGridMetrics,
  getResolvedColumnWidth,
  type GridRect,
} from "./gridMetrics.js";
import {
  createColumnSelection,
  createColumnSliceSelection,
  createGridSelection,
  createRangeSelection,
  createRowSelection,
  createRowSliceSelection,
  createSheetSelection,
  formatSelectionSummary,
  rectangleToAddresses,
} from "./gridSelection.js";
import { resolveActivatedCell, resolveSelectionChange } from "./gridSelectionSync.js";
import {
  createPointerGeometry,
  resolveColumnResizeTarget,
  resolveHeaderSelection as resolveHeaderSelectionFromGeometry,
  resolveHeaderSelectionForDrag as resolveHeaderSelectionForDragFromGeometry,
  resolvePointerCell as resolvePointerCellFromGeometry,
  type HeaderSelection,
  type PointerGeometry,
  type VisibleRegionState,
} from "./gridPointer.js";
import {
  resolveBodyDragSelection,
  resolveBodyPointerUpResult,
  resolveHeaderDragSelection,
} from "./gridDragSelection.js";
import { parseClipboardPlainText } from "./gridClipboard.js";
import { cellToEditorSeed, cellToGridCell, getResolvedCellFontFamily } from "./gridCells.js";
import {
  isClipboardShortcut,
  isHandledGridKey,
  isNavigationKey,
  isPrintableKey,
  normalizeKeyboardKey,
} from "./gridKeyboard.js";
import { resolveGridKeyAction } from "./gridKeyActions.js";
import { getEditorTextAlign, getGridTheme, getOverlayStyle } from "./gridPresentation.js";
import {
  resolveBodyDoubleClickIntent,
  resolveHeaderClickIntent,
  shouldSkipGridSelectionChange,
} from "./gridEventPolicy.js";
import {
  buildInternalClipboardRange,
  matchesInternalClipboardPaste,
  type InternalClipboardRange,
} from "./gridInternalClipboard.js";
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

interface BorderOverlaySegment {
  key: string;
  style: CSSProperties;
}

function isCellEditorInputFocused(): boolean {
  if (typeof document === "undefined") {
    return false;
  }
  const activeElement = document.activeElement;
  return (
    activeElement instanceof HTMLInputElement &&
    activeElement.dataset["testid"] === "cell-editor-input"
  );
}

function isEditableElement(element: Element | null): element is HTMLElement {
  return element instanceof HTMLElement && element.isContentEditable;
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
        setBorderOverlayRevision((current) => current + 1);
      });
    }
    return engine.subscribeCells(sheetName, visibleAddresses, () => {
      editorRef.current?.updateCells(visibleDamage);
      setBorderOverlayRevision((current) => current + 1);
    });
  }, [engine, sheetName, subscribeViewport, viewport, visibleAddresses, visibleDamage]);

  useLayoutEffect(() => {
    const hostBounds = hostRef.current?.getBoundingClientRect();
    const editor = editorRef.current;
    if (!hostBounds || !editor) {
      setBorderSegments([]);
      return;
    }

    const segments = new Map<string, BorderOverlaySegment>();
    for (const [col, row] of visibleItems) {
      const snapshot = engine.getCell(sheetName, formatAddress(row, col));
      const style = engine.getCellStyle(snapshot.styleId);
      if (!style?.borders) {
        continue;
      }
      const bounds = editor.getBounds(col, row);
      if (!bounds) {
        continue;
      }

      const rect = {
        x: bounds.x - hostBounds.left,
        y: bounds.y - hostBounds.top,
        width: bounds.width,
        height: bounds.height,
      };

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
        const descriptor = createBorderOverlayDescriptor(rect, side, border);
        if (!descriptor) {
          continue;
        }
        segments.set(descriptor.key, descriptor.segment);
      }
    }

    setBorderSegments([...segments.values()]);
  }, [borderOverlayRevision, engine, sheetName, visibleItems, visibleRegion]);

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
      if (values.length === 0 || values[0]?.length === 0) {
        return;
      }

      const internalClipboard = internalClipboardRef.current;
      if (matchesInternalClipboardPaste(internalClipboard, values)) {
        if (!internalClipboard) {
          return;
        }
        onCopyRange(
          internalClipboard.sourceStartAddress,
          internalClipboard.sourceEndAddress,
          formatAddress(target[1], target[0]),
          formatAddress(
            target[1] + internalClipboard.rowCount - 1,
            target[0] + internalClipboard.colCount - 1,
          ),
        );
        return;
      }

      onPaste(sheetName, formatAddress(target[1], target[0]), values);
    },
    [onCopyRange, onPaste, sheetName],
  );

  const captureInternalClipboardSelection = useCallback(() => {
    const range = gridSelection.current?.range;
    if (!range || gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
      internalClipboardRef.current = null;
      return;
    }

    const values = Array.from({ length: range.height }, (_rowEntry, rowOffset) =>
      Array.from({ length: range.width }, (_colEntry, colOffset) =>
        cellToEditorSeed(
          engine.getCell(sheetName, formatAddress(range.y + rowOffset, range.x + colOffset)),
        ),
      ),
    );

    internalClipboardRef.current = buildInternalClipboardRange(range, values);
  }, [engine, gridSelection, sheetName]);

  useEffect(() => {
    onSelectionLabelChange?.(selectionSummary);
  }, [onSelectionLabelChange, selectionSummary]);

  const currentSelectionCellCol = gridSelection.current?.cell?.[0] ?? null;
  const currentSelectionCellRow = gridSelection.current?.cell?.[1] ?? null;
  const currentSelectionRangeX = gridSelection.current?.range?.x ?? null;
  const currentSelectionRangeY = gridSelection.current?.range?.y ?? null;
  const currentSelectionRangeWidth = gridSelection.current?.range?.width ?? null;
  const currentSelectionRangeHeight = gridSelection.current?.range?.height ?? null;

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
      const currentSelectionCell: [number, number] | null =
        currentSelectionCellCol === null || currentSelectionCellRow === null
          ? null
          : [currentSelectionCellCol, currentSelectionCellRow];
      const currentSelectionRange =
        currentSelectionRangeX === null ||
        currentSelectionRangeY === null ||
        currentSelectionRangeWidth === null ||
        currentSelectionRangeHeight === null
          ? null
          : {
              x: currentSelectionRangeX,
              y: currentSelectionRangeY,
              width: currentSelectionRangeWidth,
              height: currentSelectionRangeHeight,
            };

      const action = resolveGridKeyAction({
        event,
        isEditingCell,
        editorValue,
        editorInputFocused: isCellEditorInputFocused(),
        pendingTypeSeed: pendingTypeSeedRef.current,
        selectedCell: [selectedCell.col, selectedCell.row],
        currentSelectionCell,
        currentRangeAnchor: currentSelectionCell,
        currentSelectionRange,
      });

      if (action.kind === "none") {
        return;
      }

      event.preventDefault();
      event.cancel?.();

      switch (action.kind) {
        case "edit-append":
          onEditorChange(action.value);
          return;
        case "commit-edit":
          onCommitEdit(action.movement);
          return;
        case "cancel-edit":
          onCancelEdit();
          return;
        case "begin-edit":
          pendingTypeSeedRef.current = action.pendingTypeSeed;
          beginSelectedEdit(action.seed, action.selectionBehavior);
          return;
        case "move-selection":
          setGridSelection(createGridSelection(action.cell[0], action.cell[1]));
          onSelect(formatAddress(action.cell[1], action.cell[0]));
          return;
        case "extend-selection":
          setGridSelection(
            createRangeSelection(
              createGridSelection(action.anchor[0], action.anchor[1]),
              action.anchor,
              action.target,
            ),
          );
          return;
        case "clear-cell":
          pendingTypeSeedRef.current = action.pendingTypeSeed;
          onClearCell();
          return;
        case "clipboard-copy": {
          captureInternalClipboardSelection();
          const clipboard = internalClipboardRef.current;
          if (clipboard && typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
            void navigator.clipboard.writeText(clipboard.plainText).catch(() => {});
          }
          return;
        }
        case "clipboard-cut":
          captureInternalClipboardSelection();
          return;
        case "clipboard-paste": {
          pendingKeyboardPasteSequenceRef.current += 1;
          const sequence = pendingKeyboardPasteSequenceRef.current;
          if (typeof navigator !== "undefined" && navigator.clipboard?.readText) {
            void navigator.clipboard
              .readText()
              .then((rawText) => {
                if (pendingKeyboardPasteSequenceRef.current !== sequence) {
                  return undefined;
                }
                pendingKeyboardPasteSequenceRef.current = 0;
                const values = parseClipboardPlainText(rawText);
                applyClipboardValues(action.target, values);
                suppressNextNativePasteRef.current = true;
                return undefined;
              })
              .catch(() => {
                if (pendingKeyboardPasteSequenceRef.current === sequence) {
                  pendingKeyboardPasteSequenceRef.current = 0;
                }
                return undefined;
              });
          }
          return;
        }
        case "select-row":
          setGridSelection(createRowSelection(action.col, action.row));
          return;
        case "select-column":
          setGridSelection(createColumnSelection(action.col, action.row));
          return;
        case "select-all":
          setGridSelection(createSheetSelection());
          return;
      }
    },
    [
      applyClipboardValues,
      beginSelectedEdit,
      captureInternalClipboardSelection,
      currentSelectionCellCol,
      currentSelectionCellRow,
      currentSelectionRangeHeight,
      currentSelectionRangeWidth,
      currentSelectionRangeX,
      currentSelectionRangeY,
      editorValue,
      isEditingCell,
      onCancelEdit,
      onClearCell,
      onCommitEdit,
      onEditorChange,
      onSelect,
      selectedCell.col,
      selectedCell.row,
    ],
  );

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      const normalizedKey = normalizeKeyboardKey(event.key, event.code);
      const activeElement = document.activeElement;
      if (
        activeElement instanceof HTMLInputElement ||
        activeElement instanceof HTMLTextAreaElement ||
        activeElement instanceof HTMLSelectElement ||
        isEditableElement(activeElement)
      ) {
        return;
      }

      const withinGridHost = Boolean(activeElement && hostRef.current?.contains(activeElement));
      const onDocumentBody =
        activeElement === document.body ||
        activeElement === document.documentElement ||
        activeElement === null;
      if (withinGridHost) {
        return;
      }
      if (!onDocumentBody) {
        return;
      }

      if (
        !isHandledGridKey({
          altKey: event.altKey,
          ctrlKey: event.ctrlKey,
          key: normalizedKey,
          metaKey: event.metaKey,
          shiftKey: event.shiftKey,
        })
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
        const snapshot = engine.getCell(sheetName, formatAddress(row, col));
        const style = engine.getCellStyle(snapshot.styleId);
        const displayText =
          "displayData" in cell
            ? String(cell.displayData ?? "")
            : "copyData" in cell
              ? String(cell.copyData ?? "")
              : "";
        context.font = `400 ${gridTheme.editorFontSize} ${getResolvedCellFontFamily(style?.font?.family)}`;
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
      drawContent();
      const snapshot = engine.getCell(sheetName, formatAddress(args.row, args.col));
      const style = engine.getCellStyle(snapshot.styleId);
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
    <div className="relative flex min-h-0 flex-1 flex-col bg-white" data-testid="sheet-grid-shell">
      <div
        className="sheet-grid-host min-h-0 flex-1 bg-white pr-2 pb-2 [--gdg-accent-color:#1a73e8]"
        data-column-width-overrides={JSON.stringify(columnWidths)}
        data-default-column-width={String(gridMetrics.columnWidth)}
        data-testid="sheet-grid"
        role="grid"
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
          if (suppressNextNativePasteRef.current) {
            suppressNextNativePasteRef.current = false;
            event.preventDefault();
            event.stopPropagation();
            return;
          }
          const rawText = event.clipboardData?.getData("text/plain") ?? "";
          const values = parseClipboardPlainText(rawText);
          if (values.length === 0 || values[0]?.length === 0) {
            return;
          }
          if (pendingKeyboardPasteSequenceRef.current !== 0) {
            pendingKeyboardPasteSequenceRef.current = 0;
          }

          const target = gridSelection.current?.cell ?? [selectedCell.col, selectedCell.row];
          applyClipboardValues(target, values);
          event.preventDefault();
          event.stopPropagation();
        }}
        onKeyDown={(event) => handleGridKey(event)}
        onDoubleClickCapture={(event) => {
          const activeGeometry = resolvePointerGeometry(visibleRegion);
          if (!activeGeometry) {
            return;
          }
          const doubleClickIntent = resolveBodyDoubleClickIntent({
            resizeTarget: resolveColumnResizeTarget(
              event.clientX,
              event.clientY,
              visibleRegion,
              activeGeometry,
              columnWidths,
              gridMetrics.columnWidth,
            ),
            bodyCell: resolvePointerCell(
              event.clientX,
              event.clientY,
              visibleRegion,
              activeGeometry,
            ),
            lastBodyClickCell: lastBodyClickCellRef.current,
          });
          if (doubleClickIntent.kind === "ignore") {
            return;
          }
          event.preventDefault();
          event.stopPropagation();
          if (doubleClickIntent.kind === "edit-cell") {
            const editAddress = formatAddress(doubleClickIntent.cell[1], doubleClickIntent.cell[0]);
            setGridSelection(
              createGridSelection(doubleClickIntent.cell[0], doubleClickIntent.cell[1]),
            );
            onSelect(editAddress);
            beginEditAt(editAddress);
            return;
          }
          columnResizeActiveRef.current = false;
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragHeaderSelectionRef.current = null;
          dragGeometryRef.current = null;
          dragDidMoveRef.current = false;
          dragViewportRef.current = null;
          postDragSelectionExpiryRef.current = 0;
          if (onAutofitColumn) {
            void Promise.resolve(
              onAutofitColumn(
                doubleClickIntent.columnIndex,
                computeAutofitColumnWidth(doubleClickIntent.columnIndex),
              ),
            );
            return;
          }
          applyColumnWidth(
            doubleClickIntent.columnIndex,
            computeAutofitColumnWidth(doubleClickIntent.columnIndex),
          );
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
              dragGeometryRef.current,
            );
            if (!nextHeader || nextHeader.index === headerAnchor.index) {
              return;
            }
            dragPointerCellRef.current = null;
            dragDidMoveRef.current = true;
            setGridSelection(
              resolveHeaderDragSelection(headerAnchor, nextHeader.index, [
                selectedCell.col,
                selectedCell.row,
              ]).selection,
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
            dragGeometryRef.current,
          );
          if (!pointerCell) {
            return;
          }
          const currentPointer = dragPointerCellRef.current;
          if (
            currentPointer &&
            currentPointer[0] === pointerCell[0] &&
            currentPointer[1] === pointerCell[1]
          ) {
            return;
          }
          dragPointerCellRef.current = pointerCell;
          if (
            pointerCell[0] !== dragAnchorCellRef.current[0] ||
            pointerCell[1] !== dragAnchorCellRef.current[1]
          ) {
            dragDidMoveRef.current = true;
            ignoreNextPointerSelectionRef.current = false;
            setGridSelection(resolveBodyDragSelection(dragAnchorCellRef.current, pointerCell));
          }
        }}
        onPointerDownCapture={(event) => {
          if (event.button !== 0) {
            return;
          }
          const activeGeometry = resolvePointerGeometry(visibleRegion);
          if (
            activeGeometry &&
            resolveColumnResizeTarget(
              event.clientX,
              event.clientY,
              visibleRegion,
              activeGeometry,
              columnWidths,
              gridMetrics.columnWidth,
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
              setGridSelection(
                createRowSliceSelection(
                  selectedCell.col,
                  headerSelection.index,
                  headerSelection.index,
                ),
              );
              onSelect(formatAddress(headerSelection.index, selectedCell.col));
              focusGrid();
              return;
            }
            ignoreNextPointerSelectionRef.current = true;
            setGridSelection(
              createColumnSliceSelection(
                headerSelection.index,
                headerSelection.index,
                selectedCell.row,
              ),
            );
            onSelect(formatAddress(selectedCell.row, headerSelection.index));
            focusGrid();
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
            const finalHeader =
              resolveHeaderSelectionForPointerDrag(
                headerAnchor.kind,
                event.clientX,
                event.clientY,
                dragViewportRef.current ?? visibleRegion,
                dragGeometryRef.current,
              ) ?? headerAnchor;
            const resolvedHeaderDrag = resolveHeaderDragSelection(headerAnchor, finalHeader.index, [
              selectedCell.col,
              selectedCell.row,
            ]);
            setGridSelection(resolvedHeaderDrag.selection);
            onSelect(resolvedHeaderDrag.addr);
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
            const pointerCell =
              resolvePointerCell(
                event.clientX,
                event.clientY,
                dragViewportRef.current ?? visibleRegion,
                dragGeometryRef.current,
              ) ??
              dragPointerCellRef.current ??
              anchorCell;
            const pointerUpResult = resolveBodyPointerUpResult(anchorCell, pointerCell, true);
            postDragSelectionExpiryRef.current = pointerUpResult.shouldSetDragExpiry
              ? window.performance.now() + 200
              : 0;
            if (pointerUpResult.selection) {
              setGridSelection(pointerUpResult.selection);
            }
            if (pointerUpResult.addr) {
              onSelect(pointerUpResult.addr);
            }
            lastBodyClickCellRef.current = pointerUpResult.clickedCell;
          } else {
            const pointerUpResult = resolveBodyPointerUpResult(
              anchorCell,
              dragPointerCellRef.current ?? anchorCell,
              false,
            );
            lastBodyClickCellRef.current = pointerUpResult.clickedCell;
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
            columnResizeActiveRef.current = true;
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
            dragHeaderSelectionRef.current = null;
            dragGeometryRef.current = null;
            dragDidMoveRef.current = false;
            dragViewportRef.current = null;
            postDragSelectionExpiryRef.current = 0;
          }}
          onColumnResize={(_column: GridColumn, newSize: number, columnIndex: number) => {
            applyColumnWidth(columnIndex, newSize);
          }}
          onColumnResizeEnd={(_column: GridColumn, newSize: number, columnIndex: number) => {
            applyColumnWidth(columnIndex, newSize);
            window.requestAnimationFrame(() => {
              columnResizeActiveRef.current = false;
            });
          }}
          onCellActivated={([col, row]) => {
            const cell = resolveActivatedCell(
              [col, row],
              dragAnchorCellRef.current,
              pendingPointerCellRef.current,
            );
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
            const headerClickIntent = resolveHeaderClickIntent({
              isEdge: event.isEdge,
              isDoubleClick: Boolean(event.isDoubleClick),
              columnResizeActive: columnResizeActiveRef.current,
              columnIndex: col,
              selectedRow: selectedCell.row,
            });
            if (headerClickIntent.kind === "ignore") {
              return;
            }
            if (headerClickIntent.kind === "autofit-column") {
              applyColumnWidth(
                headerClickIntent.columnIndex,
                computeAutofitColumnWidth(headerClickIntent.columnIndex),
              );
              columnResizeActiveRef.current = false;
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
            setGridSelection(
              createColumnSelection(headerClickIntent.columnIndex, headerClickIntent.selectedRow),
            );
            onSelect(headerClickIntent.addr);
            focusGrid();
          }}
          onGridSelectionChange={(nextSelection) => {
            const selectionPolicy = shouldSkipGridSelectionChange({
              columnResizeActive: columnResizeActiveRef.current,
              postDragSelectionExpiry: postDragSelectionExpiryRef.current,
              now: window.performance.now(),
              ignoreNextPointerSelection: ignoreNextPointerSelectionRef.current,
              hasDragViewport: dragViewportRef.current !== null,
            });
            if (selectionPolicy.clearPostDragSelectionExpiry) {
              postDragSelectionExpiryRef.current = 0;
            }
            if (selectionPolicy.consumeIgnoreNextPointerSelection) {
              ignoreNextPointerSelectionRef.current = false;
            }
            if (selectionPolicy.skip) {
              return;
            }
            const resolvedSelection = resolveSelectionChange({
              nextSelection,
              anchorCell: dragAnchorCellRef.current ?? pendingPointerCellRef.current,
              pointerCell: dragPointerCellRef.current ?? pendingPointerCellRef.current,
              selectedCell: [selectedCell.col, selectedCell.row],
            });
            if (!resolvedSelection) {
              return;
            }
            setGridSelection(resolvedSelection.selection);
            if (isEditingCell) {
              onCommitEdit();
            }
            onSelect(resolvedSelection.addr);
          }}
          onKeyDown={(event) => {
            if (
              !isPrintableKey(event) &&
              !isClipboardShortcut(event) &&
              !isNavigationKey(event.key) &&
              event.key !== "Enter" &&
              event.key !== "Tab" &&
              event.key !== "F2" &&
              event.key !== "Backspace" &&
              event.key !== "Delete"
            ) {
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
    getResolvedCellFontFamily(style.font?.family),
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

function createBorderOverlayDescriptor(
  rect: Pick<Rectangle, "x" | "y" | "width" | "height">,
  side: "top" | "right" | "bottom" | "left",
  border: NonNullable<NonNullable<CellStyleRecord["borders"]>["top"]>,
): { key: string; segment: BorderOverlaySegment } | null {
  const thickness = border.weight === "thick" ? 3 : border.weight === "medium" ? 2 : 1;
  const isHorizontal = side === "top" || side === "bottom";
  // Glide cell bounds include the trailing shared gridline pixel. Canonicalize right/bottom
  // edges onto that shared line so adjacent cells collapse into one rendered border segment.
  const edgeX = side === "left" ? rect.x : side === "right" ? rect.x + rect.width - 1 : rect.x;
  const edgeY = side === "top" ? rect.y : side === "bottom" ? rect.y + rect.height - 1 : rect.y;
  const length = isHorizontal ? rect.width : rect.height;
  const offset = thickness / 2;
  const style: CSSProperties = {
    position: "absolute",
    pointerEvents: "none",
    backgroundColor: border.color,
    left: isHorizontal ? edgeX : edgeX - offset,
    top: isHorizontal ? edgeY - offset : edgeY,
    width: isHorizontal ? length : thickness,
    height: isHorizontal ? thickness : length,
  };

  if (border.style === "dashed" || border.style === "dotted") {
    style.backgroundColor = "transparent";
    style.backgroundImage = isHorizontal
      ? `repeating-linear-gradient(90deg, ${border.color} 0 ${border.style === "dashed" ? 6 : 1}px, transparent ${border.style === "dashed" ? 6 : 1}px ${border.style === "dashed" ? 10 : 4}px)`
      : `repeating-linear-gradient(180deg, ${border.color} 0 ${border.style === "dashed" ? 6 : 1}px, transparent ${border.style === "dashed" ? 6 : 1}px ${border.style === "dashed" ? 10 : 4}px)`;
  }

  if (border.style === "double") {
    style.backgroundColor = "transparent";
    style.backgroundImage = isHorizontal
      ? `linear-gradient(to bottom, ${border.color} 0 1px, transparent 1px calc(100% - 1px), ${border.color} calc(100% - 1px) 100%)`
      : `linear-gradient(to right, ${border.color} 0 1px, transparent 1px calc(100% - 1px), ${border.color} calc(100% - 1px) 100%)`;
    style.height = isHorizontal ? Math.max(3, thickness + 2) : length;
    style.width = isHorizontal ? length : Math.max(3, thickness + 2);
    style.left = isHorizontal ? edgeX : edgeX - Math.max(3, thickness + 2) / 2;
    style.top = isHorizontal ? edgeY - Math.max(3, thickness + 2) / 2 : edgeY;
  }

  const left = Number(style.left);
  const top = Number(style.top);
  const width = Number(style.width);
  const height = Number(style.height);
  if (
    !Number.isFinite(left) ||
    !Number.isFinite(top) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    return null;
  }

  const key = [
    Math.round(edgeX * 100) / 100,
    Math.round(edgeY * 100) / 100,
    Math.round(length * 100) / 100,
    isHorizontal ? "h" : "v",
    border.style,
    border.weight ?? "thin",
    border.color ?? "#111827",
  ].join(":");

  return {
    key,
    segment: {
      key,
      style,
    },
  };
}
