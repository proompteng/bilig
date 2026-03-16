import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { selectors, type SpreadsheetEngine } from "@bilig/core";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { MAX_COLS, MAX_ROWS, ValueTag } from "@bilig/protocol";
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
  selectedAddr: string;
  editorValue: string;
  resolvedValue: string;
  isEditingCell: boolean;
  onSelect(addr: string): void;
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

const DEFAULT_COLUMN_WIDTH = 120;
const DEFAULT_ROW_HEIGHT = 28;
const HEADER_HEIGHT = 30;
const ROW_MARKER_WIDTH = 60;

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
        data: `#${snapshot.value.code}`,
        displayData: `#${snapshot.value.code}`,
        readonly: false,
        copyData: snapshot.formula ? rawValue : `#${snapshot.value.code}`,
        themeOverride: {
          textDark: "#991b1b",
          bgCell: "#fff4f4"
        }
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
        return `#${snapshot.value.code}`;
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
  selectedAddr,
  editorValue,
  resolvedValue,
  isEditingCell,
  onSelect,
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
  const pendingPointerCellRef = useRef<Item | null>(null);
  const dragAnchorCellRef = useRef<Item | null>(null);
  const dragPointerCellRef = useRef<Item | null>(null);
  const dragViewportRef = useRef<VisibleRegionState | null>(null);
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
      setOverlayBounds((current) => (sameBounds(current, next) ? current : next));
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
    (clientX: number, clientY: number, region: VisibleRegionState = visibleRegion): Item | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      if (!hostBounds) {
        return null;
      }

      const dataLeft = hostBounds.left + ROW_MARKER_WIDTH;
      const dataTop = hostBounds.top + HEADER_HEIGHT + region.ty;
      if (clientX < dataLeft || clientY < dataTop) {
        return null;
      }

      const col = region.range.x + Math.floor((clientX - dataLeft) / DEFAULT_COLUMN_WIDTH);
      const row = region.range.y + Math.floor((clientY - dataTop) / DEFAULT_ROW_HEIGHT);
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return null;
      }

      return [col, row];
    },
    [visibleRegion]
  );

  const selectionSummary = useMemo(
    () => formatSelectionSummary(gridSelection, selectedAddr),
    [gridSelection, selectedAddr]
  );

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
      <div
        className="sheet-grid-host"
        data-testid="sheet-grid"
        onKeyDownCapture={() => {
          pendingPointerCellRef.current = null;
          dragAnchorCellRef.current = null;
          dragPointerCellRef.current = null;
          dragViewportRef.current = null;
        }}
        onKeyDown={(event) => handleGridKey(event)}
        onMouseMoveCapture={(event) => {
          if ((event.buttons & 1) !== 1 || dragAnchorCellRef.current === null) {
            return;
          }
          const pointerCell = resolvePointerCell(event.clientX, event.clientY, dragViewportRef.current ?? visibleRegion);
          if (!pointerCell) {
            return;
          }
          const currentPointer = dragPointerCellRef.current;
          if (currentPointer && currentPointer[0] === pointerCell[0] && currentPointer[1] === pointerCell[1]) {
            return;
          }
          dragPointerCellRef.current = pointerCell;
          setGridSelection(
            createRangeSelection(
              createGridSelection(dragAnchorCellRef.current[0], dragAnchorCellRef.current[1]),
              dragAnchorCellRef.current,
              pointerCell
            )
          );
        }}
        onMouseDownCapture={(event) => {
          if (event.button !== 0) {
            return;
          }
          const pointerCell = resolvePointerCell(event.clientX, event.clientY);
          pendingPointerCellRef.current = pointerCell;
          dragAnchorCellRef.current = pointerCell;
          dragPointerCellRef.current = pointerCell;
          dragViewportRef.current = visibleRegion;
          if (pointerCell) {
            setGridSelection(createGridSelection(pointerCell[0], pointerCell[1]));
            if (isEditingCell) {
              onCancelEdit();
            }
            onSelect(formatAddress(pointerCell[1], pointerCell[0]));
          }
          hostRef.current?.focus();
        }}
        onMouseUpCapture={(event) => {
          const anchorCell = dragAnchorCellRef.current;
          if (!anchorCell) {
            return;
          }

          const pointerCell = resolvePointerCell(event.clientX, event.clientY, dragViewportRef.current ?? visibleRegion)
            ?? dragPointerCellRef.current
            ?? anchorCell;

          dragPointerCellRef.current = pointerCell;
          setGridSelection(
            createRangeSelection(
              createGridSelection(anchorCell[0], anchorCell[1]),
              anchorCell,
              pointerCell
            )
          );
          onSelect(formatAddress(anchorCell[1], anchorCell[0]));

          window.requestAnimationFrame(() => {
            pendingPointerCellRef.current = null;
            dragAnchorCellRef.current = null;
            dragPointerCellRef.current = null;
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
          onGridSelectionChange={(nextSelection) => {
            if (dragViewportRef.current) {
              return;
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
              onCancelEdit();
            }
            onSelect(formatAddress(cell[1], cell[0]));
          }}
          onKeyDown={(event) => {
            if (!isPrintableKey(event) && event.key !== "Enter" && event.key !== "Tab" && event.key !== "F2" && event.key !== "Backspace" && event.key !== "Delete") {
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
          rowMarkers={{ kind: "number", width: ROW_MARKER_WIDTH }}
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
