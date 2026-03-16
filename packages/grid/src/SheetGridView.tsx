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
  const pendingPointerCellRef = useRef<Item | null>(null);
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0
  });
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined);
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName]);

  const columns = useMemo<readonly GridColumn[]>(
    () =>
      Array.from({ length: MAX_COLS }, (_, index) => ({
        id: indexToColumn(index),
        title: indexToColumn(index),
        width: DEFAULT_COLUMN_WIDTH
      })),
    []
  );

  const gridSelection = useMemo<GridSelection>(
    () => ({
      current: {
        cell: [selectedCell.col, selectedCell.row],
        range: { x: selectedCell.col, y: selectedCell.row, width: 1, height: 1 },
        rangeStack: []
      },
      columns: CompactSelection.empty(),
      rows: CompactSelection.empty()
    }),
    [selectedCell.col, selectedCell.row]
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
    (clientX: number, clientY: number): Item | null => {
      const hostBounds = hostRef.current?.getBoundingClientRect();
      if (!hostBounds) {
        return null;
      }

      const dataLeft = hostBounds.left + ROW_MARKER_WIDTH;
      const dataTop = hostBounds.top + HEADER_HEIGHT + visibleRegion.ty;
      if (clientX < dataLeft || clientY < dataTop) {
        return null;
      }

      const col = visibleRegion.range.x + Math.floor((clientX - dataLeft) / DEFAULT_COLUMN_WIDTH);
      const row = visibleRegion.range.y + Math.floor((clientY - dataTop) / DEFAULT_ROW_HEIGHT);
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return null;
      }

      return [col, row];
    },
    [visibleRegion.range.x, visibleRegion.range.y, visibleRegion.ty]
  );

  const handleGridKey = useCallback(
    (event: {
      key: string;
      ctrlKey: boolean;
      metaKey: boolean;
      altKey: boolean;
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
    [beginSelectedEdit, isEditingCell, onClearCell]
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
      if (!withinGridHost && !onDocumentBody) {
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
            {sheetName}!{selectedAddr}
          </span>
          <span>{resolvedValue || "∅"}</span>
        </div>
      </div>
      <div
        className="sheet-grid-host"
        data-testid="sheet-grid"
        onKeyDown={(event) => handleGridKey(event)}
        onMouseDownCapture={(event) => {
          pendingPointerCellRef.current = resolvePointerCell(event.clientX, event.clientY);
          hostRef.current?.focus();
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
            const cell = pendingPointerCellRef.current ?? [col, row];
            pendingPointerCellRef.current = null;
            const addr = formatAddress(cell[1], cell[0]);
            onSelect(addr);
            beginEditAt(addr);
          }}
          onDelete={() => {
            onClearCell();
            return false;
          }}
          onGridSelectionChange={(nextSelection) => {
            const nextCell = nextSelection.current?.cell;
            if (!nextCell) {
              return;
            }
            const cell = pendingPointerCellRef.current ?? nextCell;
            pendingPointerCellRef.current = null;
            if (isEditingCell) {
              onCancelEdit();
            }
            onSelect(formatAddress(cell[1], cell[0]));
          }}
          onKeyDown={(event) => {
            if (!isPrintableKey(event) && event.key !== "Enter" && event.key !== "F2" && event.key !== "Backspace" && event.key !== "Delete") {
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
