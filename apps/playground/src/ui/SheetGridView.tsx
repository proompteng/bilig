import React, { useEffect, useMemo, useRef, useState } from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { useCell } from "./useCell.js";

interface SheetGridViewProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  selectedAddr: string;
  onSelect(addr: string): void;
}

const GRID_ROW_COUNT = 2000;
const GRID_COL_COUNT = 52;
const ROW_HEIGHT = 42;
const COL_WIDTH = 120;
const HEADER_HEIGHT = 46;
const ROW_HEADER_WIDTH = 72;
const OVERSCAN = 3;

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function cellContent(snapshot: ReturnType<typeof useCell>) {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return String(snapshot.value.value);
    case ValueTag.Boolean:
      return String(snapshot.value.value);
    case ValueTag.String:
      return snapshot.value.value;
    case ValueTag.Error:
      return `#${snapshot.value.code}`;
    default:
      return "";
  }
}

interface GridCellProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  addr: string;
  row: number;
  col: number;
  isSelected: boolean;
  onSelect(addr: string): void;
}

const GridCell = React.memo(function GridCell({
  engine,
  sheetName,
  addr,
  row,
  col,
  isSelected,
  onSelect
}: GridCellProps) {
  const snapshot = useCell(engine, sheetName, addr);

  return (
    <button
      aria-label={`Cell ${addr}`}
      aria-selected={isSelected}
      className={isSelected ? "grid-cell selected" : "grid-cell"}
      data-addr={addr}
      data-selected={isSelected ? "true" : "false"}
      onClick={() => onSelect(addr)}
      style={{
        left: ROW_HEADER_WIDTH + col * COL_WIDTH,
        top: HEADER_HEIGHT + row * ROW_HEIGHT,
        width: COL_WIDTH,
        height: ROW_HEIGHT
      }}
      type="button"
    >
      <span className="grid-cell-value">{cellContent(snapshot)}</span>
    </button>
  );
});

export function SheetGridView({ engine, sheetName, selectedAddr, onSelect }: SheetGridViewProps) {
  const scrollerRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportSize, setViewportSize] = useState({ width: 960, height: 640 });

  useEffect(() => {
    const node = scrollerRef.current;
    if (!node) return;

    const resizeObserver = new ResizeObserver((entries) => {
      const nextEntry = entries[0];
      if (!nextEntry) return;
      const { width, height } = nextEntry.contentRect;
      setViewportSize((current) => {
        if (current.width === width && current.height === height) {
          return current;
        }
        return { width, height };
      });
    });

    resizeObserver.observe(node);
    setViewportSize({ width: node.clientWidth, height: node.clientHeight });

    return () => {
      resizeObserver.disconnect();
    };
  }, []);

  const visibleRowCount = Math.max(1, Math.ceil(Math.max(0, viewportSize.height - HEADER_HEIGHT) / ROW_HEIGHT));
  const visibleColCount = Math.max(1, Math.ceil(Math.max(0, viewportSize.width - ROW_HEADER_WIDTH) / COL_WIDTH));

  const viewport = useMemo(() => {
    const rowStart = clamp(Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN, 0, GRID_ROW_COUNT - 1);
    const colStart = clamp(Math.floor(scrollLeft / COL_WIDTH) - OVERSCAN, 0, GRID_COL_COUNT - 1);

    return {
      rowStart,
      rowEnd: clamp(rowStart + visibleRowCount + OVERSCAN * 2, rowStart, GRID_ROW_COUNT - 1),
      colStart,
      colEnd: clamp(colStart + visibleColCount + OVERSCAN * 2, colStart, GRID_COL_COUNT - 1)
    };
  }, [scrollLeft, scrollTop, visibleColCount, visibleRowCount]);

  const rows = useMemo(
    () => Array.from({ length: viewport.rowEnd - viewport.rowStart + 1 }, (_, rowIndex) => viewport.rowStart + rowIndex),
    [viewport]
  );
  const cols = useMemo(
    () => Array.from({ length: viewport.colEnd - viewport.colStart + 1 }, (_, colIndex) => viewport.colStart + colIndex),
    [viewport]
  );
  const selectedCell = parseCellAddress(selectedAddr, sheetName);

  useEffect(() => {
    const node = scrollerRef.current;
    if (!node) return;

    const cellTop = HEADER_HEIGHT + selectedCell.row * ROW_HEIGHT;
    const cellBottom = cellTop + ROW_HEIGHT;
    const cellLeft = ROW_HEADER_WIDTH + selectedCell.col * COL_WIDTH;
    const cellRight = cellLeft + COL_WIDTH;

    let nextScrollTop = node.scrollTop;
    let nextScrollLeft = node.scrollLeft;

    if (cellTop < node.scrollTop + HEADER_HEIGHT) {
      nextScrollTop = Math.max(0, cellTop - HEADER_HEIGHT);
    } else if (cellBottom > node.scrollTop + node.clientHeight) {
      nextScrollTop = cellBottom - node.clientHeight;
    }

    if (cellLeft < node.scrollLeft + ROW_HEADER_WIDTH) {
      nextScrollLeft = Math.max(0, cellLeft - ROW_HEADER_WIDTH);
    } else if (cellRight > node.scrollLeft + node.clientWidth) {
      nextScrollLeft = cellRight - node.clientWidth;
    }

    if (nextScrollTop !== node.scrollTop || nextScrollLeft !== node.scrollLeft) {
      node.scrollTo({ top: nextScrollTop, left: nextScrollLeft });
    }
  }, [selectedCell.col, selectedCell.row, sheetName]);

  const totalCanvasWidth = ROW_HEADER_WIDTH + GRID_COL_COUNT * COL_WIDTH;
  const totalCanvasHeight = HEADER_HEIGHT + GRID_ROW_COUNT * ROW_HEIGHT;

  return (
    <div className="sheet-grid-panel">
      <div className="sheet-grid-toolbar">
        <div>
          <p className="panel-eyebrow">Viewport</p>
          <strong>{sheetName}</strong>
        </div>
        <div className="viewport-meta">
          <span>{GRID_ROW_COUNT.toLocaleString()} rows</span>
          <span>{GRID_COL_COUNT} columns</span>
        </div>
      </div>
      <div
        className="sheet-grid-shell"
        onKeyDown={(event) => {
          const rowDelta =
            event.key === "ArrowDown" || event.key === "Enter"
              ? 1
              : event.key === "ArrowUp"
                ? -1
                : 0;
          const colDelta =
            event.key === "ArrowRight" || event.key === "Tab"
              ? 1
              : event.key === "ArrowLeft"
                ? -1
                : 0;

          if (event.key === "Home") {
            event.preventDefault();
            onSelect(formatAddress(selectedCell.row, 0));
            return;
          }

          if (event.key === "End") {
            event.preventDefault();
            onSelect(formatAddress(selectedCell.row, GRID_COL_COUNT - 1));
            return;
          }

          if (rowDelta === 0 && colDelta === 0) {
            return;
          }

          event.preventDefault();
          const nextRow = clamp(
            selectedCell.row + (event.shiftKey && event.key === "Enter" ? -1 : rowDelta),
            0,
            GRID_ROW_COUNT - 1
          );
          const nextCol = clamp(
            selectedCell.col + (event.shiftKey && event.key === "Tab" ? -1 : colDelta),
            0,
            GRID_COL_COUNT - 1
          );
          onSelect(formatAddress(nextRow, nextCol));
        }}
      >
        <div
          aria-colcount={GRID_COL_COUNT}
          aria-label={`${sheetName} grid`}
          aria-rowcount={GRID_ROW_COUNT}
          className="sheet-grid-scroller"
          data-testid="sheet-grid"
          onScroll={(event) => {
            const target = event.currentTarget;
            setScrollTop(target.scrollTop);
            setScrollLeft(target.scrollLeft);
          }}
          ref={scrollerRef}
          role="grid"
          tabIndex={0}
        >
          <div className="sheet-grid-canvas" style={{ height: totalCanvasHeight, width: totalCanvasWidth }}>
            {rows.flatMap((row) =>
              cols.map((col) => {
                const addr = formatAddress(row, col);
                const isSelected = addr === selectedAddr;

                return (
                  <GridCell
                    addr={addr}
                    col={col}
                    engine={engine}
                    isSelected={isSelected}
                    key={`${row}-${col}`}
                    onSelect={onSelect}
                    row={row}
                    sheetName={sheetName}
                  />
                );
              })
            )}
          </div>
        </div>
        <div className="grid-corner" />
        <div className="column-headers" style={{ left: ROW_HEADER_WIDTH }}>
          <div className="column-header-track" style={{ transform: `translateX(${-scrollLeft}px)` }}>
            {cols.map((col) => (
              <div
                className="grid-header"
                key={`header-${col}`}
                style={{ left: col * COL_WIDTH, width: COL_WIDTH, height: HEADER_HEIGHT }}
              >
                {indexToColumn(col)}
              </div>
            ))}
          </div>
        </div>
        <div className="row-headers" style={{ top: HEADER_HEIGHT }}>
          <div className="row-header-track" style={{ transform: `translateY(${-scrollTop}px)` }}>
            {rows.map((row) => (
              <div
                className={row === selectedCell.row ? "grid-row-header active" : "grid-row-header"}
                key={`row-header-${row}`}
                style={{ top: row * ROW_HEIGHT, width: ROW_HEADER_WIDTH, height: ROW_HEIGHT }}
              >
                {row + 1}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
