import React from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, indexToColumn } from "@bilig/formula";
import { useViewport } from "./useViewport.js";

interface SheetGridViewProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  selectedAddr: string;
  onSelect(addr: string): void;
}

export function SheetGridView({ engine, sheetName, selectedAddr, onSelect }: SheetGridViewProps) {
  const viewport = { rowStart: 0, rowEnd: 19, colStart: 0, colEnd: 7 };
  useViewport(engine, sheetName, viewport);
  const rows = Array.from({ length: viewport.rowEnd - viewport.rowStart + 1 }, (_, rowIndex) => rowIndex);
  const cols = Array.from({ length: viewport.colEnd - viewport.colStart + 1 }, (_, colIndex) => colIndex);

  return (
    <div className="sheet-grid">
      <div className="grid-row header">
        <div className="grid-corner" />
        {cols.map((col) => (
          <div className="grid-header" key={`header-${col}`}>
            {indexToColumn(col)}
          </div>
        ))}
      </div>
      {rows.map((row) => (
        <div className="grid-row" key={`row-${row}`}>
          <div className="grid-header">{row + 1}</div>
          {cols.map((col) => {
            const addr = formatAddress(row, col);
            const snapshot = engine.getCell(sheetName, addr);
            const isSelected = addr === selectedAddr;
            const content =
              snapshot.value.tag === 1
                ? snapshot.value.value
                : snapshot.value.tag === 2
                  ? String(snapshot.value.value)
                  : snapshot.value.tag === 3
                    ? snapshot.value.value
                    : snapshot.value.tag === 4
                      ? `#${snapshot.value.code}`
                      : "";
            return (
              <button
                aria-label={`Cell ${addr}`}
                className={isSelected ? "grid-cell selected" : "grid-cell"}
                data-addr={addr}
                key={`${row}-${col}`}
                onClick={() => onSelect(addr)}
                type="button"
              >
                {content}
              </button>
            );
          })}
        </div>
      ))}
    </div>
  );
}
