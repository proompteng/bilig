import React from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { SheetGridView } from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: SpreadsheetEngine;
  sheetName: string;
  selectedAddr: string;
  onSelect(addr: string): void;
}

export function WorkbookView({ engine, sheetName, selectedAddr, onSelect }: WorkbookViewProps) {
  return (
    <div className="panel workbook-panel">
      <div className="sheet-tabs">
        {[...engine.workbook.sheetsByName.keys()].map((name) => (
          <span className={name === sheetName ? "sheet-tab active" : "sheet-tab"} key={name}>
            {name}
          </span>
        ))}
      </div>
      <SheetGridView engine={engine} sheetName={sheetName} selectedAddr={selectedAddr} onSelect={onSelect} />
    </div>
  );
}
