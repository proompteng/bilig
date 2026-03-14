import React from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { SheetGridView } from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: SpreadsheetEngine;
  workbookName: string;
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  onSelectSheet(sheetName: string): void;
  onSelect(addr: string): void;
}

export function WorkbookView({
  engine,
  workbookName,
  sheetNames,
  sheetName,
  selectedAddr,
  onSelectSheet,
  onSelect
}: WorkbookViewProps) {
  return (
    <div className="panel workbook-panel">
      <div className="workbook-header">
        <div>
          <p className="panel-eyebrow">Workbook</p>
          <h2>{workbookName}</h2>
        </div>
        <div aria-label="Selected cell" aria-live="polite" className="selection-chip" data-testid="selection-chip">
          {sheetName}!{selectedAddr}
        </div>
      </div>
      <div className="sheet-tabs" aria-label="Sheets" role="tablist">
        {sheetNames.map((name) => (
          <button
            aria-selected={name === sheetName}
            className={name === sheetName ? "sheet-tab active" : "sheet-tab"}
            key={name}
            onClick={() => onSelectSheet(name)}
            role="tab"
            type="button"
          >
            {name}
          </button>
        ))}
      </div>
      <SheetGridView engine={engine} sheetName={sheetName} selectedAddr={selectedAddr} onSelect={onSelect} />
    </div>
  );
}
