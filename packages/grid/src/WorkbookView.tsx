import React from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { SheetGridView } from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: SpreadsheetEngine;
  workbookName: string;
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  editorValue: string;
  resolvedValue: string;
  isEditingCell: boolean;
  onSelectSheet(sheetName: string): void;
  onSelect(addr: string): void;
  onBeginEdit(): void;
  onEditorChange(next: string): void;
  onCommitEdit(): void;
  onCancelEdit(): void;
}

export function WorkbookView({
  engine,
  workbookName,
  sheetNames,
  sheetName,
  selectedAddr,
  editorValue,
  resolvedValue,
  isEditingCell,
  onSelectSheet,
  onSelect,
  onBeginEdit,
  onEditorChange,
  onCommitEdit,
  onCancelEdit
}: WorkbookViewProps) {
  return (
    <div className="panel workbook-panel">
      <div className="workbook-header">
        <div>
          <p className="panel-eyebrow">Workbook</p>
          <h2>{workbookName}</h2>
        </div>
        <div
          aria-label="Selected cell"
          aria-live="polite"
          className="selection-chip"
          data-testid="selection-chip"
        >
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
      <SheetGridView
        editorValue={editorValue}
        engine={engine}
        isEditingCell={isEditingCell}
        onBeginEdit={onBeginEdit}
        onCancelEdit={onCancelEdit}
        onCommitEdit={onCommitEdit}
        onEditorChange={onEditorChange}
        onSelect={onSelect}
        resolvedValue={resolvedValue}
        selectedAddr={selectedAddr}
        sheetName={sheetName}
      />
    </div>
  );
}
