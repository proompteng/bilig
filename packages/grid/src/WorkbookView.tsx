import React from "react";
import type { SpreadsheetEngine } from "@bilig/core";
import { FormulaBar } from "./FormulaBar.js";
import { SheetGridView, type EditMovement } from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: SpreadsheetEngine;
  workbookName: string;
  variant?: "playground" | "product";
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  editorValue: string;
  resolvedValue: string;
  isEditing: boolean;
  isEditingCell: boolean;
  onSelectSheet(sheetName: string): void;
  onSelect(addr: string): void;
  onAddressCommit(addr: string): void;
  onBeginEdit(seed?: string): void;
  onBeginFormulaEdit(seed?: string): void;
  onEditorChange(next: string): void;
  onCommitEdit(movement?: EditMovement): void;
  onCancelEdit(): void;
  onClearCell(): void;
  onPaste(addr: string, values: readonly (readonly string[])[]): void;
  ribbon?: React.ReactNode;
  sidebar?: React.ReactNode;
  statusBar?: React.ReactNode;
}

export function WorkbookView({
  engine,
  workbookName,
  variant = "playground",
  sheetNames,
  sheetName,
  selectedAddr,
  editorValue,
  resolvedValue,
  isEditing,
  isEditingCell,
  onSelectSheet,
  onSelect,
  onAddressCommit,
  onBeginEdit,
  onBeginFormulaEdit,
  onEditorChange,
  onCommitEdit,
  onCancelEdit,
  onClearCell,
  onPaste,
  ribbon,
  sidebar,
  statusBar
}: WorkbookViewProps) {
  return (
    <section className={variant === "product" ? "workbook-shell workbook-shell-product" : "workbook-shell"}>
      {ribbon ? <div className="workbook-ribbon">{ribbon}</div> : null}
      <div className="workbook-content">
        <div className="workbook-main">
          <div className="workbook-header">
            <div>
              {variant === "playground" ? <p className="panel-eyebrow">Workbook</p> : null}
              <h1>{workbookName}</h1>
            </div>
            {variant === "playground" ? (
              <div className="workbook-header-meta">
                <span className="selection-chip">
                  {sheetName}!{selectedAddr}
                </span>
                <span className="surface-chip">Excel-scale surface</span>
              </div>
            ) : null}
          </div>
          <FormulaBar
            address={selectedAddr}
            isEditing={isEditing}
            onBeginEdit={onBeginFormulaEdit}
            onAddressCommit={onAddressCommit}
            onCancel={onCancelEdit}
            onChange={onEditorChange}
            onClear={onClearCell}
            onCommit={() => onCommitEdit()}
            resolvedValue={resolvedValue}
            sheetName={sheetName}
            value={editorValue}
            variant={variant}
          />
          <SheetGridView
            editorValue={editorValue}
            engine={engine}
            isEditingCell={isEditingCell}
            onBeginEdit={onBeginEdit}
            onCancelEdit={onCancelEdit}
            onClearCell={onClearCell}
            onCommitEdit={onCommitEdit}
            onEditorChange={onEditorChange}
            onPaste={onPaste}
            onSelect={onSelect}
            resolvedValue={resolvedValue}
            selectedAddr={selectedAddr}
            sheetName={sheetName}
            variant={variant}
          />
          <div className="workbook-footer">
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
            {statusBar ? <div className="workbook-status">{statusBar}</div> : null}
          </div>
        </div>
        {sidebar ? <aside className="workbook-sidebar">{sidebar}</aside> : null}
      </div>
    </section>
  );
}
