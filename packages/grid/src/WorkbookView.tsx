import React from "react";
import type { Viewport } from "@bilig/protocol";
import { FormulaBar } from "./FormulaBar.js";
import type { GridEngineLike } from "./grid-engine.js";
import {
  SheetGridView,
  type EditMovement,
  type EditSelectionBehavior,
  type SheetGridViewportSubscription
} from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: GridEngineLike;
  workbookName: string;
  variant?: "playground" | "product";
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  editorValue: string;
  editorSelectionBehavior: EditSelectionBehavior;
  resolvedValue: string;
  isEditing: boolean;
  isEditingCell: boolean;
  onSelectSheet(this: void, sheetName: string): void;
  onSelect(this: void, addr: string): void;
  onAddressCommit(this: void, addr: string): void;
  onBeginEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void;
  onBeginFormulaEdit(this: void, seed?: string): void;
  onEditorChange(this: void, next: string): void;
  onCommitEdit(this: void, movement?: EditMovement): void;
  onCancelEdit(this: void): void;
  onClearCell(this: void): void;
  onFillRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onCopyRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void;
  onPaste(this: void, addr: string, values: readonly (readonly string[])[]): void;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
  ribbon?: React.ReactNode;
  sidebar?: React.ReactNode;
  statusBar?: React.ReactNode;
  subscribeViewport?: SheetGridViewportSubscription | undefined;
  columnWidths?: Readonly<Record<number, number>> | undefined;
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined;
  onAutofitColumn?: ((columnIndex: number, fallbackWidth: number) => void | Promise<void>) | undefined;
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined;
}

export function WorkbookView({
  engine,
  workbookName,
  variant = "playground",
  sheetNames,
  sheetName,
  selectedAddr,
  editorValue,
  editorSelectionBehavior,
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
  onFillRange,
  onCopyRange,
  onPaste,
  onSelectionLabelChange,
  ribbon,
  sidebar,
  statusBar,
  subscribeViewport,
  columnWidths,
  onColumnWidthChange,
  onAutofitColumn,
  onVisibleViewportChange
}: WorkbookViewProps) {
  const showWorkbookHeader = variant !== "product";
  return (
    <section className={variant === "product" ? "workbook-shell workbook-shell-product" : "workbook-shell"}>
      {ribbon ? <div className="workbook-ribbon">{ribbon}</div> : null}
      <div className="workbook-content">
        <div className="workbook-main">
          {showWorkbookHeader ? (
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
          ) : null}
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
            editorSelectionBehavior={editorSelectionBehavior}
            engine={engine}
            isEditingCell={isEditingCell}
            onBeginEdit={onBeginEdit}
            onCancelEdit={onCancelEdit}
            onClearCell={onClearCell}
            onCommitEdit={onCommitEdit}
            onCopyRange={onCopyRange}
            onEditorChange={onEditorChange}
            onFillRange={onFillRange}
            onPaste={onPaste}
            onSelectionLabelChange={onSelectionLabelChange}
            onSelect={onSelect}
            subscribeViewport={subscribeViewport}
            columnWidths={columnWidths}
            onColumnWidthChange={onColumnWidthChange}
            onAutofitColumn={onAutofitColumn}
            onVisibleViewportChange={onVisibleViewportChange}
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
