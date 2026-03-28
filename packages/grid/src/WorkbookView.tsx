import React from "react";
import type { Viewport } from "@bilig/protocol";
import { FormulaBar } from "./FormulaBar.js";
import type { GridEngineLike } from "./grid-engine.js";
import {
  SheetGridView,
  type EditMovement,
  type EditSelectionBehavior,
  type SheetGridViewportSubscription,
} from "./SheetGridView.js";

interface WorkbookViewProps {
  engine: GridEngineLike;
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
  onPaste(this: void, addr: string, values: readonly (readonly string[])[]): void;
  onSelectionLabelChange?: ((label: string) => void) | undefined;
  ribbon?: React.ReactNode;
  statusBar?: React.ReactNode;
  subscribeViewport?: SheetGridViewportSubscription | undefined;
  columnWidths?: Readonly<Record<number, number>> | undefined;
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined;
  onAutofitColumn?:
    | ((columnIndex: number, fallbackWidth: number) => void | Promise<void>)
    | undefined;
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined;
}

export function WorkbookView({
  engine,
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
  statusBar,
  subscribeViewport,
  columnWidths,
  onColumnWidthChange,
  onAutofitColumn,
  onVisibleViewportChange,
}: WorkbookViewProps) {
  return (
    <section className="flex h-screen flex-col overflow-hidden bg-[#f8f9fa]">
      {ribbon ? <div className="shrink-0">{ribbon}</div> : null}
      <div className="flex min-h-0 flex-1 flex-col bg-white">
        <FormulaBar
          address={selectedAddr}
          isEditing={isEditing}
          onBeginEdit={onBeginFormulaEdit}
          onAddressCommit={onAddressCommit}
          onCancel={onCancelEdit}
          onChange={onEditorChange}
          onCommit={() => onCommitEdit()}
          resolvedValue={resolvedValue}
          sheetName={sheetName}
          value={editorValue}
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
        />
        <div className="flex min-h-8 items-center justify-between gap-3 border-t border-[#dadce0] bg-[#f8f9fa] px-2">
          <div aria-label="Sheets" className="flex items-end gap-1" role="tablist">
            {sheetNames.map((name) => (
              <button
                aria-selected={name === sheetName}
                className={`inline-flex h-7 items-center rounded-t-[3px] border border-b-0 px-3 text-[12px] ${
                  name === sheetName
                    ? "border-[#dadce0] bg-white text-[#202124]"
                    : "border-transparent bg-transparent text-[#5f6368] hover:bg-[#eef2f6]"
                }`}
                key={name}
                onClick={() => onSelectSheet(name)}
                role="tab"
                type="button"
              >
                {name}
              </button>
            ))}
          </div>
          {statusBar ? (
            <div className="inline-flex flex-wrap items-center gap-3 text-[11px] text-[#5f6368]">
              {statusBar}
            </div>
          ) : null}
        </div>
      </div>
    </section>
  );
}
