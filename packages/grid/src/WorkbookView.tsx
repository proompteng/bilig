import React from "react";
import { Tabs } from "@base-ui/react/tabs";
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
  onPaste(
    this: void,
    sheetName: string,
    addr: string,
    values: readonly (readonly string[])[],
  ): void;
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
    <section className="flex h-screen flex-col overflow-hidden bg-[#f6f8fb]">
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
        <div className="flex min-h-9 items-center justify-between gap-3 border-t border-[#d7dce5] bg-[#f8fafc] px-2.5">
          <Tabs.Root value={sheetName} onValueChange={(value) => onSelectSheet(String(value))}>
            <Tabs.List aria-label="Sheets" className="flex items-end gap-1">
              {sheetNames.map((name) => (
                <Tabs.Tab
                  className="inline-flex h-8 items-center rounded-t-[6px] border border-b-0 border-transparent bg-transparent px-3 text-[12px] font-medium text-[#5f6368] outline-none transition-[background-color,border-color,color] hover:bg-[#eef3f9] focus-visible:border-[#1a73e8] focus-visible:bg-white focus-visible:ring-2 focus-visible:ring-[#d2e3fc] data-[active]:border-[#d6dbe6] data-[active]:bg-white data-[active]:text-[#202124]"
                  key={name}
                  value={name}
                >
                  {name}
                </Tabs.Tab>
              ))}
            </Tabs.List>
          </Tabs.Root>
          {statusBar ? (
            <div className="inline-flex flex-wrap items-center gap-2 text-[11px] text-[#5f6368]">
              {statusBar}
            </div>
          ) : null}
        </div>
      </div>
    </section>
  );
}
