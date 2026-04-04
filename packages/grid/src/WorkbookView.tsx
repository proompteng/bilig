import React from "react";
import type { Viewport } from "@bilig/protocol";
import { FormulaBar } from "./FormulaBar.js";
import type { GridEngineLike } from "./grid-engine.js";
import { WorkbookSheetTabs } from "./WorkbookSheetTabs.js";
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
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
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
  onToggleBooleanCell?:
    | ((sheetName: string, address: string, nextValue: boolean) => void)
    | undefined;
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
  onCreateSheet,
  onRenameSheet,
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
  onToggleBooleanCell,
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
    <section
      className="flex h-screen flex-col overflow-hidden bg-[var(--wb-surface)] font-sans"
      data-testid="workbook-shell"
    >
      {ribbon ? <div className="shrink-0">{ribbon}</div> : null}
      <div className="flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]">
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
          onToggleBooleanCell={onToggleBooleanCell}
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
        <WorkbookSheetTabs
          onCreateSheet={onCreateSheet}
          onRenameSheet={onRenameSheet}
          onSelectSheet={onSelectSheet}
          sheetName={sheetName}
          sheetNames={sheetNames}
          statusBar={statusBar}
        />
      </div>
    </section>
  );
}
