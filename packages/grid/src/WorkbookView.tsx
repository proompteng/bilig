import React from "react";
import type { CellSnapshot, Viewport } from "@bilig/protocol";
import { FormulaBar } from "./FormulaBar.js";
import type { GridEngineLike } from "./grid-engine.js";
import { WorkbookSheetTabs } from "./WorkbookSheetTabs.js";
import {
  WorkbookGridSurface,
  type EditMovement,
  type EditSelectionBehavior,
  type SheetGridViewportSubscription,
} from "./WorkbookGridSurface.js";

interface WorkbookViewProps {
  engine: GridEngineLike;
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  selectedCellSnapshot: CellSnapshot;
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
  onMoveRange(
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
  selectionStatus?: React.ReactNode;
  headerStatus?: React.ReactNode;
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
  selectedCellSnapshot,
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
  onMoveRange,
  onToggleBooleanCell,
  onPaste,
  onSelectionLabelChange,
  ribbon,
  selectionStatus,
  headerStatus,
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
      {ribbon || headerStatus ? (
        <div className="flex shrink-0 items-start justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface)]">
          <div className="min-w-0 flex-1">{ribbon}</div>
          {headerStatus ? (
            <div className="flex min-h-10 shrink-0 items-center px-2.5 py-1">{headerStatus}</div>
          ) : null}
        </div>
      ) : null}
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
        <WorkbookGridSurface
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
          onMoveRange={onMoveRange}
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
          selectedCellSnapshot={selectedCellSnapshot}
          sheetName={sheetName}
        />
        <WorkbookSheetTabs
          onCreateSheet={onCreateSheet}
          onRenameSheet={onRenameSheet}
          onSelectSheet={onSelectSheet}
          sheetName={sheetName}
          sheetNames={sheetNames}
          selectionStatus={selectionStatus}
        />
      </div>
    </section>
  );
}
