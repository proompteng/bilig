import React, { useCallback, useRef, useState } from "react";
import type { CellSnapshot, Viewport, WorkbookDefinedNameSnapshot } from "@bilig/protocol";
import { FormulaBar } from "./FormulaBar.js";
import type { GridEngineLike } from "./grid-engine.js";
import { WorkbookSheetTabs } from "./WorkbookSheetTabs.js";
import {
  WorkbookGridSurface,
  type EditMovement,
  type EditSelectionBehavior,
  type WorkbookGridPreviewRange,
  type SheetGridViewportSubscription,
} from "./WorkbookGridSurface.js";

interface WorkbookViewProps {
  engine: GridEngineLike;
  sheetNames: string[];
  sheetName: string;
  selectedAddr: string;
  selectedCellSnapshot: CellSnapshot;
  definedNames?: readonly WorkbookDefinedNameSnapshot[];
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
  sideRail?: React.ReactNode;
  sideRailId?: string | undefined;
  sideRailWidth?: number | undefined;
  onSideRailWidthChange?: ((width: number) => void) | undefined;
  subscribeViewport?: SheetGridViewportSubscription | undefined;
  columnWidths?: Readonly<Record<number, number>> | undefined;
  hiddenColumns?: Readonly<Record<number, true>> | undefined;
  hiddenRows?: Readonly<Record<number, true>> | undefined;
  rowHeights?: Readonly<Record<number, number>> | undefined;
  freezeRows?: number | undefined;
  freezeCols?: number | undefined;
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined;
  onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined;
  onSetColumnHidden?: ((columnIndex: number, hidden: boolean) => void) | undefined;
  onSetRowHidden?: ((rowIndex: number, hidden: boolean) => void) | undefined;
  onInsertRows?: ((startRow: number, count: number) => void) | undefined;
  onDeleteRows?: ((startRow: number, count: number) => void) | undefined;
  onInsertColumns?: ((startCol: number, count: number) => void) | undefined;
  onDeleteColumns?: ((startCol: number, count: number) => void) | undefined;
  onSetFreezePane?: ((rows: number, cols: number) => void) | undefined;
  onAutofitColumn?:
    | ((columnIndex: number, fallbackWidth: number) => void | Promise<void>)
    | undefined;
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined;
  previewRanges?: readonly WorkbookGridPreviewRange[] | undefined;
  restoreViewportTarget?:
    | {
        readonly token: number;
        readonly viewport: Viewport;
      }
    | undefined;
}

const MIN_SIDE_RAIL_WIDTH = 304;
const MAX_SIDE_RAIL_WIDTH = 520;

function clampSideRailWidth(width: number): number {
  return Math.min(MAX_SIDE_RAIL_WIDTH, Math.max(MIN_SIDE_RAIL_WIDTH, Math.round(width)));
}

export function WorkbookView({
  engine,
  sheetNames,
  sheetName,
  selectedAddr,
  selectedCellSnapshot,
  definedNames,
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
  sideRail,
  sideRailId,
  sideRailWidth,
  onSideRailWidthChange,
  subscribeViewport,
  columnWidths,
  hiddenColumns,
  hiddenRows,
  rowHeights,
  freezeRows,
  freezeCols,
  onColumnWidthChange,
  onRowHeightChange,
  onSetColumnHidden,
  onSetRowHidden,
  onInsertRows,
  onDeleteRows,
  onInsertColumns,
  onDeleteColumns,
  onSetFreezePane,
  onAutofitColumn,
  onVisibleViewportChange,
  previewRanges,
  restoreViewportTarget,
}: WorkbookViewProps) {
  const resizeStateRef = useRef<{
    pointerId: number;
    startX: number;
    startWidth: number;
  } | null>(null);
  const [isResizingSideRail, setIsResizingSideRail] = useState(false);
  const resolvedSideRailWidth = clampSideRailWidth(sideRailWidth ?? 344);

  const handleSideRailPointerMove = useCallback(
    (event: React.PointerEvent<HTMLDivElement>) => {
      const resizeState = resizeStateRef.current;
      if (!resizeState || event.pointerId !== resizeState.pointerId || !onSideRailWidthChange) {
        return;
      }
      const nextWidth = resizeState.startWidth + (resizeState.startX - event.clientX);
      onSideRailWidthChange(clampSideRailWidth(nextWidth));
    },
    [onSideRailWidthChange],
  );

  const finishSideRailResize = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const resizeState = resizeStateRef.current;
    if (!resizeState || event.pointerId !== resizeState.pointerId) {
      return;
    }
    resizeStateRef.current = null;
    setIsResizingSideRail(false);
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }
  }, []);

  return (
    <section
      className="flex h-screen flex-col overflow-hidden bg-[var(--wb-surface)] font-sans"
      data-testid="workbook-shell"
    >
      {ribbon || headerStatus ? (
        <div className="flex shrink-0 items-start justify-between gap-3 bg-[var(--wb-surface)]">
          <div className="min-w-0 flex-1">{ribbon}</div>
          {headerStatus ? (
            <div className="flex min-h-10 shrink-0 items-center justify-end px-2.5 py-1">
              {headerStatus}
            </div>
          ) : null}
        </div>
      ) : null}
      <div className="flex min-h-0 flex-1 bg-[var(--wb-surface)]">
        <div className="flex min-h-0 min-w-0 flex-1 flex-col">
          <FormulaBar
            address={selectedAddr}
            {...(definedNames ? { definedNames } : {})}
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
            hiddenColumns={hiddenColumns}
            hiddenRows={hiddenRows}
            rowHeights={rowHeights}
            freezeRows={freezeRows}
            freezeCols={freezeCols}
            onColumnWidthChange={onColumnWidthChange}
            onRowHeightChange={onRowHeightChange}
            onSetColumnHidden={onSetColumnHidden}
            onSetRowHidden={onSetRowHidden}
            onInsertRows={onInsertRows}
            onDeleteRows={onDeleteRows}
            onInsertColumns={onInsertColumns}
            onDeleteColumns={onDeleteColumns}
            onSetFreezePane={onSetFreezePane}
            onAutofitColumn={onAutofitColumn}
            onVisibleViewportChange={onVisibleViewportChange}
            previewRanges={previewRanges}
            restoreViewportTarget={restoreViewportTarget}
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
        {sideRail ? (
          <aside
            className="relative flex h-full shrink-0 bg-[var(--wb-app-bg)]"
            data-testid="workbook-side-rail"
            id={sideRailId}
            style={{
              flexBasis: `${resolvedSideRailWidth}px`,
              width: `${resolvedSideRailWidth}px`,
            }}
          >
            {onSideRailWidthChange ? (
              <div
                aria-label="Resize workbook side rail"
                aria-orientation="vertical"
                className={[
                  "absolute inset-y-0 left-0 z-10 w-4 -translate-x-2 cursor-ew-resize touch-none",
                  "after:absolute after:inset-y-0 after:left-1/2 after:w-px after:-translate-x-1/2 after:bg-[var(--color-mauve-200)] after:transition-colors",
                  isResizingSideRail
                    ? "after:bg-[var(--color-mauve-500)]"
                    : "hover:after:bg-[var(--color-mauve-300)]",
                ].join(" ")}
                data-testid="workbook-side-rail-resize-handle"
                role="separator"
                onPointerCancel={finishSideRailResize}
                onPointerDown={(event) => {
                  resizeStateRef.current = {
                    pointerId: event.pointerId,
                    startX: event.clientX,
                    startWidth: resolvedSideRailWidth,
                  };
                  setIsResizingSideRail(true);
                  event.currentTarget.setPointerCapture(event.pointerId);
                }}
                onPointerMove={handleSideRailPointerMove}
                onPointerUp={finishSideRailResize}
              />
            ) : null}
            {sideRail}
          </aside>
        ) : null}
      </div>
    </section>
  );
}
