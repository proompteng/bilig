import React, { Profiler, useCallback, useEffect, useRef, useState } from 'react'
import type { CellSnapshot, Viewport, WorkbookDefinedNameSnapshot } from '@bilig/protocol'
import { FormulaBar } from './FormulaBar.js'
import type { GridEngineLike } from './grid-engine.js'
import { formatSelectionSnapshotSummary } from './gridSelection.js'
import { WorkbookSelectionStatus } from './WorkbookSelectionStatus.js'
import { WorkbookSheetTabs } from './WorkbookSheetTabs.js'
import {
  WorkbookGridSurface,
  type EditMovement,
  type EditSelectionBehavior,
  type GridSelectionSnapshot,
  type WorkbookGridPreviewRange,
  type SheetGridViewportSubscription,
} from './WorkbookGridSurface.js'
import type { WorkbookRendererMode } from './workbookRendererMode.js'

interface WorkbookViewProps {
  engine: GridEngineLike
  sheetNames: string[]
  sheetName: string
  selectedAddr: string
  selectedCellSnapshot: CellSnapshot
  selectionSnapshot: GridSelectionSnapshot
  definedNames?: readonly WorkbookDefinedNameSnapshot[]
  editorValue: string
  editorSelectionBehavior: EditSelectionBehavior
  resolvedValue: string
  isEditing: boolean
  isEditingCell: boolean
  onSelectSheet(this: void, sheetName: string): void
  onCreateSheet?: (() => void) | undefined
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined
  onDeleteSheet?: ((sheetName: string) => void) | undefined
  onSelectionChange(this: void, selection: GridSelectionSnapshot): void
  onAddressCommit(this: void, addr: string): boolean
  getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  onBeginEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void
  onBeginFormulaEdit(this: void, seed?: string): void
  onEditorChange(this: void, next: string): void
  onCommitEdit(this: void, movement?: EditMovement, valueOverride?: string): void
  onCancelEdit(this: void): void
  onClearCell(this: void): void
  onFillRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onCopyRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onMoveRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onToggleBooleanCell?: ((sheetName: string, address: string, nextValue: boolean) => void) | undefined
  onPaste(this: void, sheetName: string, addr: string, values: readonly (readonly string[])[]): void
  onSelectionLabelChange?: ((label: string) => void) | undefined
  ribbon?: React.ReactNode
  sidePanel?: React.ReactNode
  sidePanelId?: string | undefined
  sidePanelWidth?: number | undefined
  onSidePanelWidthChange?: ((width: number) => void) | undefined
  subscribeViewport?: SheetGridViewportSubscription | undefined
  columnWidths?: Readonly<Record<number, number>> | undefined
  hiddenColumns?: Readonly<Record<number, true>> | undefined
  hiddenRows?: Readonly<Record<number, true>> | undefined
  rowHeights?: Readonly<Record<number, number>> | undefined
  freezeRows?: number | undefined
  freezeCols?: number | undefined
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined
  onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined
  onSetColumnHidden?: ((columnIndex: number, hidden: boolean) => void) | undefined
  onSetRowHidden?: ((rowIndex: number, hidden: boolean) => void) | undefined
  onInsertRows?: ((startRow: number, count: number) => void) | undefined
  onDeleteRows?: ((startRow: number, count: number) => void) | undefined
  onInsertColumns?: ((startCol: number, count: number) => void) | undefined
  onDeleteColumns?: ((startCol: number, count: number) => void) | undefined
  onSetFreezePane?: ((rows: number, cols: number) => void) | undefined
  onAutofitColumn?: ((columnIndex: number, fallbackWidth: number) => void | Promise<void>) | undefined
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  previewRanges?: readonly WorkbookGridPreviewRange[] | undefined
  restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
  rendererMode?: WorkbookRendererMode | undefined
}

const MIN_SIDE_PANEL_WIDTH = 280
const MAX_SIDE_PANEL_WIDTH = 420
const SIDE_PANEL_VIEWPORT_FRACTION = 0.42

function noteSurfaceCommit(surface: string): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteSurfaceCommit?: (surface: string) => void } }).__biligScrollPerf?.noteSurfaceCommit?.(
    surface,
  )
}

function clampSidePanelWidth(width: number): number {
  const viewportWidth = typeof window === 'undefined' ? null : window.innerWidth
  const viewportAwareMax =
    viewportWidth && Number.isFinite(viewportWidth)
      ? Math.min(MAX_SIDE_PANEL_WIDTH, Math.max(MIN_SIDE_PANEL_WIDTH, Math.round(viewportWidth * SIDE_PANEL_VIEWPORT_FRACTION)))
      : MAX_SIDE_PANEL_WIDTH
  return Math.min(viewportAwareMax, Math.max(MIN_SIDE_PANEL_WIDTH, Math.round(width)))
}

function sameFormulaBarProps(left: React.ComponentProps<typeof FormulaBar>, right: React.ComponentProps<typeof FormulaBar>): boolean {
  return (
    left.address === right.address &&
    left.definedNames === right.definedNames &&
    left.isEditing === right.isEditing &&
    left.resolvedValue === right.resolvedValue &&
    left.selectionLabel === right.selectionLabel &&
    left.sheetName === right.sheetName &&
    left.value === right.value
  )
}

const MemoFormulaBar = React.memo(FormulaBar, sameFormulaBarProps)
const MemoWorkbookSelectionStatus = React.memo(WorkbookSelectionStatus)
function sameWorkbookSheetTabsProps(
  left: React.ComponentProps<typeof WorkbookSheetTabs>,
  right: React.ComponentProps<typeof WorkbookSheetTabs>,
): boolean {
  return left.sheetName === right.sheetName && left.sheetNames === right.sheetNames && left.trailingContent === right.trailingContent
}

const MemoWorkbookSheetTabs = React.memo(WorkbookSheetTabs, sameWorkbookSheetTabsProps)

const MemoFormulaBarSurface = React.memo(function MemoFormulaBarSurface(props: React.ComponentProps<typeof FormulaBar>) {
  return (
    <Profiler id="workbook-formula-bar" onRender={() => noteSurfaceCommit('formulaBar')}>
      <MemoFormulaBar {...props} />
    </Profiler>
  )
}, sameFormulaBarProps)

const MemoWorkbookSelectionStatusSurface = React.memo(function MemoWorkbookSelectionStatusSurface(
  props: React.ComponentProps<typeof WorkbookSelectionStatus>,
) {
  return (
    <Profiler id="workbook-status" onRender={() => noteSurfaceCommit('statusBar')}>
      <MemoWorkbookSelectionStatus {...props} />
    </Profiler>
  )
})

const MemoWorkbookSheetTabsSurface = React.memo(function MemoWorkbookSheetTabsSurface(
  props: React.ComponentProps<typeof WorkbookSheetTabs>,
) {
  return (
    <Profiler id="workbook-sheet-tabs" onRender={() => noteSurfaceCommit('sheetTabs')}>
      <MemoWorkbookSheetTabs {...props} />
    </Profiler>
  )
}, sameWorkbookSheetTabsProps)

export function WorkbookView({
  engine,
  sheetNames,
  sheetName,
  selectedAddr,
  selectedCellSnapshot,
  selectionSnapshot,
  definedNames,
  editorValue,
  editorSelectionBehavior,
  resolvedValue,
  isEditing,
  isEditingCell,
  onSelectSheet,
  onCreateSheet,
  onRenameSheet,
  onDeleteSheet,
  onSelectionChange,
  onAddressCommit,
  getCellEditorSeed,
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
  sidePanel,
  sidePanelId,
  sidePanelWidth,
  onSidePanelWidthChange,
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
  rendererMode,
}: WorkbookViewProps) {
  const resizeStateRef = useRef<{
    pointerId: number
    startX: number
    startWidth: number
  } | null>(null)
  const [isResizingSidePanel, setIsResizingSidePanel] = useState(false)
  const [selectionLabel, setSelectionLabel] = useState(formatSelectionSnapshotSummary(selectionSnapshot))
  const resolvedSidePanelWidth = clampSidePanelWidth(sidePanelWidth ?? 344)

  useEffect(() => {
    const nextSelectionLabel = formatSelectionSnapshotSummary(selectionSnapshot)
    setSelectionLabel(nextSelectionLabel)
  }, [selectionSnapshot, sheetName])

  const handleSelectionLabelChange = useCallback(
    (label: string) => {
      setSelectionLabel(label)
      onSelectionLabelChange?.(label)
    },
    [onSelectionLabelChange],
  )

  const handleSidePanelPointerMove = useCallback(
    (event: React.PointerEvent<HTMLDivElement>) => {
      const resizeState = resizeStateRef.current
      if (!resizeState || event.pointerId !== resizeState.pointerId || !onSidePanelWidthChange) {
        return
      }
      const nextWidth = resizeState.startWidth + (resizeState.startX - event.clientX)
      onSidePanelWidthChange(clampSidePanelWidth(nextWidth))
    },
    [onSidePanelWidthChange],
  )

  const finishSidePanelResize = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const resizeState = resizeStateRef.current
    if (!resizeState || event.pointerId !== resizeState.pointerId) {
      return
    }
    resizeStateRef.current = null
    setIsResizingSidePanel(false)
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }, [])

  const sheetTabsTrailingContent = React.useMemo(
    () => <MemoWorkbookSelectionStatusSurface engine={engine} selectionLabel={selectionLabel} selectionSnapshot={selectionSnapshot} />,
    [engine, selectionLabel, selectionSnapshot],
  )

  return (
    <section className="flex h-full min-h-0 flex-col overflow-hidden bg-[var(--wb-surface)] font-sans" data-testid="workbook-shell">
      {ribbon ? (
        <Profiler id="workbook-ribbon" onRender={() => noteSurfaceCommit('ribbon')}>
          <div className="shrink-0 bg-[var(--wb-surface)]">{ribbon}</div>
        </Profiler>
      ) : null}
      <div className="flex min-h-0 flex-1 bg-[var(--wb-surface)]">
        <div className="flex min-h-0 min-w-0 flex-1 flex-col">
          <MemoFormulaBarSurface
            address={selectedAddr}
            {...(definedNames ? { definedNames } : {})}
            isEditing={isEditing}
            onBeginEdit={onBeginFormulaEdit}
            onAddressCommit={onAddressCommit}
            onCancel={onCancelEdit}
            onChange={onEditorChange}
            onCommit={() => onCommitEdit()}
            resolvedValue={resolvedValue}
            selectionLabel={selectionLabel}
            sheetName={sheetName}
            value={editorValue}
          />
          <Profiler id="workbook-grid" onRender={() => noteSurfaceCommit('grid')}>
            <WorkbookGridSurface
              editorValue={editorValue}
              editorSelectionBehavior={editorSelectionBehavior}
              engine={engine}
              isEditingCell={isEditingCell}
              getCellEditorSeed={getCellEditorSeed}
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
              onSelectionLabelChange={handleSelectionLabelChange}
              onSelectionChange={onSelectionChange}
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
              rendererMode={rendererMode}
              resolvedValue={resolvedValue}
              selectedAddr={selectedAddr}
              selectedCellSnapshot={selectedCellSnapshot}
              selectionSnapshot={selectionSnapshot}
              sheetName={sheetName}
            />
          </Profiler>
          <MemoWorkbookSheetTabsSurface
            onCreateSheet={onCreateSheet}
            onDeleteSheet={onDeleteSheet}
            onRenameSheet={onRenameSheet}
            onSelectSheet={onSelectSheet}
            sheetName={sheetName}
            sheetNames={sheetNames}
            trailingContent={sheetTabsTrailingContent}
          />
        </div>
        {sidePanel ? (
          <aside
            className="relative flex h-full shrink-0 bg-[var(--wb-app-bg)]"
            data-testid="workbook-side-panel"
            id={sidePanelId}
            style={{
              flexBasis: `${resolvedSidePanelWidth}px`,
              width: `${resolvedSidePanelWidth}px`,
            }}
          >
            {onSidePanelWidthChange ? (
              <div
                aria-label="Resize workbook side panel"
                aria-orientation="vertical"
                className={[
                  'absolute inset-y-0 left-0 z-10 w-4 -translate-x-2 cursor-ew-resize touch-none',
                  'after:absolute after:inset-y-0 after:left-1/2 after:w-px after:-translate-x-1/2 after:bg-[var(--wb-border)] after:transition-colors',
                  isResizingSidePanel ? 'after:bg-[var(--wb-accent)]' : 'hover:after:bg-[var(--wb-border-strong)]',
                ].join(' ')}
                data-testid="workbook-side-panel-resize-handle"
                role="separator"
                onPointerCancel={finishSidePanelResize}
                onPointerDown={(event) => {
                  resizeStateRef.current = {
                    pointerId: event.pointerId,
                    startX: event.clientX,
                    startWidth: resolvedSidePanelWidth,
                  }
                  setIsResizingSidePanel(true)
                  event.currentTarget.setPointerCapture(event.pointerId)
                }}
                onPointerMove={handleSidePanelPointerMove}
                onPointerUp={finishSidePanelResize}
              />
            ) : null}
            {sidePanel}
          </aside>
        ) : null}
      </div>
    </section>
  )
}
