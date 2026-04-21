import { useCallback, useMemo } from 'react'
import { parseCellAddress } from '@bilig/formula'
import { CellEditorOverlay } from './CellEditorOverlay.js'
import { GridFillHandleOverlay } from './GridFillHandleOverlay.js'
import { WorkbookGridContextMenu } from './WorkbookGridContextMenu.js'
import { createGridGeometrySnapshot } from './gridGeometry.js'
import { WorkbookPaneRendererV2, buildDynamicGridOverlayPacket } from './renderer-v2/index.js'
import { resolveResizeGuideColumn, resolveResizeGuideRow } from './useGridResizeState.js'
import { useWorkbookGridInteractions } from './useWorkbookGridInteractions.js'
import { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'
import type { WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
export { hasSelectionTargetChanged } from './workbookGridViewport.js'
export type {
  EditTargetSelection,
  EditMovement,
  EditSelectionBehavior,
  GridSelectionSnapshot,
  SheetGridViewportSubscription,
  WorkbookGridPreviewRange,
  WorkbookGridSurfaceProps,
} from './workbookGridSurfaceTypes.js'

export function WorkbookGridSurface(props: WorkbookGridSurfaceProps) {
  const renderState = useWorkbookGridRenderState({
    engine: props.engine,
    sheetName: props.sheetName,
    selectedAddr: props.selectedAddr,
    selectedCellSnapshot: props.selectedCellSnapshot,
    editorValue: props.editorValue,
    isEditingCell: props.isEditingCell,
    subscribeViewport: props.subscribeViewport,
    controlledColumnWidths: props.columnWidths,
    controlledRowHeights: props.rowHeights,
    controlledHiddenColumns: props.hiddenColumns,
    controlledHiddenRows: props.hiddenRows,
    getCellEditorSeed: props.getCellEditorSeed,
    freezeRows: props.freezeRows,
    freezeCols: props.freezeCols,
    onVisibleViewportChange: props.onVisibleViewportChange,
    onColumnWidthChange: props.onColumnWidthChange,
    onRowHeightChange: props.onRowHeightChange,
    restoreViewportTarget: props.restoreViewportTarget,
  })
  const interactions = useWorkbookGridInteractions({
    engine: props.engine,
    sheetName: props.sheetName,
    selectedAddr: props.selectedAddr,
    editorValue: props.editorValue,
    isEditingCell: props.isEditingCell,
    onAutofitColumn: props.onAutofitColumn,
    onBeginEdit: props.onBeginEdit,
    onCancelEdit: props.onCancelEdit,
    onClearCell: props.onClearCell,
    onColumnWidthChange: props.onColumnWidthChange,
    onCommitEdit: props.onCommitEdit,
    onCopyRange: props.onCopyRange,
    onEditorChange: props.onEditorChange,
    onFillRange: props.onFillRange,
    onMoveRange: props.onMoveRange,
    onPaste: props.onPaste,
    hiddenColumns: props.hiddenColumns,
    hiddenRows: props.hiddenRows,
    getCellEditorSeed: props.getCellEditorSeed,
    onSetColumnHidden: props.onSetColumnHidden,
    onSetRowHidden: props.onSetRowHidden,
    onInsertRows: props.onInsertRows,
    onDeleteRows: props.onDeleteRows,
    onInsertColumns: props.onInsertColumns,
    onDeleteColumns: props.onDeleteColumns,
    onSetFreezePane: props.onSetFreezePane,
    onSelectionChange: props.onSelectionChange,
    onSelectionLabelChange: props.onSelectionLabelChange,
    selectionSnapshot: props.selectionSnapshot,
    onToggleBooleanCell: props.onToggleBooleanCell,
    onRowHeightChange: props.onRowHeightChange,
    selectedCellSnapshot: props.selectedCellSnapshot,
    renderState,
  })
  const visibleRange = renderState.visibleRegion.range
  const getCellLocalBounds = renderState.getCellLocalBounds
  const v2Geometry = useMemo(
    () =>
      renderState.hostElement
        ? createGridGeometrySnapshot({
            columnWidths: renderState.columnWidths,
            dpr: typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1),
            freezeCols: props.freezeCols,
            freezeRows: props.freezeRows,
            gridMetrics: renderState.gridMetrics,
            hiddenColumns: props.hiddenColumns,
            hiddenRows: props.hiddenRows,
            hostHeight: renderState.hostElement.clientHeight,
            hostWidth: renderState.hostElement.clientWidth,
            rowHeights: renderState.rowHeights,
            scrollLeft: renderState.scrollViewportRef.current?.scrollLeft ?? 0,
            scrollTop: renderState.scrollViewportRef.current?.scrollTop ?? 0,
            sheetName: props.sheetName,
          })
        : null,
    [
      props.freezeCols,
      props.freezeRows,
      props.hiddenColumns,
      props.hiddenRows,
      props.sheetName,
      renderState.columnWidths,
      renderState.gridMetrics,
      renderState.hostElement,
      renderState.rowHeights,
      renderState.scrollViewportRef,
    ],
  )
  const dynamicOverlayBuilder = useCallback(
    (geometry: NonNullable<typeof v2Geometry>) =>
      buildDynamicGridOverlayPacket({
        geometry,
        selectionRange: null,
        showFillHandle: false,
        resizeGuideColumn: resolveResizeGuideColumn({
          activeResizeColumn: renderState.activeResizeColumn,
          cursor: renderState.hoverState.cursor,
          header: renderState.hoverState.header,
        }),
        resizeGuideRow: resolveResizeGuideRow({
          activeResizeRow: renderState.activeResizeRow,
          cursor: renderState.hoverState.cursor,
          header: renderState.hoverState.header,
        }),
      }),
    [renderState.activeResizeColumn, renderState.activeResizeRow, renderState.hoverState.cursor, renderState.hoverState.header],
  )
  const previewRects = useMemo(() => {
    return (props.previewRanges ?? [])
      .filter((range) => range.sheetName === props.sheetName)
      .flatMap((range) => {
        const start = parseCellAddress(range.startAddress, range.sheetName)
        const end = parseCellAddress(range.endAddress, range.sheetName)
        const colStart = Math.max(Math.min(start.col, end.col), visibleRange.x)
        const colEnd = Math.min(Math.max(start.col, end.col), visibleRange.x + visibleRange.width - 1)
        const rowStart = Math.max(Math.min(start.row, end.row), visibleRange.y)
        const rowEnd = Math.min(Math.max(start.row, end.row), visibleRange.y + visibleRange.height - 1)
        if (colStart > colEnd || rowStart > rowEnd) {
          return []
        }
        const topLeft = getCellLocalBounds(colStart, rowStart)
        const bottomRight = getCellLocalBounds(colEnd, rowEnd)
        if (!topLeft || !bottomRight) {
          return []
        }
        return [
          {
            key: `${range.role}:${range.sheetName}:${range.startAddress}:${range.endAddress}`,
            role: range.role,
            bounds: {
              x: topLeft.x,
              y: topLeft.y,
              width: bottomRight.x + bottomRight.width - topLeft.x,
              height: bottomRight.y + bottomRight.height - topLeft.y,
            },
          },
        ]
      })
  }, [props.previewRanges, props.sheetName, getCellLocalBounds, visibleRange.height, visibleRange.width, visibleRange.x, visibleRange.y])

  return (
    <div className="relative flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]">
      <div
        className="sheet-grid-host min-h-0 flex-1 overflow-hidden bg-[var(--wb-surface)] pr-2 pb-2"
        data-column-width-overrides={renderState.columnWidthOverridesAttr}
        data-default-column-width={renderState.gridMetrics.columnWidth}
        data-testid="sheet-grid"
        role="grid"
        style={{ cursor: renderState.hoverState.cursor }}
        onFocus={interactions.handleHostFocus}
        onKeyDownCapture={interactions.handleHostKeyDownCapture}
        onClickCapture={interactions.handleHostClickCapture}
        onCopyCapture={interactions.handleHostCopyCapture}
        onContextMenuCapture={interactions.handleHostContextMenuCapture}
        onPasteCapture={interactions.handleHostPasteCapture}
        onKeyDown={interactions.handleHostKeyDown}
        onDoubleClickCapture={interactions.handleHostDoubleClickCapture}
        onPointerMoveCapture={interactions.handleHostPointerMoveCapture}
        onPointerLeave={interactions.handleHostPointerLeave}
        onPointerDownCapture={interactions.handleHostPointerDownCapture}
        onPointerUpCapture={interactions.handleHostPointerUpCapture}
        ref={renderState.handleHostRef}
        // oxlint-disable-next-line jsx-a11y/no-noninteractive-tabindex
        tabIndex={0}
      >
        <div
          aria-label={`${props.sheetName} grid focus target`}
          className="pointer-events-none absolute h-px w-px overflow-hidden opacity-0"
          data-testid="sheet-grid-focus-target"
          ref={renderState.focusTargetRef}
          tabIndex={-1}
        />
        <div
          ref={renderState.scrollViewportRef}
          aria-hidden="true"
          className="absolute inset-0 overflow-auto"
          data-testid="grid-scroll-viewport"
        >
          <div style={{ height: renderState.totalGridHeight, width: renderState.totalGridWidth }} />
        </div>
        <WorkbookPaneRendererV2
          active={renderState.hostElement !== null}
          cameraStore={renderState.gridCameraStore}
          geometry={v2Geometry}
          host={renderState.hostElement}
          overlayBuilder={dynamicOverlayBuilder}
          panes={renderState.renderPanes}
          scrollTransformStore={renderState.scrollTransformStore}
        />
        <button
          aria-label="Select entire sheet"
          className="absolute z-20 flex items-center justify-center border-r border-b border-[var(--wb-border-subtle)] bg-[var(--wb-muted)] text-[var(--wb-text-muted)] outline-none transition-colors hover:bg-[var(--wb-muted-strong)] hover:text-[var(--wb-text)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent)] focus-visible:ring-offset-0"
          data-testid="grid-select-entire-sheet"
          onClick={interactions.handleSelectEntireSheet}
          style={{
            height: renderState.gridMetrics.headerHeight,
            left: 0,
            top: 0,
            width: renderState.gridMetrics.rowMarkerWidth,
          }}
          type="button"
        >
          <span
            aria-hidden="true"
            className="block h-0 w-0 border-t-[11px] border-r-[11px] border-t-transparent border-r-current opacity-80"
            style={{
              color: renderState.isEntireSheetSelected ? 'var(--wb-accent)' : 'currentColor',
              transform: 'translate(2px, 1px)',
            }}
          />
        </button>
        <GridFillHandleOverlay
          getCellBounds={renderState.getCellLocalBounds}
          hidden={
            renderState.hostElement === null ||
            !renderState.selectionRange ||
            renderState.gridSelection.columns.length > 0 ||
            renderState.gridSelection.rows.length > 0 ||
            Boolean(renderState.fillPreviewRange) ||
            renderState.isRangeMoveDragging
          }
          hostHeight={renderState.hostElement?.clientHeight ?? 0}
          hostWidth={renderState.hostElement?.clientWidth ?? 0}
          minX={renderState.gridMetrics.rowMarkerWidth}
          minY={renderState.gridMetrics.headerHeight}
          onPointerDown={interactions.handleFillHandlePointerDown}
          scrollTransformStore={renderState.scrollTransformStore}
          selectionRange={renderState.selectionRange}
        />
        {renderState.fillPreviewBounds ? (
          <div
            aria-hidden="true"
            className="pointer-events-none absolute z-20 box-border border border-dashed"
            data-grid-fill-preview="true"
            style={{
              borderColor: renderState.gridTheme.textMedium,
              height: renderState.fillPreviewBounds.height,
              left: renderState.fillPreviewBounds.x,
              top: renderState.fillPreviewBounds.y,
              width: renderState.fillPreviewBounds.width,
            }}
          />
        ) : null}
        {previewRects.map((previewRect) => (
          <div
            key={previewRect.key}
            aria-hidden="true"
            className="pointer-events-none absolute z-20 box-border border border-dashed"
            style={{
              backgroundColor: previewRect.role === 'target' ? 'rgba(56, 189, 248, 0.08)' : 'rgba(148, 163, 184, 0.06)',
              borderColor: previewRect.role === 'target' ? 'rgba(14, 116, 144, 0.9)' : 'rgba(100, 116, 139, 0.9)',
              height: previewRect.bounds.height,
              left: previewRect.bounds.x,
              top: previewRect.bounds.y,
              width: previewRect.bounds.width,
            }}
          />
        ))}
        <div className="pointer-events-none absolute inset-0 z-[1]" />
      </div>
      {interactions.contextMenu.contextMenuState ? (
        <WorkbookGridContextMenu
          canUnfreezePanes={interactions.contextMenu.canUnfreezePanes}
          menuRef={interactions.contextMenu.menuRef}
          onClose={interactions.contextMenu.closeContextMenu}
          onDeleteTarget={interactions.contextMenu.deleteTarget}
          onFreezeTarget={interactions.contextMenu.freezeTarget}
          onInsertAfterTarget={interactions.contextMenu.insertAfterTarget}
          onInsertBeforeTarget={interactions.contextMenu.insertBeforeTarget}
          onToggleTargetHidden={interactions.contextMenu.toggleTargetHidden}
          onUnfreezePanes={interactions.contextMenu.unfreezePanes}
          state={interactions.contextMenu.contextMenuState}
        />
      ) : null}
      {props.isEditingCell && renderState.overlayStyle ? (
        <CellEditorOverlay
          label={`${props.sheetName}!${props.selectedAddr}`}
          targetSelection={{ sheetName: props.sheetName, address: props.selectedAddr }}
          onCancel={props.onCancelEdit}
          onChange={props.onEditorChange}
          onCommit={props.onCommitEdit}
          backgroundColor={renderState.editorPresentation.backgroundColor}
          color={renderState.editorPresentation.color}
          font={renderState.editorPresentation.font}
          fontSize={renderState.editorPresentation.fontSize}
          resolvedValue={props.resolvedValue}
          selectionBehavior={props.editorSelectionBehavior}
          textAlign={renderState.editorTextAlign}
          underline={renderState.editorPresentation.underline}
          value={props.editorValue}
          style={renderState.overlayStyle}
        />
      ) : null}
    </div>
  )
}
