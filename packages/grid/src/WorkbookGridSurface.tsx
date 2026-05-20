import { useCallback, useEffect, useLayoutEffect, useMemo, useRef } from 'react'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { CellEditorOverlay } from './CellEditorOverlay.js'
import { GridFillHandleOverlay } from './GridFillHandleOverlay.js'
import { GridSelectionVisualOverlay } from './GridSelectionVisualOverlay.js'
import { WorkbookGridContextMenu } from './WorkbookGridContextMenu.js'
import { createGridSelection } from './gridSelection.js'
import { WorkbookPaneRendererV3 } from './renderer-v3/WorkbookPaneRendererV3.js'
import { buildDynamicGridOverlayBatchV3 } from './renderer-v3/dynamic-overlay-batch.js'
import { resolveResizeGuideColumn, resolveResizeGuideRow } from './useGridResizeState.js'
import { useWorkbookGridInteractions } from './useWorkbookGridInteractions.js'
import { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from './workbookTheme.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import type { WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
export { hasSelectionTargetChanged } from './workbookGridViewport.js'
export type {
  EditTargetSelection,
  EditMovement,
  EditSelectionBehavior,
  GridSelectionSnapshot,
  WorkbookGridPreviewRange,
  WorkbookGridSurfaceProps,
} from './workbookGridSurfaceTypes.js'

export function resolveWorkbookGridSurfaceDisplaySelection(input: {
  readonly activeHeaderDrag: unknown
  readonly committedCellSelection: GridSelection
  readonly isEditingCell: boolean
  readonly isFillHandleDragging: boolean
  readonly isRangeMoveDragging: boolean
  readonly hasPendingLocalSelection?: boolean | undefined
  readonly renderGridSelection: GridSelection
  readonly renderSelectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null | undefined
  readonly selectedCell: Item
}): GridSelection {
  if (input.isEditingCell) {
    return input.committedCellSelection
  }
  if (input.isFillHandleDragging || input.isRangeMoveDragging || input.activeHeaderDrag) {
    return input.renderGridSelection
  }
  if (input.hasPendingLocalSelection) {
    return input.renderGridSelection
  }
  const currentCell = input.renderGridSelection.current?.cell ?? null
  const currentCellMatchesSelected = currentCell?.[0] === input.selectedCell[0] && currentCell[1] === input.selectedCell[1]
  const renderRangeContainsSelected =
    input.renderSelectionRange !== null &&
    input.renderSelectionRange !== undefined &&
    input.selectedCell[0] >= input.renderSelectionRange.x &&
    input.selectedCell[0] < input.renderSelectionRange.x + input.renderSelectionRange.width &&
    input.selectedCell[1] >= input.renderSelectionRange.y &&
    input.selectedCell[1] < input.renderSelectionRange.y + input.renderSelectionRange.height
  if (!currentCellMatchesSelected) {
    return input.committedCellSelection
  }
  if (currentCellMatchesSelected && !renderRangeContainsSelected) {
    return input.committedCellSelection
  }
  const renderSelectionIsSingleCell =
    input.renderGridSelection.columns.length === 0 &&
    input.renderGridSelection.rows.length === 0 &&
    input.renderSelectionRange?.width === 1 &&
    input.renderSelectionRange.height === 1
  return renderSelectionIsSingleCell ? input.committedCellSelection : input.renderGridSelection
}

export function resolveWorkbookGridSurfaceTextOcclusionRanges(input: {
  readonly gridSelection: GridSelection
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
}): readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] {
  const axisRanges: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] = []
  for (const [start, endExclusive] of input.gridSelection.columns.ranges) {
    axisRanges.push({ x: start, y: 0, width: endExclusive - start, height: MAX_ROWS })
  }
  for (const [start, endExclusive] of input.gridSelection.rows.ranges) {
    axisRanges.push({ x: 0, y: start, width: MAX_COLS, height: endExclusive - start })
  }
  if (axisRanges.length > 0) {
    return axisRanges
  }
  return input.selectionRange ? [input.selectionRange] : []
}

export function resolveWorkbookGridSurfaceDisplayCell(input: {
  readonly committedCell: Item
  readonly displayGridSelection: GridSelection
}): Item {
  return input.displayGridSelection.current?.cell ?? input.committedCell
}

export function WorkbookGridSurface(props: WorkbookGridSurfaceProps) {
  const renderState = useWorkbookGridRenderState({
    engine: props.engine,
    sheetName: props.sheetName,
    selectedAddr: props.selectedAddr,
    selectedCellSnapshot: props.selectedCellSnapshot,
    editorTargetSelection: props.editorTargetSelection,
    editorValue: props.editorValue,
    isEditingCell: props.isEditingCell,
    sheetId: props.sheetId,
    sheetOrdinal: props.sheetOrdinal,
    renderTileSource: props.renderTileSource,
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
  const renderRevisionSnapshot = props.engine.getRenderRevisionSnapshot?.()
  const committedSelectionCol = props.isEditingCell ? renderState.editorCell.col : renderState.selectedCell.col
  const committedSelectionRow = props.isEditingCell ? renderState.editorCell.row : renderState.selectedCell.row
  const committedCellSelection = useMemo(
    () => createGridSelection(committedSelectionCol, committedSelectionRow),
    [committedSelectionCol, committedSelectionRow],
  )
  const hasPendingLocalSelection = renderState.gridRuntimeHost.input.hasPendingLocalSelection({
    currentSelection: renderState.gridSelection,
    externalSnapshot: props.selectionSnapshot,
    sheetName: props.sheetName,
  })
  const displayGridSelection = resolveWorkbookGridSurfaceDisplaySelection({
    activeHeaderDrag: renderState.activeHeaderDrag,
    committedCellSelection,
    hasPendingLocalSelection,
    isEditingCell: props.isEditingCell,
    isFillHandleDragging: renderState.isFillHandleDragging,
    isRangeMoveDragging: renderState.isRangeMoveDragging,
    renderGridSelection: renderState.gridSelection,
    renderSelectionRange: renderState.selectionRange,
    selectedCell: [committedSelectionCol, committedSelectionRow],
  })
  const displaySelectionCell = resolveWorkbookGridSurfaceDisplayCell({
    committedCell: [committedSelectionCol, committedSelectionRow],
    displayGridSelection,
  })
  const displaySelectionCol = displaySelectionCell[0]
  const displaySelectionRow = displaySelectionCell[1]
  const displaySelectionRange = displayGridSelection.current?.range ?? null
  const displayTextOcclusionRanges = useMemo(
    () =>
      resolveWorkbookGridSurfaceTextOcclusionRanges({
        gridSelection: displayGridSelection,
        selectionRange: displaySelectionRange,
      }),
    [displayGridSelection, displaySelectionRange],
  )
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
    onExternalSelectionSync: props.onExternalSelectionSync,
    onSelectionLabelChange: props.onSelectionLabelChange,
    selectionSnapshot: props.selectionSnapshot,
    onToggleBooleanCell: props.onToggleBooleanCell,
    onRowHeightChange: props.onRowHeightChange,
    selectedCellSnapshot: props.selectedCellSnapshot,
    interactionGridSelection: displayGridSelection,
    interactionSelectionCell: displaySelectionCell,
    interactionSelectionRange: displaySelectionRange,
    renderState,
  })
  const focusGrid = renderState.focusGrid
  const lastFocusRequestTokenRef = useRef(props.focusRequestToken)
  useLayoutEffect(() => {
    const focusApiRef = props.focusApiRef
    if (!focusApiRef) {
      return
    }
    focusApiRef.current = focusGrid
    return () => {
      if (focusApiRef.current === focusGrid) {
        focusApiRef.current = null
      }
    }
  }, [focusGrid, props.focusApiRef])
  useLayoutEffect(() => {
    if (props.focusRequestToken === undefined || props.focusRequestToken === lastFocusRequestTokenRef.current) {
      return
    }
    lastFocusRequestTokenRef.current = props.focusRequestToken
    focusGrid()
  }, [focusGrid, props.focusRequestToken])
  useEffect(() => {
    if (props.isEditingCell) {
      return
    }
    if (typeof document !== 'undefined') {
      const activeElement = document.activeElement
      const onDocumentBody = activeElement === document.body || activeElement === document.documentElement || activeElement === null
      if (!onDocumentBody) {
        return
      }
    }
    focusGrid()
  }, [focusGrid, props.isEditingCell, props.selectedAddr, props.sheetName])
  const visibleRange = renderState.visibleRegion.range
  const getCellLocalBounds = renderState.getCellLocalBounds
  const renderHostElement = renderState.hostElement
  const getLiveGeometrySnapshot = renderState.getLiveGeometrySnapshot
  const activeHeaderDrag = renderState.activeHeaderDrag
  const activeResizeColumn = renderState.activeResizeColumn
  const activeResizeRow = renderState.activeResizeRow
  const fillPreviewRange = renderState.fillPreviewRange
  const getPreviewColumnWidth = renderState.getPreviewColumnWidth
  const getPreviewRowHeight = renderState.getPreviewRowHeight
  const hoverCell = renderState.hoverState.cell
  const isRangeMoveDragging = renderState.isRangeMoveDragging
  const activePreviewColumnWidth = activeResizeColumn === null ? null : getPreviewColumnWidth(activeResizeColumn)
  const activePreviewRowHeight = activeResizeRow === null ? null : getPreviewRowHeight(activeResizeRow)
  const resizeGuideColumn = resolveResizeGuideColumn({
    activeResizeColumn,
    cursor: renderState.hoverState.cursor,
    header: renderState.hoverState.header,
  })
  const resizeGuideRow = resolveResizeGuideRow({
    activeResizeRow,
    cursor: renderState.hoverState.cursor,
    header: renderState.hoverState.header,
  })
  const resizeGuideColumnWidth = resizeGuideColumn === activeResizeColumn ? activePreviewColumnWidth : null
  const resizeGuideRowHeight = resizeGuideRow === activeResizeRow ? activePreviewRowHeight : null
  const editorCellCol = renderState.editorCell.col
  const editorCellRow = renderState.editorCell.row
  const suppressedEditorTextCell = useMemo(
    () => (props.isEditingCell ? { col: editorCellCol, row: editorCellRow } : null),
    [editorCellCol, editorCellRow, props.isEditingCell],
  )
  const v2Geometry = useMemo(() => (renderHostElement ? getLiveGeometrySnapshot() : null), [getLiveGeometrySnapshot, renderHostElement])
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
  const showSelectionFillHandle =
    !props.isEditingCell &&
    displaySelectionRange !== null &&
    displayGridSelection.columns.length === 0 &&
    displayGridSelection.rows.length === 0 &&
    fillPreviewRange === null &&
    !isRangeMoveDragging
  const dynamicOverlayBuilder = useCallback(
    (geometry: NonNullable<typeof v2Geometry>) => {
      return buildDynamicGridOverlayBatchV3({
        geometry,
        activeHeaderDrag,
        fillPreviewRange,
        gridSelection: displayGridSelection,
        hoveredCell: hoverCell,
        previewRects,
        selectedCell: [displaySelectionCol, displaySelectionRow],
        selectionOverlayMode: 'fills-only',
        selectionRange: displaySelectionRange,
        showFillHandle: showSelectionFillHandle,
        showHoverOverlay: false,
        showSelectionOverlay: true,
        resizeGuideColumn,
        resizeGuideColumnWidth,
        resizeGuideRow,
        resizeGuideRowHeight,
      })
    },
    [
      activeHeaderDrag,
      displayGridSelection,
      displaySelectionRange,
      fillPreviewRange,
      hoverCell,
      showSelectionFillHandle,
      previewRects,
      resizeGuideColumn,
      resizeGuideColumnWidth,
      resizeGuideRow,
      resizeGuideRowHeight,
      displaySelectionCol,
      displaySelectionRow,
    ],
  )
  const editorTargetAddress =
    props.isEditingCell && props.editorTargetSelection?.sheetName === props.sheetName
      ? props.editorTargetSelection.address
      : props.selectedAddr
  const displayTargetAddress = props.isEditingCell ? editorTargetAddress : formatAddress(displaySelectionRow, displaySelectionCol)

  return (
    <div className="relative flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]">
      <div
        className="sheet-grid-host min-h-0 flex-1 overflow-hidden bg-[var(--wb-surface)] pr-2 pb-2"
        data-column-width-overrides={renderState.columnWidthOverridesAttr}
        data-default-column-width={renderState.gridMetrics.columnWidth}
        data-default-row-height={renderState.gridMetrics.rowHeight}
        data-render-authoritative-revision={renderRevisionSnapshot?.authoritativeRevision ?? ''}
        data-render-local-revision={renderRevisionSnapshot?.localRevision ?? ''}
        data-render-projected-revision={renderRevisionSnapshot?.projectedRevision ?? ''}
        data-render-tile-scene-camera-seq={renderRevisionSnapshot?.tileSceneCameraSeq ?? ''}
        data-render-tile-scene-revision={renderRevisionSnapshot?.tileSceneRevision ?? ''}
        data-row-height-overrides={renderState.rowHeightOverridesAttr}
        data-testid="sheet-grid"
        aria-label={`${props.sheetName} worksheet grid`}
        role="grid"
        style={{
          cursor: renderState.hoverState.cursor,
          fontFamily: WORKBOOK_FONT_SANS,
          fontSize: workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE),
          lineHeight: 1.2,
        }}
        onFocus={interactions.handleHostFocus}
        onKeyDownCapture={interactions.handleHostKeyDownCapture}
        onClickCapture={interactions.handleHostClickCapture}
        onCopyCapture={interactions.handleHostCopyCapture}
        onCutCapture={interactions.handleHostCutCapture}
        onContextMenuCapture={interactions.handleHostContextMenuCapture}
        onPasteCapture={interactions.handleHostPasteCapture}
        onKeyDown={interactions.handleHostKeyDown}
        onDoubleClickCapture={interactions.handleHostDoubleClickCapture}
        onPointerMoveCapture={interactions.handleHostPointerMoveCapture}
        onPointerLeave={interactions.handleHostPointerLeave}
        onPointerDownCapture={interactions.handleHostPointerDownCapture}
        onPointerUpCapture={interactions.handleHostPointerUpCapture}
        ref={renderState.handleHostRef}
      >
        <div
          aria-rowindex={displaySelectionRow + 1}
          className="pointer-events-none absolute h-px w-px overflow-hidden opacity-0"
          role="row"
        >
          <div
            aria-colindex={displaySelectionCol + 1}
            aria-label={`${props.sheetName} ${displayTargetAddress}`}
            aria-selected="true"
            data-testid="sheet-grid-focus-target"
            ref={renderState.focusTargetRef}
            role="gridcell"
            tabIndex={0}
          />
        </div>
        <div
          ref={renderState.scrollViewportRef}
          aria-hidden="true"
          className="absolute inset-0 overflow-auto"
          data-testid="grid-scroll-viewport"
        >
          <div style={{ height: renderState.totalGridHeight, width: renderState.totalGridWidth }} />
        </div>
        <WorkbookPaneRendererV3
          active={renderState.hostElement !== null}
          cameraStore={renderState.gridCameraStore}
          geometry={v2Geometry}
          headerPanes={renderState.headerPanes}
          host={renderState.hostElement}
          overlayBuilder={dynamicOverlayBuilder}
          renderRevisionSnapshot={renderRevisionSnapshot}
          scrollTransformStore={renderState.scrollTransformStore}
          selectionOcclusionRanges={displayTextOcclusionRanges}
          suppressedTextCell={suppressedEditorTextCell}
          tilePanes={renderState.renderTilePanes}
          preloadTilePanes={renderState.preloadDataPanes}
        />
        <GridSelectionVisualOverlay
          geometry={v2Geometry}
          getGeometrySnapshot={getLiveGeometrySnapshot}
          gridSelection={displayGridSelection}
          hoverCell={hoverCell}
          scrollTransformStore={renderState.scrollTransformStore}
          selectedCell={[displaySelectionCol, displaySelectionRow]}
          selectionChromeMode="chrome-only"
          selectionRange={displaySelectionRange}
          showFillHandle={showSelectionFillHandle}
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
          getGeometrySnapshot={getLiveGeometrySnapshot}
          hidden={
            renderState.hostElement === null ||
            props.isEditingCell ||
            !displaySelectionRange ||
            displayGridSelection.columns.length > 0 ||
            displayGridSelection.rows.length > 0 ||
            Boolean(renderState.fillPreviewRange) ||
            renderState.isRangeMoveDragging
          }
          hostHeight={renderState.hostElement?.clientHeight ?? 0}
          hostWidth={renderState.hostElement?.clientWidth ?? 0}
          minX={renderState.gridMetrics.rowMarkerWidth}
          minY={renderState.gridMetrics.headerHeight}
          onPointerDown={interactions.handleFillHandlePointerDown}
          scrollTransformStore={renderState.scrollTransformStore}
          selectionRange={displaySelectionRange}
        />
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
          label={`${props.sheetName}!${editorTargetAddress}`}
          targetSelection={{ sheetName: props.sheetName, address: editorTargetAddress }}
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
