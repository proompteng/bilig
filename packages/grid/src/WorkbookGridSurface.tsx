import { CellEditorOverlay } from "./CellEditorOverlay.js";
import { GridGpuSurface } from "./GridGpuSurface.js";
import { GridTextOverlay } from "./GridTextOverlay.js";
import { useWorkbookGridInteractions } from "./useWorkbookGridInteractions.js";
import { useWorkbookGridRenderState } from "./useWorkbookGridRenderState.js";
import type { WorkbookGridSurfaceProps } from "./workbookGridSurfaceTypes.js";
export { hasSelectionTargetChanged } from "./workbookGridViewport.js";
export type {
  EditMovement,
  EditSelectionBehavior,
  SheetGridViewportSubscription,
  WorkbookGridSurfaceProps,
} from "./workbookGridSurfaceTypes.js";

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
    onVisibleViewportChange: props.onVisibleViewportChange,
    onColumnWidthChange: props.onColumnWidthChange,
  });
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
    onSelect: props.onSelect,
    onSelectionLabelChange: props.onSelectionLabelChange,
    onToggleBooleanCell: props.onToggleBooleanCell,
    renderState,
  });

  return (
    <div className="relative flex min-h-0 flex-1 flex-col bg-[var(--wb-surface)]">
      <div
        className="sheet-grid-host min-h-0 flex-1 bg-[var(--wb-surface)] pr-2 pb-2"
        data-column-width-overrides={renderState.columnWidthOverridesAttr}
        data-default-column-width={renderState.gridMetrics.columnWidth}
        data-testid="sheet-grid"
        role="grid"
        style={{ cursor: renderState.hoverState.cursor }}
        onFocus={interactions.handleHostFocus}
        onKeyDownCapture={interactions.handleHostKeyDownCapture}
        onCopyCapture={interactions.handleHostCopyCapture}
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
        >
          <div style={{ height: renderState.totalGridHeight, width: renderState.totalGridWidth }} />
        </div>
        <GridGpuSurface
          host={renderState.hostElement}
          scene={renderState.gpuScene}
          onActiveChange={renderState.setIsWebGpuActive}
        />
        <GridTextOverlay
          active={renderState.hostElement !== null && renderState.isWebGpuActive}
          scene={renderState.textScene}
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
              color: renderState.isEntireSheetSelected ? "var(--wb-accent)" : "currentColor",
              transform: "translate(2px, 1px)",
            }}
          />
        </button>
        {renderState.fillHandleBounds ? (
          <button
            aria-label="Fill handle"
            className="absolute z-30 cursor-crosshair rounded-full border-0 bg-[#1f7a43] shadow-[0_0_0_1px_rgba(31,122,67,0.45)] outline-none"
            data-grid-fill-handle="true"
            onClick={(event) => {
              event.preventDefault();
              event.stopPropagation();
            }}
            onPointerDown={interactions.handleFillHandlePointerDown}
            style={{
              height: renderState.fillHandleBounds.height,
              left: renderState.fillHandleBounds.x,
              touchAction: "none",
              top: renderState.fillHandleBounds.y,
              width: renderState.fillHandleBounds.width,
            }}
            type="button"
          />
        ) : null}
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
        <div className="pointer-events-none absolute inset-0 z-[1]" />
      </div>
      {props.isEditingCell && renderState.overlayStyle ? (
        <CellEditorOverlay
          label={`${props.sheetName}!${props.selectedAddr}`}
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
  );
}
