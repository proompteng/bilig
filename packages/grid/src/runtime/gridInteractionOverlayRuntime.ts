import { resolveFillHandlePreviewBounds } from '../gridFillHandle.js'
import type { GridHoverState } from '../gridHover.js'
import type { HeaderSelection } from '../gridPointer.js'
import { createGridSelection, isSheetSelection, normalizeGridSelection } from '../gridSelection.js'
import type { GridSelection, Rectangle } from '../gridTypes.js'
import { resolveRequiresLiveViewportState } from '../useGridSelectionState.js'

type RuntimeStateAction<T> = T | ((current: T) => T)

export interface GridInteractionOverlaySnapshot {
  readonly activeHeaderDrag: HeaderSelection | null
  readonly fillPreviewRange: Rectangle | null
  readonly gridSelection: GridSelection
  readonly hoverState: GridHoverState
  readonly isFillHandleDragging: boolean
  readonly isRangeMoveDragging: boolean
}

export interface GridInteractionOverlayResolvedState extends GridInteractionOverlaySnapshot {
  readonly fillPreviewBounds: Rectangle | undefined
  readonly isEntireSheetSelected: boolean
  readonly requiresLiveViewportState: boolean
  readonly selectionRange: Rectangle | null
}

const DEFAULT_HOVER_STATE: GridHoverState = Object.freeze({
  cell: null,
  cursor: 'default',
  header: null,
})

export class GridInteractionOverlayRuntime {
  private readonly listeners = new Set<() => void>()
  private snapshotValue: GridInteractionOverlaySnapshot = {
    activeHeaderDrag: null,
    fillPreviewRange: null,
    gridSelection: createGridSelection(0, 0),
    hoverState: DEFAULT_HOVER_STATE,
    isFillHandleDragging: false,
    isRangeMoveDragging: false,
  }

  snapshot(): GridInteractionOverlaySnapshot {
    return this.snapshotValue
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  setActiveHeaderDrag(action: RuntimeStateAction<HeaderSelection | null>): void {
    this.update('activeHeaderDrag', action)
  }

  setFillPreviewRange(action: RuntimeStateAction<Rectangle | null>): void {
    this.update('fillPreviewRange', action)
  }

  setGridSelection(action: RuntimeStateAction<GridSelection>): void {
    const current = this.snapshotValue.gridSelection
    const next = typeof action === 'function' ? (action as (current: GridSelection) => GridSelection)(current) : action
    const normalized = normalizeGridSelection(next)
    if (Object.is(current, normalized)) {
      return
    }
    this.snapshotValue = {
      ...this.snapshotValue,
      gridSelection: normalized,
    }
    this.emit()
  }

  setHoverState(action: RuntimeStateAction<GridHoverState>): void {
    this.update('hoverState', action)
  }

  setIsFillHandleDragging(action: RuntimeStateAction<boolean>): void {
    this.update('isFillHandleDragging', action)
  }

  setIsRangeMoveDragging(action: RuntimeStateAction<boolean>): void {
    this.update('isRangeMoveDragging', action)
  }

  syncSelectedCell(input: { readonly selectedCol: number; readonly selectedRow: number }): void {
    const current = this.snapshotValue.gridSelection
    if (current.columns.length > 0 || current.rows.length > 0 || current.current?.range.width !== 1 || current.current.range.height !== 1) {
      return
    }
    if (current.current.cell[0] === input.selectedCol && current.current.cell[1] === input.selectedRow) {
      return
    }
    this.setGridSelection(createGridSelection(input.selectedCol, input.selectedRow))
  }

  resolveState(input: {
    readonly activeResizeColumn: number | null
    readonly activeResizeRow: number | null
    readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
    readonly hasColumnResizePreview: boolean
    readonly hasRowResizePreview: boolean
    readonly isEditingCell: boolean
    readonly snapshot: GridInteractionOverlaySnapshot
    readonly visibleRange: Rectangle
  }): GridInteractionOverlayResolvedState {
    const fillPreviewBounds = input.snapshot.fillPreviewRange
      ? resolveFillHandlePreviewBounds({
          previewRange: input.snapshot.fillPreviewRange,
          visibleRange: input.visibleRange,
          hostBounds: { left: 0, top: 0 },
          getCellBounds: input.getCellLocalBounds,
        })
      : undefined
    return {
      ...input.snapshot,
      fillPreviewBounds,
      isEntireSheetSelected: isSheetSelection(input.snapshot.gridSelection),
      requiresLiveViewportState: resolveRequiresLiveViewportState({
        fillPreviewActive: input.snapshot.fillPreviewRange !== null,
        hasActiveHeaderDrag: input.snapshot.activeHeaderDrag !== null,
        hasActiveResizeColumn: input.activeResizeColumn !== null,
        hasActiveResizeRow: input.activeResizeRow !== null,
        hasColumnResizePreview: input.hasColumnResizePreview,
        hasRowResizePreview: input.hasRowResizePreview,
        isEditingCell: input.isEditingCell,
        isFillHandleDragging: input.snapshot.isFillHandleDragging,
      }),
      selectionRange: input.snapshot.gridSelection.current?.range ?? null,
    }
  }

  private update<Key extends keyof GridInteractionOverlaySnapshot>(
    key: Key,
    action: RuntimeStateAction<GridInteractionOverlaySnapshot[Key]>,
  ): void {
    const current = this.snapshotValue[key]
    const next =
      typeof action === 'function'
        ? (action as (current: GridInteractionOverlaySnapshot[Key]) => GridInteractionOverlaySnapshot[Key])(current)
        : action
    if (Object.is(current, next)) {
      return
    }
    this.snapshotValue = {
      ...this.snapshotValue,
      [key]: next,
    }
    this.emit()
  }

  private emit(): void {
    this.listeners.forEach((listener) => {
      listener()
    })
  }
}
