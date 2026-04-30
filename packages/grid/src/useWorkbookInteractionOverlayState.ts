import { useLayoutEffect, useMemo, useSyncExternalStore, type Dispatch, type SetStateAction } from 'react'
import type { GridHoverState } from './gridHover.js'
import type { HeaderSelection } from './gridPointer.js'
import { GridInteractionOverlayRuntime } from './runtime/gridInteractionOverlayRuntime.js'
import type { GridSelection, Rectangle } from './gridTypes.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

export interface WorkbookInteractionOverlayState {
  readonly activeHeaderDrag: HeaderSelection | null
  readonly fillPreviewBounds: Rectangle | undefined
  readonly fillPreviewRange: Rectangle | null
  readonly gridSelection: GridSelection
  readonly hoverState: GridHoverState
  readonly isEntireSheetSelected: boolean
  readonly isFillHandleDragging: boolean
  readonly isRangeMoveDragging: boolean
  readonly requiresLiveViewportState: boolean
  readonly selectionRange: Rectangle | null
  readonly setActiveHeaderDrag: Dispatch<SetStateAction<HeaderSelection | null>>
  readonly setFillPreviewRange: Dispatch<SetStateAction<Rectangle | null>>
  readonly setGridSelection: Dispatch<SetStateAction<GridSelection>>
  readonly setHoverState: Dispatch<SetStateAction<GridHoverState>>
  readonly setIsFillHandleDragging: Dispatch<SetStateAction<boolean>>
  readonly setIsRangeMoveDragging: Dispatch<SetStateAction<boolean>>
}

export function useWorkbookInteractionOverlayState(input: {
  readonly activeResizeColumn: number | null
  readonly activeResizeRow: number | null
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridRuntimeHost?: GridRuntimeHost | undefined
  readonly hasColumnResizePreview: boolean
  readonly hasRowResizePreview: boolean
  readonly isEditingCell: boolean
  readonly selectedCol: number
  readonly selectedRow: number
  readonly visibleRange: Rectangle
}): WorkbookInteractionOverlayState {
  const {
    activeResizeColumn,
    activeResizeRow,
    getCellLocalBounds,
    gridRuntimeHost,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol,
    selectedRow,
    visibleRange,
  } = input
  const runtime = useMemo(() => gridRuntimeHost?.interactionOverlays ?? new GridInteractionOverlayRuntime(), [gridRuntimeHost])
  const snapshot = useSyncExternalStore(
    (listener) => runtime.subscribe(listener),
    () => runtime.snapshot(),
    () => runtime.snapshot(),
  )

  useLayoutEffect(() => {
    runtime.syncSelectedCell({ selectedCol, selectedRow })
  }, [runtime, selectedCol, selectedRow])

  const resolvedState = useMemo(
    () =>
      runtime.resolveState({
        activeResizeColumn,
        activeResizeRow,
        getCellLocalBounds,
        hasColumnResizePreview,
        hasRowResizePreview,
        isEditingCell,
        snapshot,
        visibleRange,
      }),
    [
      activeResizeColumn,
      activeResizeRow,
      getCellLocalBounds,
      hasColumnResizePreview,
      hasRowResizePreview,
      isEditingCell,
      runtime,
      snapshot,
      visibleRange,
    ],
  )

  return {
    ...resolvedState,
    setActiveHeaderDrag: (action) => runtime.setActiveHeaderDrag(action),
    setFillPreviewRange: (action) => runtime.setFillPreviewRange(action),
    setGridSelection: (action) => runtime.setGridSelection(action),
    setHoverState: (action) => runtime.setHoverState(action),
    setIsFillHandleDragging: (action) => runtime.setIsFillHandleDragging(action),
    setIsRangeMoveDragging: (action) => runtime.setIsRangeMoveDragging(action),
  }
}
