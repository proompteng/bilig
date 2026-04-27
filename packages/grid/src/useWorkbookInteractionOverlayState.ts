import { useLayoutEffect, useMemo, useState, type Dispatch, type SetStateAction } from 'react'
import { resolveFillHandlePreviewBounds } from './gridFillHandle.js'
import type { GridHoverState } from './gridHover.js'
import type { HeaderSelection } from './gridPointer.js'
import { createGridSelection, isSheetSelection } from './gridSelection.js'
import type { GridSelection, Rectangle } from './gridTypes.js'
import { resolveRequiresLiveViewportState } from './useGridSelectionState.js'

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
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol,
    selectedRow,
    visibleRange,
  } = input
  const [fillPreviewRange, setFillPreviewRange] = useState<Rectangle | null>(null)
  const [isFillHandleDragging, setIsFillHandleDragging] = useState(false)
  const [isRangeMoveDragging, setIsRangeMoveDragging] = useState(false)
  const [hoverState, setHoverState] = useState<GridHoverState>({
    cell: null,
    header: null,
    cursor: 'default',
  })
  const [activeHeaderDrag, setActiveHeaderDrag] = useState<HeaderSelection | null>(null)
  const [gridSelection, setGridSelection] = useState<GridSelection>(() => createGridSelection(selectedCol, selectedRow))

  useLayoutEffect(() => {
    setGridSelection((current) => {
      if (
        current.columns.length > 0 ||
        current.rows.length > 0 ||
        current.current?.range.width !== 1 ||
        current.current.range.height !== 1
      ) {
        return current
      }
      if (current.current.cell[0] === selectedCol && current.current.cell[1] === selectedRow) {
        return current
      }
      return createGridSelection(selectedCol, selectedRow)
    })
  }, [selectedCol, selectedRow])

  const selectionRange = gridSelection.current?.range ?? null
  const fillPreviewBounds = useMemo<Rectangle | undefined>(() => {
    if (!fillPreviewRange) {
      return undefined
    }
    return resolveFillHandlePreviewBounds({
      previewRange: fillPreviewRange,
      visibleRange,
      hostBounds: { left: 0, top: 0 },
      getCellBounds: getCellLocalBounds,
    })
  }, [fillPreviewRange, getCellLocalBounds, visibleRange])

  const requiresLiveViewportState = resolveRequiresLiveViewportState({
    fillPreviewActive: fillPreviewRange !== null,
    hasActiveHeaderDrag: activeHeaderDrag !== null,
    hasActiveResizeColumn: activeResizeColumn !== null,
    hasActiveResizeRow: activeResizeRow !== null,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    isFillHandleDragging,
  })

  return {
    activeHeaderDrag,
    fillPreviewBounds,
    fillPreviewRange,
    gridSelection,
    hoverState,
    isEntireSheetSelected: isSheetSelection(gridSelection),
    isFillHandleDragging,
    isRangeMoveDragging,
    requiresLiveViewportState,
    selectionRange,
    setActiveHeaderDrag,
    setFillPreviewRange,
    setGridSelection,
    setHoverState,
    setIsFillHandleDragging,
    setIsRangeMoveDragging,
  }
}
