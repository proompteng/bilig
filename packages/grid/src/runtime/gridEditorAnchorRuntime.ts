import type { CellSnapshot } from '@bilig/protocol'
import { snapshotToRenderCell } from '../gridCells.js'
import { applyEditorOverlayBounds, resolveEditorOverlayScreenBounds } from '../gridEditorOverlayGeometry.js'
import type { GridEngineLike } from '../grid-engine.js'
import { getEditorPresentation, getEditorTextAlign, getOverlayStyle, type GridEditorPresentation } from '../gridPresentation.js'
import type { Rectangle } from '../gridTypes.js'
import { sameBounds } from '../useGridOverlayState.js'
import type { GridCameraStore } from './gridCameraStore.js'

export type GridEditorAnchorOverlayStyle = ReturnType<typeof getOverlayStyle>

export interface GridEditorAnchorRefreshInput {
  readonly col: number
  readonly row: number
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridCameraStore: GridCameraStore
  readonly hostElement: HTMLElement | null
}

export interface GridEditorAnchorPresentationInput {
  readonly engine: GridEngineLike
  readonly selectedCellSnapshot: CellSnapshot
}

export class GridEditorAnchorRuntime {
  refreshOverlayBounds(input: GridEditorAnchorRefreshInput): Rectangle | null {
    const next = resolveEditorOverlayScreenBounds({
      col: input.col,
      row: input.row,
      geometry: input.gridCameraStore.getSnapshot(),
      getCellLocalBounds: input.getCellLocalBounds,
      hostElement: input.hostElement,
    })
    if (!next) {
      return null
    }
    applyEditorOverlayBounds(next)
    return next
  }

  resolveCommittedBounds(current: Rectangle | undefined, next: Rectangle): Rectangle | undefined {
    return sameBounds(current, next) ? current : next
  }

  resolveOverlayStyle(isEditingCell: boolean, overlayBounds: Rectangle | undefined): GridEditorAnchorOverlayStyle {
    return getOverlayStyle(isEditingCell, overlayBounds)
  }

  resolvePresentation(input: GridEditorAnchorPresentationInput): GridEditorPresentation {
    const selectedCellStyle = input.engine.getCellStyle(input.selectedCellSnapshot.styleId)
    const renderCell = snapshotToRenderCell(input.selectedCellSnapshot, selectedCellStyle)
    return getEditorPresentation({
      renderCell,
      fillColor: selectedCellStyle?.fill?.backgroundColor,
    })
  }

  resolveTextAlign(editorValue: string): 'left' | 'right' {
    return getEditorTextAlign(editorValue)
  }
}
