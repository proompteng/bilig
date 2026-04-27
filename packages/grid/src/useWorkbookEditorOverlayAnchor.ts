import { useCallback, useEffect, useLayoutEffect, useMemo, useState } from 'react'
import type { CellSnapshot } from '@bilig/protocol'
import { snapshotToRenderCell } from './gridCells.js'
import { applyEditorOverlayBounds, resolveEditorOverlayScreenBounds } from './gridEditorOverlayGeometry.js'
import type { GridEngineLike } from './grid-engine.js'
import { getEditorPresentation, getEditorTextAlign, getOverlayStyle, type GridEditorPresentation } from './gridPresentation.js'
import type { Rectangle } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import { sameBounds } from './useGridOverlayState.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

export interface WorkbookEditorOverlayAnchorState {
  readonly editorPresentation: GridEditorPresentation
  readonly editorTextAlign: 'left' | 'right'
  readonly overlayStyle: ReturnType<typeof getOverlayStyle>
}

export function useWorkbookEditorOverlayAnchor(input: {
  readonly editorValue: string
  readonly engine: GridEngineLike
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridCameraStore: GridCameraStore
  readonly hostElement: HTMLElement | null
  readonly isEditingCell: boolean
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly selectedCellSnapshot: CellSnapshot
  readonly selectedCol: number
  readonly selectedRow: number
}): WorkbookEditorOverlayAnchorState {
  const {
    editorValue,
    engine,
    getCellLocalBounds,
    gridCameraStore,
    hostElement,
    isEditingCell,
    scrollTransformStore,
    selectedCellSnapshot,
    selectedCol,
    selectedRow,
  } = input
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined)

  const refreshOverlayBounds = useCallback(
    (options?: { readonly commitReactState?: boolean }) => {
      const next = resolveEditorOverlayScreenBounds({
        col: selectedCol,
        row: selectedRow,
        geometry: gridCameraStore.getSnapshot(),
        getCellLocalBounds,
        hostElement,
      })
      if (!next) {
        return
      }
      applyEditorOverlayBounds(next)
      if (options?.commitReactState === false) {
        return
      }
      setOverlayBounds((current) => {
        return sameBounds(current, next) ? current : next
      })
    },
    [getCellLocalBounds, gridCameraStore, hostElement, selectedCol, selectedRow],
  )

  useLayoutEffect(() => {
    if (!isEditingCell) {
      setOverlayBounds(undefined)
      return
    }

    refreshOverlayBounds({ commitReactState: true })
    const frame = window.requestAnimationFrame(() => refreshOverlayBounds({ commitReactState: true }))
    const unsubscribeScrollTransform = scrollTransformStore.subscribe(() => refreshOverlayBounds({ commitReactState: false }))
    return () => {
      window.cancelAnimationFrame(frame)
      unsubscribeScrollTransform()
    }
  }, [isEditingCell, refreshOverlayBounds, scrollTransformStore])

  useEffect(() => {
    if (!isEditingCell) {
      return
    }
    const handleWindowResize = () => refreshOverlayBounds()
    window.addEventListener('resize', handleWindowResize)
    return () => {
      window.removeEventListener('resize', handleWindowResize)
    }
  }, [isEditingCell, refreshOverlayBounds])

  const overlayStyle = useMemo(() => getOverlayStyle(isEditingCell, overlayBounds), [isEditingCell, overlayBounds])
  const editorPresentation = useMemo(() => {
    const selectedCellStyle = engine.getCellStyle(selectedCellSnapshot.styleId)
    const renderCell = snapshotToRenderCell(selectedCellSnapshot, selectedCellStyle)
    return getEditorPresentation({
      renderCell,
      fillColor: selectedCellStyle?.fill?.backgroundColor,
    })
  }, [engine, selectedCellSnapshot])
  const editorTextAlign = useMemo<'left' | 'right'>(() => getEditorTextAlign(editorValue), [editorValue])

  return {
    editorPresentation,
    editorTextAlign,
    overlayStyle,
  }
}
