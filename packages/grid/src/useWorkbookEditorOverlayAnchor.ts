import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridEditorPresentation } from './gridPresentation.js'
import type { Rectangle } from './gridTypes.js'
import { GridEditorAnchorRuntime, type GridEditorAnchorOverlayStyle } from './runtime/gridEditorAnchorRuntime.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

export interface WorkbookEditorOverlayAnchorState {
  readonly editorPresentation: GridEditorPresentation
  readonly editorTextAlign: 'left' | 'right'
  readonly overlayStyle: GridEditorAnchorOverlayStyle
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
  const editorAnchorRuntimeRef = useRef<GridEditorAnchorRuntime | null>(null)
  editorAnchorRuntimeRef.current ??= new GridEditorAnchorRuntime()
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined)

  const refreshOverlayBounds = useCallback(
    (options?: { readonly commitReactState?: boolean }) => {
      const next = editorAnchorRuntimeRef.current!.refreshOverlayBounds({
        col: selectedCol,
        row: selectedRow,
        getCellLocalBounds,
        gridCameraStore,
        hostElement,
      })
      if (!next) {
        return
      }
      if (options?.commitReactState === false) {
        return
      }
      setOverlayBounds((current) => {
        return editorAnchorRuntimeRef.current!.resolveCommittedBounds(current, next)
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

  const overlayStyle = useMemo(
    () => editorAnchorRuntimeRef.current!.resolveOverlayStyle(isEditingCell, overlayBounds),
    [isEditingCell, overlayBounds],
  )
  const editorPresentation = useMemo(
    () =>
      editorAnchorRuntimeRef.current!.resolvePresentation({
        engine,
        selectedCellSnapshot,
      }),
    [engine, selectedCellSnapshot],
  )
  const editorTextAlign = useMemo<'left' | 'right'>(() => editorAnchorRuntimeRef.current!.resolveTextAlign(editorValue), [editorValue])

  return {
    editorPresentation,
    editorTextAlign,
    overlayStyle,
  }
}
