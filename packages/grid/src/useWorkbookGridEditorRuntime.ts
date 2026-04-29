import { useCallback, type MutableRefObject } from 'react'
import type { CellSnapshot } from '@bilig/protocol'
import { useWorkbookColumnAutofit } from './useWorkbookColumnAutofit.js'
import { useWorkbookEditorOverlayAnchor, type WorkbookEditorOverlayAnchorState } from './useWorkbookEditorOverlayAnchor.js'
import type { GridMetrics } from './gridMetrics.js'
import type { GridEngineLike } from './grid-engine.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { Rectangle } from './gridTypes.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'

export interface WorkbookGridEditorRuntimeState extends WorkbookEditorOverlayAnchorState {
  readonly computeAutofitColumnWidth: (columnIndex: number) => number
  readonly focusGrid: () => void
}

export function useWorkbookGridEditorRuntime(input: {
  readonly editorFontSize: string
  readonly editorValue: string
  readonly engine: GridEngineLike
  readonly focusTargetRef: MutableRefObject<HTMLDivElement | null>
  readonly freezeRows: number
  readonly getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly getVisibleRegion: () => VisibleRegionState
  readonly gridCameraStore: GridCameraStore
  readonly gridMetrics: GridMetrics
  readonly headerFontStyle: string
  readonly hostElement: HTMLElement | null
  readonly hostRef: MutableRefObject<HTMLDivElement | null>
  readonly isEditingCell: boolean
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly selectedCell: {
    readonly col: number
    readonly row: number
  }
  readonly selectedCellSnapshot: CellSnapshot
  readonly sheetName: string
}): WorkbookGridEditorRuntimeState {
  const {
    editorFontSize,
    editorValue,
    engine,
    focusTargetRef,
    freezeRows,
    getCellEditorSeed,
    getCellLocalBounds,
    getVisibleRegion,
    gridCameraStore,
    gridMetrics,
    headerFontStyle,
    hostElement,
    hostRef,
    isEditingCell,
    scrollTransformStore,
    selectedCell,
    selectedCellSnapshot,
    sheetName,
  } = input

  const overlayAnchor = useWorkbookEditorOverlayAnchor({
    editorValue,
    engine,
    getCellLocalBounds,
    gridCameraStore,
    hostElement,
    isEditingCell,
    scrollTransformStore,
    selectedCellSnapshot,
    selectedCol: selectedCell.col,
    selectedRow: selectedCell.row,
  })

  const focusGrid = useCallback(() => {
    const activeElement = typeof document === 'undefined' ? null : document.activeElement
    if (
      (activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement) &&
      activeElement.dataset['testid'] === 'cell-editor-input'
    ) {
      return
    }
    const focusTarget = focusTargetRef.current
    if (focusTarget) {
      focusTarget.focus({ preventScroll: true })
      return
    }
    hostRef.current?.focus({ preventScroll: true })
  }, [focusTargetRef, hostRef])

  const computeAutofitColumnWidth = useWorkbookColumnAutofit({
    editorFontSize,
    engine,
    freezeRows,
    getCellEditorSeed,
    getVisibleRegion,
    gridMetrics,
    headerFontStyle,
    selectedCell,
    selectedCellSnapshot,
    sheetName,
  })

  return {
    ...overlayAnchor,
    computeAutofitColumnWidth,
    focusGrid,
  }
}
