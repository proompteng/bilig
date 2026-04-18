import { formatAddress } from '@bilig/formula'
import { createColumnSliceSelection, createGridSelection, createRowSliceSelection } from './gridSelection.js'
import { resolveBodyDragSelection, resolveBodyPointerUpResult, resolveHeaderDragSelection } from './gridDragSelection.js'
import { resolveBodyDoubleClickIntent } from './gridEventPolicy.js'
import { resolveColumnResizeTarget, type HeaderSelection, type PointerGeometry, type VisibleRegionState } from './gridPointer.js'
import {
  beginGridBodyPointerInteraction,
  beginGridHeaderDrag,
  finishGridColumnResize,
  resetGridPointerInteraction,
  scheduleGridPointerInteractionReset,
  startGridColumnResize,
  type GridInteractionStateRefs,
} from './gridInteractionState.js'
import type { GridSelection, Item } from './gridTypes.js'

interface PointerEventLike {
  clientX: number
  clientY: number
  preventDefault(): void
  stopPropagation(): void
}

interface PointerDownEventLike extends PointerEventLike {
  button: number
  shiftKey: boolean
}

interface PointerMoveEventLike {
  clientX: number
  clientY: number
  buttons: number
}

interface GridInteractionCommonOptions {
  interactionState: GridInteractionStateRefs
  isEditingCell: boolean
  onCommitEdit(this: void): void
  onSelectionChange(this: void, selection: GridSelection): void
  selectedCell: Item
  setGridSelection(this: void, selection: GridSelection): void
  visibleRegion: VisibleRegionState
}

interface HandleGridBodyDoubleClickOptions extends GridInteractionCommonOptions {
  event: PointerEventLike
  columnWidths: Readonly<Record<number, number>>
  defaultColumnWidth: number
  lastBodyClickCell: Item | null
  onAutofitColumn?: ((this: void, columnIndex: number, fallbackWidth: number) => void | Promise<void>) | undefined
  applyColumnWidth(this: void, columnIndex: number, width: number): void
  computeAutofitColumnWidth(this: void, columnIndex: number): number
  beginEditAt(this: void, address: string): void
  resolvePointerGeometry(this: void, region?: VisibleRegionState): PointerGeometry | null
  resolvePointerCell(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): Item | null
}

interface HandleGridPointerMoveOptions extends GridInteractionCommonOptions {
  event: PointerMoveEventLike
  dragAnchorCell: Item | null
  dragHeaderSelection: HeaderSelection | null
  dragPointerCell: Item | null
  dragViewport: VisibleRegionState | null
  dragGeometry: PointerGeometry | null
  resolvePointerCell(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): Item | null
  resolveHeaderSelectionForPointerDrag(
    this: void,
    kind: HeaderSelection['kind'],
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): HeaderSelection | null
}

interface HandleGridPointerDownOptions extends GridInteractionCommonOptions {
  event: PointerDownEventLike
  columnWidths: Readonly<Record<number, number>>
  defaultColumnWidth: number
  focusGrid(this: void): void
  resolvePointerGeometry(this: void, region?: VisibleRegionState): PointerGeometry | null
  resolveColumnResizeTargetAtPointer(
    this: void,
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry: PointerGeometry,
    columnWidths: Readonly<Record<number, number>>,
    defaultWidth: number,
  ): number | null
  resolveHeaderSelectionAtPointer(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): HeaderSelection | null
  resolvePointerCell(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): Item | null
}

interface HandleGridPointerUpOptions extends GridInteractionCommonOptions {
  event: PointerEventLike
  dragAnchorCell: Item | null
  dragDidMove: boolean
  dragHeaderSelection: HeaderSelection | null
  dragPointerCell: Item | null
  dragViewport: VisibleRegionState | null
  dragGeometry: PointerGeometry | null
  lastBodyClickCellRef: { current: Item | null }
  postDragSelectionExpiryRef: { current: number }
  resolvePointerCell(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): Item | null
  resolveHeaderSelectionForPointerDrag(
    this: void,
    kind: HeaderSelection['kind'],
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ): HeaderSelection | null
}

export function handleGridBodyDoubleClick({
  event,
  columnWidths,
  defaultColumnWidth,
  interactionState,
  lastBodyClickCell,
  onAutofitColumn,
  applyColumnWidth,
  computeAutofitColumnWidth,
  beginEditAt,
  onSelectionChange,
  resolvePointerGeometry,
  resolvePointerCell,
  setGridSelection,
  visibleRegion,
}: HandleGridBodyDoubleClickOptions): void {
  const activeGeometry = resolvePointerGeometry(visibleRegion)
  if (!activeGeometry) {
    return
  }
  const doubleClickIntent = resolveBodyDoubleClickIntent({
    resizeTarget: resolveColumnResizeTarget(event.clientX, event.clientY, visibleRegion, activeGeometry, columnWidths, defaultColumnWidth),
    bodyCell: resolvePointerCell(event.clientX, event.clientY, visibleRegion, activeGeometry),
    lastBodyClickCell,
  })
  if (doubleClickIntent.kind === 'ignore') {
    return
  }
  event.preventDefault()
  event.stopPropagation()
  if (doubleClickIntent.kind === 'edit-cell') {
    const editAddress = formatAddress(doubleClickIntent.cell[1], doubleClickIntent.cell[0])
    const nextSelection = createGridSelection(doubleClickIntent.cell[0], doubleClickIntent.cell[1])
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
    beginEditAt(editAddress)
    return
  }
  interactionState.columnResizeActiveRef.current = false
  resetGridPointerInteraction(interactionState)
  if (onAutofitColumn) {
    void Promise.resolve(onAutofitColumn(doubleClickIntent.columnIndex, computeAutofitColumnWidth(doubleClickIntent.columnIndex)))
    return
  }
  applyColumnWidth(doubleClickIntent.columnIndex, computeAutofitColumnWidth(doubleClickIntent.columnIndex))
}

export function handleGridPointerMove({
  event,
  dragAnchorCell,
  dragHeaderSelection,
  dragPointerCell,
  dragViewport,
  dragGeometry,
  interactionState,
  resolvePointerCell,
  resolveHeaderSelectionForPointerDrag,
  selectedCell,
  setGridSelection,
  visibleRegion,
}: HandleGridPointerMoveOptions): void {
  if (interactionState.columnResizeActiveRef.current) {
    return
  }
  if ((event.buttons & 1) !== 1) {
    return
  }
  if (dragHeaderSelection) {
    const nextHeader = resolveHeaderSelectionForPointerDrag(
      dragHeaderSelection.kind,
      event.clientX,
      event.clientY,
      dragViewport ?? visibleRegion,
      dragGeometry,
    )
    if (!nextHeader || nextHeader.index === dragHeaderSelection.index) {
      return
    }
    interactionState.dragPointerCellRef.current = null
    interactionState.dragDidMoveRef.current = true
    setGridSelection(resolveHeaderDragSelection(dragHeaderSelection, nextHeader.index, selectedCell).selection)
    return
  }
  if (dragAnchorCell === null) {
    return
  }
  const pointerCell = resolvePointerCell(event.clientX, event.clientY, dragViewport ?? visibleRegion, dragGeometry)
  if (!pointerCell) {
    return
  }
  if (dragPointerCell && dragPointerCell[0] === pointerCell[0] && dragPointerCell[1] === pointerCell[1]) {
    return
  }
  interactionState.dragPointerCellRef.current = pointerCell
  if (pointerCell[0] !== dragAnchorCell[0] || pointerCell[1] !== dragAnchorCell[1]) {
    interactionState.dragDidMoveRef.current = true
    interactionState.ignoreNextPointerSelectionRef.current = false
    setGridSelection(resolveBodyDragSelection(dragAnchorCell, pointerCell))
  }
}

export function handleGridPointerDown({
  event,
  columnWidths,
  defaultColumnWidth,
  focusGrid,
  interactionState,
  isEditingCell,
  onCommitEdit,
  onSelectionChange,
  resolvePointerGeometry,
  resolveColumnResizeTargetAtPointer,
  resolveHeaderSelectionAtPointer,
  resolvePointerCell,
  selectedCell,
  setGridSelection,
  visibleRegion,
}: HandleGridPointerDownOptions): void {
  if (event.button !== 0) {
    return
  }
  const activeGeometry = resolvePointerGeometry(visibleRegion)
  if (
    activeGeometry &&
    resolveColumnResizeTargetAtPointer(event.clientX, event.clientY, visibleRegion, activeGeometry, columnWidths, defaultColumnWidth) !==
      null
  ) {
    resetGridPointerInteraction(interactionState)
    return
  }
  const headerSelection = resolveHeaderSelectionAtPointer(event.clientX, event.clientY)
  if (headerSelection) {
    if (isEditingCell) {
      onCommitEdit()
    }
    beginGridHeaderDrag(interactionState, headerSelection, activeGeometry, visibleRegion)
    if (headerSelection.kind === 'row') {
      interactionState.ignoreNextPointerSelectionRef.current = true
      const nextSelection = createRowSliceSelection(selectedCell[0], headerSelection.index, headerSelection.index)
      setGridSelection(nextSelection)
      onSelectionChange(nextSelection)
      focusGrid()
      return
    }
    interactionState.ignoreNextPointerSelectionRef.current = true
    const nextSelection = createColumnSliceSelection(headerSelection.index, headerSelection.index, selectedCell[1])
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
    focusGrid()
    return
  }
  const pointerCell = resolvePointerCell(event.clientX, event.clientY)
  beginGridBodyPointerInteraction(interactionState, pointerCell, activeGeometry, visibleRegion)
  if (pointerCell) {
    const anchorCell: Item = event.shiftKey ? selectedCell : pointerCell
    interactionState.dragAnchorCellRef.current = anchorCell
    interactionState.dragPointerCellRef.current = pointerCell
    interactionState.ignoreNextPointerSelectionRef.current = true
    if (isEditingCell) {
      onCommitEdit()
    }
    const nextSelection = event.shiftKey
      ? resolveBodyDragSelection(anchorCell, pointerCell)
      : createGridSelection(pointerCell[0], pointerCell[1])
    setGridSelection(nextSelection)
    onSelectionChange(nextSelection)
  } else {
    interactionState.dragAnchorCellRef.current = null
    interactionState.dragPointerCellRef.current = null
  }
  focusGrid()
}

export function handleGridPointerUp({
  event,
  dragAnchorCell,
  dragDidMove,
  dragHeaderSelection,
  dragPointerCell,
  dragViewport,
  dragGeometry,
  interactionState,
  lastBodyClickCellRef,
  onSelectionChange,
  postDragSelectionExpiryRef,
  resolvePointerCell,
  resolveHeaderSelectionForPointerDrag,
  selectedCell,
  setGridSelection,
  visibleRegion,
}: HandleGridPointerUpOptions): void {
  if (interactionState.columnResizeActiveRef.current) {
    return
  }
  if (dragHeaderSelection) {
    const finalHeader =
      resolveHeaderSelectionForPointerDrag(
        dragHeaderSelection.kind,
        event.clientX,
        event.clientY,
        dragViewport ?? visibleRegion,
        dragGeometry,
      ) ?? dragHeaderSelection
    const resolvedHeaderDrag = resolveHeaderDragSelection(dragHeaderSelection, finalHeader.index, selectedCell)
    setGridSelection(resolvedHeaderDrag.selection)
    onSelectionChange(resolvedHeaderDrag.selection)
    scheduleGridPointerInteractionReset(interactionState, {
      clearPostDragSelectionExpiry: false,
    })
    return
  }
  if (!dragAnchorCell) {
    return
  }
  if (dragDidMove) {
    const pointerCell =
      resolvePointerCell(event.clientX, event.clientY, dragViewport ?? visibleRegion, dragGeometry) ?? dragPointerCell ?? dragAnchorCell
    const pointerUpResult = resolveBodyPointerUpResult(dragAnchorCell, pointerCell, true)
    postDragSelectionExpiryRef.current = pointerUpResult.shouldSetDragExpiry ? window.performance.now() + 200 : 0
    if (pointerUpResult.selection) {
      setGridSelection(pointerUpResult.selection)
      onSelectionChange(pointerUpResult.selection)
    }
    lastBodyClickCellRef.current = pointerUpResult.clickedCell
  } else {
    const pointerUpResult = resolveBodyPointerUpResult(dragAnchorCell, dragPointerCell ?? dragAnchorCell, false)
    lastBodyClickCellRef.current = pointerUpResult.clickedCell
  }
  scheduleGridPointerInteractionReset(interactionState, {
    clearPostDragSelectionExpiry: false,
  })
}

export function startGridResize(interactionState: GridInteractionStateRefs): void {
  startGridColumnResize(interactionState)
}

export function finishGridResize(interactionState: GridInteractionStateRefs): void {
  finishGridColumnResize(interactionState)
}
