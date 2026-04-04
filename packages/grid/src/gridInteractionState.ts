import type { MutableRefObject } from "react";
import type { Item } from "./gridTypes.js";
import type { HeaderSelection, PointerGeometry, VisibleRegionState } from "./gridPointer.js";

export interface GridInteractionStateRefs {
  ignoreNextPointerSelectionRef: MutableRefObject<boolean>;
  pendingPointerCellRef: MutableRefObject<Item | null>;
  dragAnchorCellRef: MutableRefObject<Item | null>;
  dragPointerCellRef: MutableRefObject<Item | null>;
  dragHeaderSelectionRef: MutableRefObject<HeaderSelection | null>;
  dragViewportRef: MutableRefObject<VisibleRegionState | null>;
  dragGeometryRef: MutableRefObject<PointerGeometry | null>;
  dragDidMoveRef: MutableRefObject<boolean>;
  postDragSelectionExpiryRef: MutableRefObject<number>;
  columnResizeActiveRef: MutableRefObject<boolean>;
}

interface ResetGridPointerInteractionOptions {
  clearIgnoreNextPointerSelection?: boolean;
  clearPostDragSelectionExpiry?: boolean;
}

export function resetGridPointerInteraction(
  refs: GridInteractionStateRefs,
  options: ResetGridPointerInteractionOptions = {},
): void {
  if (options.clearIgnoreNextPointerSelection) {
    refs.ignoreNextPointerSelectionRef.current = false;
  }
  refs.pendingPointerCellRef.current = null;
  refs.dragAnchorCellRef.current = null;
  refs.dragPointerCellRef.current = null;
  refs.dragHeaderSelectionRef.current = null;
  refs.dragGeometryRef.current = null;
  refs.dragDidMoveRef.current = false;
  refs.dragViewportRef.current = null;
  if (options.clearPostDragSelectionExpiry ?? true) {
    refs.postDragSelectionExpiryRef.current = 0;
  }
}

export function scheduleGridPointerInteractionReset(
  refs: GridInteractionStateRefs,
  options: ResetGridPointerInteractionOptions = {},
): void {
  window.requestAnimationFrame(() => {
    resetGridPointerInteraction(refs, options);
  });
}

export function startGridColumnResize(refs: GridInteractionStateRefs): void {
  refs.columnResizeActiveRef.current = true;
  resetGridPointerInteraction(refs);
}

export function finishGridColumnResize(refs: GridInteractionStateRefs): void {
  window.requestAnimationFrame(() => {
    refs.columnResizeActiveRef.current = false;
  });
}

export function beginGridHeaderDrag(
  refs: GridInteractionStateRefs,
  headerSelection: HeaderSelection,
  geometry: PointerGeometry | null,
  visibleRegion: VisibleRegionState,
): void {
  resetGridPointerInteraction(refs);
  refs.dragHeaderSelectionRef.current = headerSelection;
  refs.dragGeometryRef.current = geometry;
  refs.dragViewportRef.current = visibleRegion;
}

export function beginGridBodyPointerInteraction(
  refs: GridInteractionStateRefs,
  pointerCell: Item | null,
  geometry: PointerGeometry | null,
  visibleRegion: VisibleRegionState,
): void {
  refs.ignoreNextPointerSelectionRef.current = pointerCell === null;
  refs.pendingPointerCellRef.current = pointerCell;
  refs.dragGeometryRef.current = geometry;
  refs.dragDidMoveRef.current = false;
  refs.dragViewportRef.current = visibleRegion;
  refs.postDragSelectionExpiryRef.current = 0;
}

export function clearGridPendingPointerActivation(refs: GridInteractionStateRefs): void {
  refs.pendingPointerCellRef.current = null;
  refs.dragAnchorCellRef.current = null;
  refs.dragPointerCellRef.current = null;
}
