import { formatAddress } from "@bilig/formula";
import { createRectangleSelectionFromRange, rectangleToAddresses } from "./gridSelection.js";
import type { GridHoverState } from "./gridHover.js";
import { resolveMovedRange, sameRectangle } from "./gridRangeMove.js";
import type { Item, Rectangle } from "./gridTypes.js";

interface PointerEventLike {
  clientX: number;
  clientY: number;
}

interface PointerListenerTarget {
  addEventListener(
    type: "pointermove" | "pointerup",
    listener: (event: PointerEventLike) => void,
    useCapture: boolean,
  ): void;
  removeEventListener(
    type: "pointermove" | "pointerup",
    listener: (event: PointerEventLike) => void,
    useCapture: boolean,
  ): void;
}

export function beginWorkbookGridRangeMove(input: {
  cleanupRef: { current: (() => void) | null };
  listenerTarget: PointerListenerTarget;
  sourceRange: Rectangle;
  pointerCell: Item;
  resolvePointerCell: (clientX: number, clientY: number) => Item | null;
  setGridSelection: (selection: ReturnType<typeof createRectangleSelectionFromRange>) => void;
  onSelect: (address: string) => void;
  onMoveRange: (
    sourceStartAddress: string,
    sourceEndAddress: string,
    targetStartAddress: string,
    targetEndAddress: string,
  ) => void;
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void;
  setIsRangeMoveDragging: (isDragging: boolean) => void;
  setHoverState: (state: GridHoverState) => void;
}): void {
  const {
    cleanupRef,
    listenerTarget,
    onMoveRange,
    onSelect,
    pointerCell,
    refreshHoverState,
    resolvePointerCell,
    setGridSelection,
    setHoverState,
    setIsRangeMoveDragging,
    sourceRange,
  } = input;
  const anchorOffset: Item = [pointerCell[0] - sourceRange.x, pointerCell[1] - sourceRange.y];
  let previewRange = sourceRange;

  cleanupRef.current?.();
  setIsRangeMoveDragging(true);
  setHoverState({ cell: null, header: null, cursor: "grabbing" });

  const move = (nativeEvent: PointerEventLike) => {
    const nextPointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY);
    if (!nextPointerCell) {
      return;
    }
    const nextRange = resolveMovedRange(sourceRange, nextPointerCell, anchorOffset);
    if (sameRectangle(previewRange, nextRange)) {
      return;
    }
    previewRange = nextRange;
    setGridSelection(createRectangleSelectionFromRange(nextRange));
  };

  const cleanup = (nativeEvent?: PointerEventLike) => {
    listenerTarget.removeEventListener("pointermove", move, true);
    listenerTarget.removeEventListener("pointerup", up, true);
    cleanupRef.current = null;
    setIsRangeMoveDragging(false);
    if (nativeEvent) {
      refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0);
    }
  };

  const up = (nativeEvent: PointerEventLike) => {
    cleanup(nativeEvent);
    setGridSelection(createRectangleSelectionFromRange(previewRange));
    onSelect(formatAddress(previewRange.y, previewRange.x));
    if (sameRectangle(sourceRange, previewRange)) {
      return;
    }
    const sourceAddresses = rectangleToAddresses(sourceRange);
    const targetAddresses = rectangleToAddresses(previewRange);
    onMoveRange(
      sourceAddresses.startAddress,
      sourceAddresses.endAddress,
      targetAddresses.startAddress,
      targetAddresses.endAddress,
    );
  };

  cleanupRef.current = () => {
    cleanup();
  };
  listenerTarget.addEventListener("pointermove", move, true);
  listenerTarget.addEventListener("pointerup", up, true);
}
