import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
} from "react";
import { formatAddress } from "@bilig/formula";
import { createColumnSliceSelection, createRowSliceSelection } from "./gridSelection.js";
import type { HeaderSelection, VisibleRegionState } from "./gridPointer.js";
import type { Item, GridSelection } from "./gridTypes.js";
import type { WorkbookGridContextMenuState } from "./WorkbookGridContextMenu.js";

export function useWorkbookGridContextMenu(input: {
  focusGrid(this: void): void;
  isEditingCell: boolean;
  onCommitEdit(this: void): void;
  onSelect(this: void, addr: string): void;
  onSetColumnHidden?: ((columnIndex: number, hidden: boolean) => void) | undefined;
  onSetRowHidden?: ((rowIndex: number, hidden: boolean) => void) | undefined;
  resolveHeaderSelectionAtPointer(
    this: void,
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
  ): HeaderSelection | null;
  selectedCell: Item;
  setGridSelection(this: void, selection: GridSelection): void;
  visibleRegion: VisibleRegionState;
}) {
  const {
    focusGrid,
    isEditingCell,
    onCommitEdit,
    onSelect,
    onSetColumnHidden,
    onSetRowHidden,
    resolveHeaderSelectionAtPointer,
    selectedCell,
    setGridSelection,
    visibleRegion,
  } = input;
  const [contextMenuState, setContextMenuState] = useState<WorkbookGridContextMenuState | null>(
    null,
  );
  const menuRef = useRef<HTMLDivElement | null>(null);

  const closeContextMenu = useCallback(() => {
    setContextMenuState(null);
  }, []);

  useEffect(() => {
    if (!contextMenuState) {
      return;
    }

    const handlePointerDown = (event: PointerEvent) => {
      if (
        menuRef.current &&
        event.target instanceof Node &&
        menuRef.current.contains(event.target)
      ) {
        return;
      }
      setContextMenuState(null);
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        setContextMenuState(null);
      }
    };

    window.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("keydown", handleKeyDown, true);
    return () => {
      window.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [contextMenuState]);

  const hideTarget = useCallback(() => {
    if (!contextMenuState) {
      return;
    }
    if (contextMenuState.target.kind === "row") {
      onSetRowHidden?.(contextMenuState.target.index, true);
    } else {
      onSetColumnHidden?.(contextMenuState.target.index, true);
    }
    setContextMenuState(null);
  }, [contextMenuState, onSetColumnHidden, onSetRowHidden]);

  const handleHostContextMenuCapture = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const headerSelection = resolveHeaderSelectionAtPointer(
        event.clientX,
        event.clientY,
        visibleRegion,
      );
      if (!headerSelection) {
        setContextMenuState(null);
        return;
      }

      const canOpen =
        (headerSelection.kind === "row" && onSetRowHidden) ||
        (headerSelection.kind === "column" && onSetColumnHidden);
      if (!canOpen) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      if (isEditingCell) {
        onCommitEdit();
      }
      focusGrid();
      if (headerSelection.kind === "row") {
        setGridSelection(
          createRowSliceSelection(selectedCell[0], headerSelection.index, headerSelection.index),
        );
        onSelect(formatAddress(headerSelection.index, selectedCell[0]));
      } else {
        setGridSelection(
          createColumnSliceSelection(headerSelection.index, headerSelection.index, selectedCell[1]),
        );
        onSelect(formatAddress(selectedCell[1], headerSelection.index));
      }
      setContextMenuState({
        x: event.clientX,
        y: event.clientY,
        target: headerSelection,
      });
    },
    [
      focusGrid,
      isEditingCell,
      onCommitEdit,
      onSelect,
      onSetColumnHidden,
      onSetRowHidden,
      resolveHeaderSelectionAtPointer,
      selectedCell,
      setGridSelection,
      visibleRegion,
    ],
  );

  return useMemo(
    () => ({
      closeContextMenu,
      contextMenuState,
      handleHostContextMenuCapture,
      hideTarget,
      menuRef,
    }),
    [closeContextMenu, contextMenuState, handleHostContextMenuCapture, hideTarget],
  );
}
