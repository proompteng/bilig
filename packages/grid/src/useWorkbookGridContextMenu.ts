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
import type { WorkbookGridContextMenuTarget } from "./workbookGridContextMenuTarget.js";

export function useWorkbookGridContextMenu(input: {
  focusGrid(this: void): void;
  hiddenColumnsByIndex?: Readonly<Record<number, true>> | undefined;
  hiddenRowsByIndex?: Readonly<Record<number, true>> | undefined;
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
    hiddenColumnsByIndex,
    hiddenRowsByIndex,
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

  const isTargetHidden = useCallback(
    (target: { kind: "row" | "column"; index: number }) =>
      target.kind === "row"
        ? hiddenRowsByIndex?.[target.index] === true
        : hiddenColumnsByIndex?.[target.index] === true,
    [hiddenColumnsByIndex, hiddenRowsByIndex],
  );

  const closeContextMenu = useCallback(() => {
    setContextMenuState(null);
    focusGrid();
  }, [focusGrid]);

  useEffect(() => {
    if (!contextMenuState) {
      return;
    }
    const firstMenuItem = menuRef.current?.querySelector<HTMLElement>('[role="menuitem"]');
    firstMenuItem?.focus();
  }, [contextMenuState]);

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
      closeContextMenu();
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        closeContextMenu();
      }
    };

    window.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("keydown", handleKeyDown, true);
    return () => {
      window.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [closeContextMenu, contextMenuState]);

  const toggleTargetHidden = useCallback(() => {
    if (!contextMenuState) {
      return;
    }
    if (contextMenuState.target.kind === "row") {
      onSetRowHidden?.(contextMenuState.target.index, !contextMenuState.target.hidden);
    } else {
      onSetColumnHidden?.(contextMenuState.target.index, !contextMenuState.target.hidden);
    }
    closeContextMenu();
  }, [closeContextMenu, contextMenuState, onSetColumnHidden, onSetRowHidden]);

  const openContextMenuForTarget = useCallback(
    ({ target, x, y }: WorkbookGridContextMenuTarget): boolean => {
      const canOpen =
        (target.kind === "row" && onSetRowHidden) ||
        (target.kind === "column" && onSetColumnHidden);
      if (!canOpen) {
        return false;
      }

      if (isEditingCell) {
        onCommitEdit();
      }
      focusGrid();
      if (target.kind === "row") {
        setGridSelection(createRowSliceSelection(selectedCell[0], target.index, target.index));
        onSelect(formatAddress(target.index, selectedCell[0]));
      } else {
        setGridSelection(createColumnSliceSelection(target.index, target.index, selectedCell[1]));
        onSelect(formatAddress(selectedCell[1], target.index));
      }
      setContextMenuState({
        x,
        y,
        target: {
          ...target,
          hidden: isTargetHidden(target),
        },
      });
      return true;
    },
    [
      focusGrid,
      isTargetHidden,
      isEditingCell,
      onCommitEdit,
      onSelect,
      onSetColumnHidden,
      onSetRowHidden,
      selectedCell,
      setGridSelection,
    ],
  );

  const handleHostContextMenuCapture = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const headerSelection = resolveHeaderSelectionAtPointer(
        event.clientX,
        event.clientY,
        visibleRegion,
      );
      if (!headerSelection) {
        closeContextMenu();
        return;
      }
      if (
        !openContextMenuForTarget({
          target: headerSelection,
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
    },
    [closeContextMenu, openContextMenuForTarget, resolveHeaderSelectionAtPointer, visibleRegion],
  );

  return useMemo(
    () => ({
      closeContextMenu,
      contextMenuState,
      handleHostContextMenuCapture,
      toggleTargetHidden,
      menuRef,
      openContextMenuForTarget,
    }),
    [
      closeContextMenu,
      contextMenuState,
      handleHostContextMenuCapture,
      toggleTargetHidden,
      openContextMenuForTarget,
    ],
  );
}
