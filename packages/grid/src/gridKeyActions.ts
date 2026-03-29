import type { Item, Rectangle } from "@glideapps/glide-data-grid";
import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import { clampCell } from "./gridSelection.js";

export type GridEditSelectionBehavior = "select-all" | "caret-end";

export interface GridKeyActionEvent {
  key: string;
  ctrlKey: boolean;
  metaKey: boolean;
  altKey: boolean;
  shiftKey?: boolean;
}

export type GridKeyAction =
  | { kind: "none" }
  | { kind: "edit-append"; value: string }
  | { kind: "commit-edit"; movement?: readonly [-1 | 0 | 1, -1 | 0 | 1] }
  | { kind: "cancel-edit" }
  | {
      kind: "begin-edit";
      seed?: string;
      selectionBehavior: GridEditSelectionBehavior;
      pendingTypeSeed: string | null;
    }
  | { kind: "move-selection"; cell: Item }
  | { kind: "extend-selection"; anchor: Item; target: Item }
  | { kind: "clear-cell"; pendingTypeSeed: null }
  | { kind: "clipboard-copy" }
  | { kind: "clipboard-cut" }
  | { kind: "clipboard-paste"; target: Item }
  | { kind: "select-row"; row: number; col: number }
  | { kind: "select-column"; row: number; col: number }
  | { kind: "select-all" };

const PAGE_JUMP_ROWS = 20;

function moveSelectionToEdge(cell: Item, direction: "up" | "down" | "left" | "right"): Item {
  switch (direction) {
    case "up":
      return [cell[0], 0];
    case "down":
      return [cell[0], MAX_ROWS - 1];
    case "left":
      return [0, cell[1]];
    case "right":
      return [MAX_COLS - 1, cell[1]];
  }
}

function resolveSelectionActiveCell(
  anchorCell: Item,
  currentSelectionCell: Item | null,
  currentSelectionRange: Rectangle | null | undefined,
): Item {
  if (
    !currentSelectionRange ||
    (currentSelectionRange.width === 1 && currentSelectionRange.height === 1)
  ) {
    return currentSelectionCell ?? anchorCell;
  }

  const horizontalTarget =
    anchorCell[0] === currentSelectionRange.x
      ? currentSelectionRange.x + currentSelectionRange.width - 1
      : currentSelectionRange.x;
  const verticalTarget =
    anchorCell[1] === currentSelectionRange.y
      ? currentSelectionRange.y + currentSelectionRange.height - 1
      : currentSelectionRange.y;

  return [horizontalTarget, verticalTarget];
}

interface ResolveGridKeyActionOptions {
  event: GridKeyActionEvent;
  isEditingCell: boolean;
  editorValue: string;
  editorInputFocused: boolean;
  pendingTypeSeed: string | null;
  selectedCell: Item;
  currentSelectionCell: Item | null;
  currentRangeAnchor: Item | null;
  currentSelectionRange?: Rectangle | null;
}

export function resolveGridKeyAction(options: ResolveGridKeyActionOptions): GridKeyAction {
  const {
    event,
    isEditingCell,
    editorValue,
    editorInputFocused,
    pendingTypeSeed,
    selectedCell,
    currentSelectionCell,
    currentRangeAnchor,
    currentSelectionRange,
  } = options;

  const anchorCell = currentRangeAnchor ?? selectedCell;
  const activeCell = resolveSelectionActiveCell(
    anchorCell,
    currentSelectionCell,
    currentSelectionRange,
  );

  if (isEditingCell) {
    if (!editorInputFocused) {
      if (event.key === "Enter") {
        return { kind: "commit-edit", movement: [0, event.shiftKey ? -1 : 1] };
      }
      if (event.key === "Tab") {
        return { kind: "commit-edit", movement: [event.shiftKey ? -1 : 1, 0] };
      }
      if (event.key === "Escape") {
        return { kind: "cancel-edit" };
      }
      if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
        return { kind: "edit-append", value: `${editorValue}${event.key}` };
      }
    }
    return { kind: "none" };
  }

  if (event.key === "F2") {
    return { kind: "begin-edit", selectionBehavior: "caret-end", pendingTypeSeed: null };
  }

  const hasPrimaryModifier = event.ctrlKey || event.metaKey;
  const normalizedKey = event.key.toLowerCase();

  if (hasPrimaryModifier && !event.altKey && normalizedKey === "a") {
    return { kind: "select-all" };
  }

  if (!event.altKey && event.key === " " && hasPrimaryModifier && event.shiftKey) {
    return { kind: "select-all" };
  }

  if (!event.altKey && event.key === " " && hasPrimaryModifier) {
    return { kind: "select-column", col: activeCell[0], row: activeCell[1] };
  }

  if (!event.altKey && event.key === " " && event.shiftKey) {
    return { kind: "select-row", col: activeCell[0], row: activeCell[1] };
  }

  if (event.key === "Home") {
    const nextCell = hasPrimaryModifier ? ([0, 0] as Item) : ([0, activeCell[1]] as Item);
    if (event.shiftKey) {
      return {
        kind: "extend-selection",
        anchor: anchorCell,
        target: nextCell,
      };
    }
    return { kind: "move-selection", cell: nextCell };
  }

  if (event.key === "End") {
    const nextCell = hasPrimaryModifier
      ? ([MAX_COLS - 1, MAX_ROWS - 1] as Item)
      : ([MAX_COLS - 1, activeCell[1]] as Item);
    if (event.shiftKey) {
      return {
        kind: "extend-selection",
        anchor: anchorCell,
        target: nextCell,
      };
    }
    return { kind: "move-selection", cell: nextCell };
  }

  if (event.key === "PageUp" || event.key === "PageDown") {
    const nextCell = clampCell([
      activeCell[0],
      activeCell[1] + (event.key === "PageDown" ? PAGE_JUMP_ROWS : -PAGE_JUMP_ROWS),
    ]);
    if (event.shiftKey) {
      return {
        kind: "extend-selection",
        anchor: anchorCell,
        target: nextCell,
      };
    }
    return { kind: "move-selection", cell: nextCell };
  }

  if (event.key === "Enter") {
    return {
      kind: "move-selection",
      cell: clampCell([activeCell[0], activeCell[1] + (event.shiftKey ? -1 : 1)]),
    };
  }

  if (event.key === "Tab") {
    return {
      kind: "move-selection",
      cell: clampCell([activeCell[0] + (event.shiftKey ? -1 : 1), activeCell[1]]),
    };
  }

  if (
    event.key === "ArrowUp" ||
    event.key === "ArrowDown" ||
    event.key === "ArrowLeft" ||
    event.key === "ArrowRight"
  ) {
    const delta: Item =
      event.key === "ArrowUp"
        ? [0, -1]
        : event.key === "ArrowDown"
          ? [0, 1]
          : event.key === "ArrowLeft"
            ? [-1, 0]
            : [1, 0];
    const nextCell = hasPrimaryModifier
      ? moveSelectionToEdge(
          activeCell,
          event.key === "ArrowUp"
            ? "up"
            : event.key === "ArrowDown"
              ? "down"
              : event.key === "ArrowLeft"
                ? "left"
                : "right",
        )
      : clampCell([activeCell[0] + delta[0], activeCell[1] + delta[1]]);

    if (event.shiftKey) {
      return {
        kind: "extend-selection",
        anchor: anchorCell,
        target: nextCell,
      };
    }

    return { kind: "move-selection", cell: nextCell };
  }

  if (event.key === "Backspace" || event.key === "Delete") {
    return { kind: "clear-cell", pendingTypeSeed: null };
  }

  if (hasPrimaryModifier && !event.altKey) {
    if (normalizedKey === "c") {
      return { kind: "clipboard-copy" };
    }
    if (normalizedKey === "x") {
      return { kind: "clipboard-cut" };
    }
    if (normalizedKey === "v") {
      return { kind: "clipboard-paste", target: activeCell };
    }
  }

  if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const seed = `${pendingTypeSeed ?? ""}${event.key}`;
    return {
      kind: "begin-edit",
      seed,
      selectionBehavior: "caret-end",
      pendingTypeSeed: seed,
    };
  }

  return { kind: "none" };
}
