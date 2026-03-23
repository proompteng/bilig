import type { Item } from "@glideapps/glide-data-grid";
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
  | { kind: "clipboard-paste"; target: Item };

interface ResolveGridKeyActionOptions {
  event: GridKeyActionEvent;
  isEditingCell: boolean;
  editorValue: string;
  editorInputFocused: boolean;
  pendingTypeSeed: string | null;
  selectedCell: Item;
  currentSelectionCell: Item | null;
  currentRangeAnchor: Item | null;
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
  } = options;

  if (isEditingCell) {
    if (
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      !editorInputFocused
    ) {
      return { kind: "edit-append", value: `${editorValue}${event.key}` };
    }
    return { kind: "none" };
  }

  if (event.key === "F2") {
    return { kind: "begin-edit", selectionBehavior: "caret-end", pendingTypeSeed: null };
  }

  if (event.key === "Enter") {
    return {
      kind: "move-selection",
      cell: clampCell([selectedCell[0], selectedCell[1] + (event.shiftKey ? -1 : 1)]),
    };
  }

  if (event.key === "Tab") {
    return {
      kind: "move-selection",
      cell: clampCell([selectedCell[0] + (event.shiftKey ? -1 : 1), selectedCell[1]]),
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
    const nextCell = clampCell([selectedCell[0] + delta[0], selectedCell[1] + delta[1]]);

    if (event.shiftKey) {
      return {
        kind: "extend-selection",
        anchor: currentRangeAnchor ?? selectedCell,
        target: nextCell,
      };
    }

    return { kind: "move-selection", cell: nextCell };
  }

  if (event.key === "Backspace" || event.key === "Delete") {
    return { kind: "clear-cell", pendingTypeSeed: null };
  }

  if ((event.ctrlKey || event.metaKey) && !event.altKey) {
    const normalizedKey = event.key.toLowerCase();
    if (normalizedKey === "c") {
      return { kind: "clipboard-copy" };
    }
    if (normalizedKey === "x") {
      return { kind: "clipboard-cut" };
    }
    if (normalizedKey === "v") {
      return { kind: "clipboard-paste", target: currentSelectionCell ?? selectedCell };
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
