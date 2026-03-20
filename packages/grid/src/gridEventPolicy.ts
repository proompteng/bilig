import { formatAddress } from "@bilig/formula";
import type { Item } from "@glideapps/glide-data-grid";
import { sameItem } from "./gridSelection.js";

type GridVariant = "playground" | "product";

export function resolveBodyDoubleClickIntent(options: {
  variant: GridVariant;
  resizeTarget: number | null;
  bodyCell: Item | null;
  lastBodyClickCell: Item | null;
}):
  | { kind: "ignore" }
  | { kind: "autofit-column"; columnIndex: number }
  | { kind: "edit-cell"; cell: Item } {
  const { variant, resizeTarget, bodyCell, lastBodyClickCell } = options;
  if (variant !== "product") {
    return { kind: "ignore" };
  }
  if (resizeTarget !== null) {
    return { kind: "autofit-column", columnIndex: resizeTarget };
  }
  if (bodyCell === null || !sameItem(bodyCell, lastBodyClickCell)) {
    return { kind: "ignore" };
  }
  return { kind: "edit-cell", cell: bodyCell };
}

export function resolveHeaderClickIntent(options: {
  variant: GridVariant;
  isEdge: boolean;
  isDoubleClick: boolean;
  columnResizeActive: boolean;
  columnIndex: number;
  selectedRow: number;
}):
  | { kind: "ignore" }
  | { kind: "autofit-column"; columnIndex: number }
  | { kind: "select-column"; columnIndex: number; selectedRow: number; addr: string } {
  const { variant, isEdge, isDoubleClick, columnResizeActive, columnIndex, selectedRow } = options;
  if (variant === "product" && isEdge && isDoubleClick) {
    return { kind: "autofit-column", columnIndex };
  }
  if (columnResizeActive) {
    return { kind: "ignore" };
  }
  return {
    kind: "select-column",
    columnIndex,
    selectedRow,
    addr: formatAddress(selectedRow, columnIndex)
  };
}

export function shouldSkipGridSelectionChange(options: {
  columnResizeActive: boolean;
  postDragSelectionExpiry: number;
  now: number;
  ignoreNextPointerSelection: boolean;
  hasDragViewport: boolean;
}): {
  skip: boolean;
  clearPostDragSelectionExpiry: boolean;
  consumeIgnoreNextPointerSelection: boolean;
} {
  const {
    columnResizeActive,
    postDragSelectionExpiry,
    now,
    ignoreNextPointerSelection,
    hasDragViewport
  } = options;

  if (columnResizeActive) {
    return { skip: true, clearPostDragSelectionExpiry: false, consumeIgnoreNextPointerSelection: false };
  }
  if (postDragSelectionExpiry > 0 && now <= postDragSelectionExpiry) {
    return { skip: true, clearPostDragSelectionExpiry: true, consumeIgnoreNextPointerSelection: false };
  }
  if (ignoreNextPointerSelection) {
    return { skip: true, clearPostDragSelectionExpiry: false, consumeIgnoreNextPointerSelection: true };
  }
  if (hasDragViewport) {
    return { skip: true, clearPostDragSelectionExpiry: false, consumeIgnoreNextPointerSelection: false };
  }
  return { skip: false, clearPostDragSelectionExpiry: false, consumeIgnoreNextPointerSelection: false };
}
