import type { Item, Rectangle } from "@glideapps/glide-data-grid";
import type { HeaderSelection, PointerGeometry, VisibleRegionState } from "./gridPointer.js";
import {
  resolveColumnResizeTarget,
  resolveHeaderSelection,
  resolvePointerCell,
} from "./gridPointer.js";
import type { GridMetrics } from "./gridMetrics.js";

export type GridHoverCursor = "default" | "cell" | "pointer" | "col-resize";

export interface GridHoverState {
  readonly cell: Item | null;
  readonly header: HeaderSelection | null;
  readonly cursor: GridHoverCursor;
}

interface ResolveGridHoverStateOptions {
  readonly clientX: number;
  readonly clientY: number;
  readonly region: VisibleRegionState;
  readonly geometry: PointerGeometry;
  readonly columnWidths: Readonly<Record<number, number>>;
  readonly defaultColumnWidth: number;
  readonly gridMetrics: GridMetrics;
  readonly selectedCell: Item;
  readonly selectedCellBounds?: Rectangle | null;
  readonly selectionRange?: Rectangle | null;
  readonly hasColumnSelection: boolean;
  readonly hasRowSelection: boolean;
}

export function resolveGridHoverState(options: ResolveGridHoverStateOptions): GridHoverState {
  const {
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    defaultColumnWidth,
    gridMetrics,
    selectedCell,
    selectedCellBounds,
    selectionRange,
    hasColumnSelection,
    hasRowSelection,
  } = options;

  const resizeTarget = resolveColumnResizeTarget(
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    defaultColumnWidth,
  );
  if (resizeTarget !== null) {
    return {
      cell: null,
      header: { kind: "column", index: resizeTarget },
      cursor: "col-resize",
    };
  }

  const header = resolveHeaderSelection(
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    gridMetrics,
  );
  if (header) {
    return {
      cell: null,
      header,
      cursor: "pointer",
    };
  }

  const cell = resolvePointerCell({
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    gridMetrics,
    selectedCell,
    ...(selectedCellBounds ? { selectedCellBounds } : {}),
    selectionRange: selectionRange ?? null,
    hasColumnSelection,
    hasRowSelection,
  });
  if (cell) {
    return {
      cell,
      header: null,
      cursor: "cell",
    };
  }

  return {
    cell: null,
    header: null,
    cursor: "default",
  };
}

export function sameGridHoverState(left: GridHoverState, right: GridHoverState): boolean {
  return (
    sameItem(left.cell, right.cell) &&
    sameHeaderSelection(left.header, right.header) &&
    left.cursor === right.cursor
  );
}

function sameItem(left: Item | null, right: Item | null): boolean {
  return (
    left === right ||
    (left !== null && right !== null && left[0] === right[0] && left[1] === right[1])
  );
}

function sameHeaderSelection(left: HeaderSelection | null, right: HeaderSelection | null): boolean {
  return (
    left === right ||
    (left !== null && right !== null && left.kind === right.kind && left.index === right.index)
  );
}
