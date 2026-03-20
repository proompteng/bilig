import { formatAddress } from "@bilig/formula";
import type { GridSelection, Item } from "@glideapps/glide-data-grid";
import { clampCell, clampSelectionRange, createRangeSelection } from "./gridSelection.js";

export function resolveActivatedCell(activatedCell: Item, dragAnchorCell: Item | null, pendingPointerCell: Item | null): Item {
  return dragAnchorCell ?? pendingPointerCell ?? clampCell(activatedCell);
}

interface ResolveSelectionChangeOptions {
  nextSelection: GridSelection;
  anchorCell: Item | null;
  pointerCell: Item | null;
  selectedCell: Item;
}

type SelectionChangeKind = "cell" | "column" | "row";

interface SelectionChangeResult {
  kind: SelectionChangeKind;
  selection: GridSelection;
  addr: string;
}

export function resolveSelectionChange({
  nextSelection,
  anchorCell,
  pointerCell,
  selectedCell
}: ResolveSelectionChangeOptions): SelectionChangeResult | null {
  if (nextSelection.columns.length > 0) {
    const nextColumn = nextSelection.columns.first();
    if (nextColumn !== undefined) {
      return {
        kind: "column",
        selection: nextSelection,
        addr: formatAddress(selectedCell[1], nextColumn)
      };
    }
  }

  if (nextSelection.rows.length > 0) {
    const nextRow = nextSelection.rows.first();
    if (nextRow !== undefined) {
      return {
        kind: "row",
        selection: nextSelection,
        addr: formatAddress(nextRow, selectedCell[0])
      };
    }
  }

  const nextCell = nextSelection.current?.cell;
  if (!nextCell) {
    return null;
  }

  const selection = anchorCell && pointerCell
    ? createRangeSelection(nextSelection, anchorCell, pointerCell)
    : {
        ...nextSelection,
        current: nextSelection.current
          ? {
              ...nextSelection.current,
              cell: clampCell(nextSelection.current.cell),
              range: clampSelectionRange(nextSelection.current.range)
            }
          : nextSelection.current
      };

  const cell = selection.current?.cell ?? nextCell;
  return {
    kind: "cell",
    selection,
    addr: formatAddress(cell[1], cell[0])
  };
}
