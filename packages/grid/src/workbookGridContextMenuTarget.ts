import type { HeaderSelection } from "./gridPointer.js";
import type { GridSelection, Rectangle } from "./gridTypes.js";

export interface WorkbookGridContextMenuTarget {
  readonly target: HeaderSelection;
  readonly x: number;
  readonly y: number;
}

export function resolveKeyboardHeaderContextMenuTarget(input: {
  gridSelection: GridSelection;
  targetCellBounds?: Rectangle | undefined;
  hostLeft: number;
  hostTop: number;
  rowMarkerWidth: number;
  headerHeight: number;
}): WorkbookGridContextMenuTarget | null {
  const { gridSelection, headerHeight, hostLeft, hostTop, rowMarkerWidth, targetCellBounds } =
    input;
  if (!targetCellBounds) {
    return null;
  }

  if (gridSelection.columns.length > 0 && gridSelection.rows.length === 0) {
    const columnIndex = gridSelection.columns.first();
    if (columnIndex === undefined) {
      return null;
    }
    return {
      target: { kind: "column", index: columnIndex },
      x: targetCellBounds.x + targetCellBounds.width / 2,
      y: hostTop + headerHeight / 2,
    };
  }

  if (gridSelection.rows.length > 0 && gridSelection.columns.length === 0) {
    const rowIndex = gridSelection.rows.first();
    if (rowIndex === undefined) {
      return null;
    }
    return {
      target: { kind: "row", index: rowIndex },
      x: hostLeft + rowMarkerWidth / 2,
      y: targetCellBounds.y + targetCellBounds.height / 2,
    };
  }

  return null;
}
