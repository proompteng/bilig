import React from "react";
import type { CellProps, SheetProps, WorkbookProps } from "./descriptors.js";

export function Workbook(props: WorkbookProps): React.ReactElement {
  return React.createElement("Workbook", props);
}

export function Sheet(props: SheetProps): React.ReactElement {
  return React.createElement("Sheet", props);
}

export function Cell(props: CellProps): React.ReactElement {
  return React.createElement("Cell", props);
}
