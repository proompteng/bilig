import React from "react";
import type { CellProps, SheetProps, WorkbookProps } from "./descriptors.js";
import { RENDERER_KIND_PROP } from "./renderer-kind.js";

export function Workbook(props: WorkbookProps): React.ReactElement {
  const children = React.Children.toArray(props.children).map((child) => {
    if (!React.isValidElement<{ name?: string }>(child)) {
      return child;
    }
    const key = child.props.name;
    return key ? React.cloneElement(child, { key }) : child;
  });
  const { children: _children, ...workbookProps } = props;
  return React.createElement("Workbook", workbookProps, children);
}
Workbook[RENDERER_KIND_PROP] = "Workbook";

export function Sheet(props: SheetProps): React.ReactElement {
  const children = React.Children.toArray(props.children).map((child) => {
    if (!React.isValidElement<{ addr?: string }>(child)) {
      return child;
    }
    const key = child.props.addr;
    return key ? React.cloneElement(child, { key }) : child;
  });
  const { children: _children, ...sheetProps } = props;
  return React.createElement("Sheet", sheetProps, children);
}
Sheet[RENDERER_KIND_PROP] = "Sheet";

export function Cell(props: CellProps): React.ReactElement {
  return React.createElement("Cell", props);
}
Cell[RENDERER_KIND_PROP] = "Cell";
