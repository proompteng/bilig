import React from "react";
import type { CellProps, SheetProps, WorkbookProps } from "./descriptors.js";

export function Workbook(props: WorkbookProps): React.ReactElement {
  const children = React.Children.map(props.children, (child) => {
    if (!React.isValidElement<{ name?: string }>(child)) {
      return child;
    }
    const key = child.props.name;
    return key ? React.cloneElement(child, { key }) : child;
  });
  return React.createElement("Workbook", { ...props, children });
}

export function Sheet(props: SheetProps): React.ReactElement {
  const children = React.Children.map(props.children, (child) => {
    if (!React.isValidElement<{ addr?: string }>(child)) {
      return child;
    }
    const key = child.props.addr;
    return key ? React.cloneElement(child, { key }) : child;
  });
  return React.createElement("Sheet", { ...props, children });
}

export function Cell(props: CellProps): React.ReactElement {
  return React.createElement("Cell", props);
}
