import React from "react";
import type { CellProps, SheetProps, WorkbookProps } from "./descriptors.js";

export function Workbook(props: WorkbookProps): React.ReactElement {
  const children = React.Children.toArray(props.children as React.ReactNode).map((child) => {
    if (!React.isValidElement<{ name?: string }>(child)) {
      return child;
    }
    const key = child.props.name;
    return key ? React.cloneElement(child, { key }) : child;
  });
  const { children: _children, ...workbookProps } = props;
  return React.createElement("Workbook", workbookProps, children);
}

export function Sheet(props: SheetProps): React.ReactElement {
  const children = React.Children.toArray(props.children as React.ReactNode).map((child) => {
    if (!React.isValidElement<{ addr?: string }>(child)) {
      return child;
    }
    const key = child.props.addr;
    return key ? React.cloneElement(child, { key }) : child;
  });
  const { children: _children, ...sheetProps } = props;
  return React.createElement("Sheet", sheetProps, children);
}

export function Cell(props: CellProps): React.ReactElement {
  return React.createElement("Cell", props);
}
