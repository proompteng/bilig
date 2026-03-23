import React from "react";
import { describe, expect, it } from "vitest";
import { Cell, Sheet, Workbook } from "../components.js";

function isPropsWithChildren(value: unknown): value is { children?: React.ReactNode } {
  return typeof value === "object" && value !== null;
}

function expectElement(value: React.ReactNode): React.ReactElement {
  if (!React.isValidElement(value)) {
    throw new Error("Expected a React element");
  }
  return value;
}

describe("renderer components", () => {
  it("wraps workbook and sheet children with stable keys", () => {
    const workbookElement = Workbook({
      name: "book",
      children: [<Sheet name="Sheet1" key="ignored-1" />, <Sheet key="ignored-2" name="Sheet2" />],
    });
    expect(isPropsWithChildren(workbookElement.props)).toBe(true);
    if (!isPropsWithChildren(workbookElement.props)) {
      throw new Error("Workbook props should include children");
    }

    const workbookChildren = React.Children.toArray(workbookElement.props.children);
    expect(workbookElement.type).toBe("Workbook");
    expect(String(expectElement(workbookChildren[0]).key)).toContain("Sheet1");
    expect(String(expectElement(workbookChildren[1]).key)).toContain("Sheet2");

    const sheetElement = Sheet({
      name: "Sheet1",
      children: [
        <Cell addr="A1" value={10} key="ignored-a1" />,
        <Cell addr="B1" formula="A1*2" key="ignored-b1" />,
      ],
    });
    expect(isPropsWithChildren(sheetElement.props)).toBe(true);
    if (!isPropsWithChildren(sheetElement.props)) {
      throw new Error("Sheet props should include children");
    }

    const sheetChildren = React.Children.toArray(sheetElement.props.children);
    expect(sheetElement.type).toBe("Sheet");
    expect(String(expectElement(sheetChildren[0]).key)).toContain("A1");
    expect(String(expectElement(sheetChildren[1]).key)).toContain("B1");
  });

  it("passes cell props through unchanged", () => {
    const element = Cell({ addr: "A1", value: 10, format: "currency-usd" });
    expect(element.type).toBe("Cell");
    expect(element.props).toMatchObject({
      addr: "A1",
      value: 10,
      format: "currency-usd",
    });
  });
});
