import React from "react";
import { describe, expect, it } from "vitest";
import { Cell, Sheet, Workbook } from "../components.js";

describe("renderer components", () => {
  it("wraps workbook and sheet children with stable keys", () => {
    const workbookElement = Workbook({
      name: "book",
      children: [
        <Sheet name="Sheet1" key="ignored-1" />,
        <Sheet name="Sheet2" />
      ]
    });
    const workbookProps = workbookElement.props as { children?: React.ReactNode };

    const workbookChildren = React.Children.toArray(workbookProps.children);
    expect(workbookElement.type).toBe("Workbook");
    expect(String((workbookChildren[0] as React.ReactElement).key)).toContain("Sheet1");
    expect(String((workbookChildren[1] as React.ReactElement).key)).toContain("Sheet2");

    const sheetElement = Sheet({
      name: "Sheet1",
      children: [
        <Cell addr="A1" value={10} key="ignored-a1" />,
        <Cell addr="B1" formula="A1*2" />
      ]
    });
    const sheetProps = sheetElement.props as { children?: React.ReactNode };

    const sheetChildren = React.Children.toArray(sheetProps.children);
    expect(sheetElement.type).toBe("Sheet");
    expect(String((sheetChildren[0] as React.ReactElement).key)).toContain("A1");
    expect(String((sheetChildren[1] as React.ReactElement).key)).toContain("B1");
  });

  it("passes cell props through unchanged", () => {
    const element = Cell({ addr: "A1", value: 10, format: "currency-usd" });
    expect(element.type).toBe("Cell");
    expect(element.props).toMatchObject({
      addr: "A1",
      value: 10,
      format: "currency-usd"
    });
  });
});
