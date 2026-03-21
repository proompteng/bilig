import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { materializePivotTable } from "../pivot-engine.js";

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

describe("materializePivotTable", () => {
  it("groups rows by key columns and accumulates sum and count values", () => {
    const result = materializePivotTable(
      {
        groupBy: ["Region", "Product"],
        values: [
          { sourceColumn: "Sales", summarizeBy: "sum" },
          { sourceColumn: "Sales", summarizeBy: "count", outputLabel: "Rows" }
        ]
      },
      [
        [stringValue("Region"), stringValue("Product"), stringValue("Sales")],
        [stringValue("East"), stringValue("Widget"), numberValue(10)],
        [stringValue("East"), stringValue("Widget"), numberValue(5)],
        [stringValue("West"), stringValue("Gizmo"), numberValue(7)]
      ]
    );

    expect(result).toEqual({
      kind: "ok",
      rows: 3,
      cols: 4,
      values: [
        stringValue("Region"),
        stringValue("Product"),
        stringValue("SUM of Sales"),
        stringValue("Rows"),
        stringValue("East"),
        stringValue("Widget"),
        numberValue(15),
        numberValue(2),
        stringValue("West"),
        stringValue("Gizmo"),
        numberValue(7),
        numberValue(1)
      ]
    });
  });

  it("returns a #VALUE pivot result when configured columns are missing", () => {
    const result = materializePivotTable(
      {
        groupBy: ["Region"],
        values: [{ sourceColumn: "Sales", summarizeBy: "sum" }]
      },
      [[stringValue("Category"), stringValue("Amount")]]
    );

    expect(result).toEqual({
      kind: "error",
      code: ErrorCode.Value,
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Error, code: ErrorCode.Value }]
    });
  });
});
