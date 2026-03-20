import { describe, expect, test } from "vitest";
import { ErrorCode, ValueTag, type CellSnapshot } from "@bilig/protocol";
import { GridCellKind } from "@glideapps/glide-data-grid";
import { cellToEditorSeed, snapshotToGridCell } from "../gridCells.js";

    function makeSnapshot(overrides: Partial<CellSnapshot>): CellSnapshot {
      const snapshot: CellSnapshot = {
        address: "A1",
        input: null,
        formula: null,
        value: { tag: ValueTag.Empty },
        dependents: [],
        precedents: [],
        stale: false,
        ...overrides
      };
      return snapshot;
    }

describe("gridCells", () => {
  test("derives editor seeds from formulas, values, and inputs", () => {
    expect(cellToEditorSeed(makeSnapshot({ formula: "SUM(A1:A2)", value: { tag: ValueTag.Number, value: 3 } }))).toBe("=SUM(A1:A2)");
    expect(cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.Number, value: 42 } }))).toBe("42");
    expect(cellToEditorSeed(makeSnapshot({ input: true, value: { tag: ValueTag.Boolean, value: true } }))).toBe("TRUE");
    expect(cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.Error, code: ErrorCode.Value } }))).toBe("#VALUE!");
    expect(cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.String, value: "hello" } }))).toBe("hello");
  });

  test("converts snapshots into Glide cells", () => {
    const numberCell = snapshotToGridCell(makeSnapshot({ value: { tag: ValueTag.Number, value: 12 } }));
    expect(numberCell.kind).toBe(GridCellKind.Number);
    expect(numberCell.contentAlign).toBe("right");

    const booleanCell = snapshotToGridCell(makeSnapshot({ value: { tag: ValueTag.Boolean, value: false } }));
    expect(booleanCell.kind).toBe(GridCellKind.Boolean);

    const errorCell = snapshotToGridCell(makeSnapshot({ value: { tag: ValueTag.Error, code: ErrorCode.Ref } }));
    expect(errorCell.kind).toBe(GridCellKind.Text);
    expect(errorCell.displayData).toBe("#REF!");

    const formulaStringCell = snapshotToGridCell(
      makeSnapshot({
        formula: "A1&B1",
        value: { tag: ValueTag.String, value: "joined" }
      })
    );
    expect(formulaStringCell.kind).toBe(GridCellKind.Text);
    expect(formulaStringCell.copyData).toBe("=A1&B1");
  });
});
