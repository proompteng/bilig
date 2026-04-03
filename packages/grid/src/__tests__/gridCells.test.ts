import { describe, expect, test } from "vitest";
import { ErrorCode, ValueTag, type CellSnapshot, type CellStyleRecord } from "@bilig/protocol";
import { GridCellKind } from "@glideapps/glide-data-grid";
import { cellStyleToThemeOverride, cellToEditorSeed, snapshotToGridCell } from "../gridCells.js";

function makeSnapshot(overrides: Partial<CellSnapshot>): CellSnapshot {
  const snapshot: CellSnapshot = {
    address: "A1",
    input: null,
    formula: null,
    value: { tag: ValueTag.Empty },
    dependents: [],
    precedents: [],
    stale: false,
    ...overrides,
  };
  return snapshot;
}

describe("gridCells", () => {
  test("derives editor seeds from formulas, values, and inputs", () => {
    expect(
      cellToEditorSeed(
        makeSnapshot({ formula: "SUM(A1:A2)", value: { tag: ValueTag.Number, value: 3 } }),
      ),
    ).toBe("=SUM(A1:A2)");
    expect(cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.Number, value: 42 } }))).toBe(
      "42",
    );
    expect(
      cellToEditorSeed(
        makeSnapshot({ input: true, value: { tag: ValueTag.Boolean, value: true } }),
      ),
    ).toBe("TRUE");
    expect(
      cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.Error, code: ErrorCode.Value } })),
    ).toBe("#VALUE!");
    expect(
      cellToEditorSeed(makeSnapshot({ value: { tag: ValueTag.String, value: "hello" } })),
    ).toBe("hello");
  });

  test("converts snapshots into Glide cells", () => {
    const numberCell = snapshotToGridCell(
      makeSnapshot({ value: { tag: ValueTag.Number, value: 12 } }),
    );
    expect(numberCell.kind).toBe(GridCellKind.Number);
    expect(numberCell.contentAlign).toBe("right");

    const booleanCell = snapshotToGridCell(
      makeSnapshot({ value: { tag: ValueTag.Boolean, value: false } }),
    );
    expect(booleanCell.kind).toBe(GridCellKind.Boolean);

    const errorCell = snapshotToGridCell(
      makeSnapshot({ value: { tag: ValueTag.Error, code: ErrorCode.Ref } }),
    );
    expect(errorCell.kind).toBe(GridCellKind.Text);
    expect(errorCell.displayData).toBe("#REF!");

    const formulaStringCell = snapshotToGridCell(
      makeSnapshot({
        formula: "A1&B1",
        value: { tag: ValueTag.String, value: "joined" },
      }),
    );
    expect(formulaStringCell.kind).toBe(GridCellKind.Text);
    expect(formulaStringCell.copyData).toBe("=A1&B1");
  });

  test("keeps fill styling out of theme overrides so grid borders stay stable", () => {
    const fillOnlyStyle: CellStyleRecord = {
      id: "style-fill",
      fill: { backgroundColor: "#ea9999" },
    };

    expect(cellStyleToThemeOverride(fillOnlyStyle)).toBeUndefined();

    const fontAndFillStyle: CellStyleRecord = {
      id: "style-font-fill",
      fill: { backgroundColor: "#ea9999" },
      font: { color: "#202124", family: "Roboto", size: 12 },
    };

    expect(cellStyleToThemeOverride(fontAndFillStyle)).toEqual({
      textDark: "#202124",
      baseFontStyle: "400 12px",
      fontFamily: '"JetBrainsMono Nerd Font","JetBrains Mono",monospace',
    });
  });
});
