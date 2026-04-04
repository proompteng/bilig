import { describe, expect, test } from "vitest";
import { ErrorCode, ValueTag, type CellSnapshot, type CellStyleRecord } from "@bilig/protocol";
import {
  GridCellKind,
  cellStyleToThemeOverride,
  cellToEditorSeed,
  snapshotToGridCell,
  snapshotToRenderCell,
} from "../gridCells.js";

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

  test("derives renderer-native cell snapshots for text and autofit surfaces", () => {
    const renderCell = snapshotToRenderCell(
      makeSnapshot({
        formula: "A1&B1",
        value: { tag: ValueTag.String, value: "joined" },
      }),
      {
        id: "style-render",
        alignment: { horizontal: "center", wrap: true },
        font: { color: "#123456", size: 15, italic: true, underline: true },
      },
    );

    expect(renderCell).toMatchObject({
      kind: "string",
      displayText: "joined",
      copyText: "=A1&B1",
      align: "center",
      wrap: true,
      color: "#123456",
      fontSize: 15,
      underline: true,
    });
    expect(renderCell.font).toBe(
      'italic 400 15px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
    );

    expect(
      snapshotToRenderCell(makeSnapshot({ value: { tag: ValueTag.Boolean, value: false } })),
    ).toMatchObject({
      kind: "boolean",
      displayText: "FALSE",
      copyText: "FALSE",
    });
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

  test("makes text cells transparent when the text surface is active and routes booleans to the custom surface", () => {
    const textCell = snapshotToGridCell(
      makeSnapshot({ value: { tag: ValueTag.String, value: "hello" } }),
      undefined,
      { textSurfaceEnabled: true },
    );
    expect(textCell.kind).toBe(GridCellKind.Text);
    expect(textCell.themeOverride).toEqual({
      textDark: "rgba(32, 33, 36, 0)",
    });

    const booleanCell = snapshotToGridCell(
      makeSnapshot({ value: { tag: ValueTag.Boolean, value: true } }),
      undefined,
      { booleanSurfaceEnabled: true, textSurfaceEnabled: true },
    );
    expect(booleanCell.kind).toBe(GridCellKind.Text);
    expect(booleanCell.displayData).toBe("");
    expect(booleanCell.copyData).toBe("TRUE");
  });
});
