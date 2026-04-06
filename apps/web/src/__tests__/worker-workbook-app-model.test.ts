import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import {
  emptyCellSnapshot,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
} from "../worker-workbook-app-model.js";

describe("worker workbook app model", () => {
  it("normalizes snapshots into parsed editor input shapes", () => {
    const formulaCell = {
      ...emptyCellSnapshot("Sheet1", "A1"),
      formula: "SUM(B1:B3)",
      value: { tag: ValueTag.Number, value: 42 },
    };
    const booleanCell = {
      ...emptyCellSnapshot("Sheet1", "A2"),
      input: true,
      value: { tag: ValueTag.Boolean, value: true },
    };
    const errorCell = {
      ...emptyCellSnapshot("Sheet1", "A3"),
      value: { tag: ValueTag.Error, code: ErrorCode.Div0 },
    };

    expect(parsedEditorInputFromSnapshot(formulaCell)).toEqual({
      kind: "formula",
      formula: "SUM(B1:B3)",
    });
    expect(parsedEditorInputFromSnapshot(booleanCell)).toEqual({
      kind: "value",
      value: true,
    });
    expect(parsedEditorInputFromSnapshot(errorCell)).toEqual({
      kind: "value",
      value: "#DIV/0!",
    });
  });

  it("matches parsed editor input against authoritative snapshots", () => {
    const numericCell = {
      ...emptyCellSnapshot("Sheet1", "B4"),
      input: 17,
      value: { tag: ValueTag.Number, value: 17 },
    };
    const formulaCell = {
      ...emptyCellSnapshot("Sheet1", "B5"),
      formula: "A1+A2",
      value: { tag: ValueTag.Number, value: 9 },
    };

    expect(parsedEditorInputMatchesSnapshot({ kind: "value", value: 17 }, numericCell)).toBe(true);
    expect(
      parsedEditorInputMatchesSnapshot({ kind: "formula", formula: "A1+A2" }, formulaCell),
    ).toBe(true);
    expect(parsedEditorInputMatchesSnapshot({ kind: "clear" }, formulaCell)).toBe(false);
  });

  it("treats style or version-only drift as the same cell content", () => {
    const baseCell = {
      ...emptyCellSnapshot("Sheet1", "C7"),
      input: "remote",
      value: { tag: ValueTag.String, value: "remote" },
      version: 1,
    };
    const styleOnlyUpdate = {
      ...baseCell,
      styleId: "style-2",
      version: 2,
    };
    const contentChange = {
      ...baseCell,
      input: "local",
      value: { tag: ValueTag.String, value: "local" },
      version: 3,
    };

    expect(sameCellContent(baseCell, styleOnlyUpdate)).toBe(true);
    expect(sameCellContent(baseCell, contentChange)).toBe(false);
  });
});
