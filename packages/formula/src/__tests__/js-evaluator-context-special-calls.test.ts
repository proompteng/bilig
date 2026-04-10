import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { evaluatePlan, lowerToPlan, parseFormula } from "../index.js";

const context = {
  sheetName: "Sheet2",
  currentAddress: "C4",
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case "A1":
        return number(2);
      case "B1":
        return number(3);
      default:
        return empty();
    }
  },
  resolveRange: (_sheetName: string, _start: string, _end: string): CellValue[] => [],
  listSheetNames: () => ["Sheet1", "Sheet2", "Summary"],
  resolveFormula: (sheetName: string, address: string): string | undefined =>
    sheetName === "Sheet2" && address === "B1"
      ? "SUM(A1:A2)"
      : sheetName === "Sheet2" && address === "C1"
        ? "A1*2"
        : undefined,
};

describe("js evaluator context special calls", () => {
  it("evaluates metadata and context helpers", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("ROW()")), context)).toEqual(number(4));
    expect(evaluatePlan(lowerToPlan(parseFormula("COLUMN(B:D)")), context)).toEqual(number(2));
    expect(evaluatePlan(lowerToPlan(parseFormula("FORMULATEXT(Sheet2!B1)")), context)).toEqual(
      text("=SUM(A1:A2)"),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula("FORMULA(Sheet2!C1)")), context)).toEqual(
      text("=A1*2"),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula("PHONETIC(42)")), context)).toEqual(text("42"));
    expect(evaluatePlan(lowerToPlan(parseFormula('CHOOSE(2,"a","b")')), context)).toEqual(
      text("b"),
    );
    expect(
      evaluatePlan(lowerToPlan(parseFormula("LAMBDA(x,IF(ISOMITTED(x),1,0))()")), context),
    ).toEqual(number(1));
  });

  it("evaluates sheet and cell info helpers", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("SHEET()")), context)).toEqual(number(2));
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Summary")')), context)).toEqual(
      number(3),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula("SHEETS()")), context)).toEqual(number(3));
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("address",B1)')), context)).toEqual(
      text("$B$1"),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("row",A1)')), context)).toEqual(number(1));
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("col",B1)')), context)).toEqual(number(2));
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("type",B1)')), context)).toEqual(text("v"));
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("filename")')), context)).toEqual(text(""));
  });

  it("preserves validation and NA or REF branches", () => {
    expect(evaluatePlan(lowerToPlan(parseFormula("FORMULATEXT(1)")), context)).toEqual(
      err(ErrorCode.Ref),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Missing")')), context)).toEqual(
      err(ErrorCode.NA),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("bogus",A1)')), context)).toEqual(
      err(ErrorCode.Value),
    );
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("type")')), { ...context, currentAddress: undefined }),
    ).toEqual(err(ErrorCode.Value));
    expect(evaluatePlan(lowerToPlan(parseFormula("CHOOSE(0,1,2)")), context)).toEqual(
      err(ErrorCode.Value),
    );
    expect(evaluatePlan(lowerToPlan(parseFormula("PHONETIC()")), context)).toEqual(
      err(ErrorCode.Value),
    );
  });
});

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function empty(): CellValue {
  return { tag: ValueTag.Empty };
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}
