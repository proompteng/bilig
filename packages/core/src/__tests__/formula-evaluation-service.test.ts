import { Effect } from "effect";
import { describe, expect, it, vi } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { FormulaNode } from "@bilig/formula";
import { SpreadsheetEngine } from "../engine.js";
import { EngineFormulaEvaluationError } from "../engine/errors.js";
import type { EngineFormulaEvaluationService } from "../engine/services/formula-evaluation-service.js";
import type { EngineMutationSupportService } from "../engine/services/mutation-support-service.js";

function isEngineFormulaEvaluationService(value: unknown): value is EngineFormulaEvaluationService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "evaluateUnsupportedFormula") === "function" &&
    typeof Reflect.get(value, "resolveStructuredReference") === "function" &&
    typeof Reflect.get(value, "resolveSpillReference") === "function"
  );
}

function isEngineMutationSupportService(value: unknown): value is EngineMutationSupportService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return typeof Reflect.get(value, "clearOwnedSpill") === "function";
}

function getEvaluationService(engine: SpreadsheetEngine): EngineFormulaEvaluationService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const evaluation = Reflect.get(runtime, "evaluation");
  if (!isEngineFormulaEvaluationService(evaluation)) {
    throw new TypeError("Expected engine formula evaluation service");
  }
  return evaluation;
}

function getMutationSupportService(engine: SpreadsheetEngine): EngineMutationSupportService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const support = Reflect.get(runtime, "support");
  if (!isEngineMutationSupportService(support)) {
    throw new TypeError("Expected engine mutation support service");
  }
  return support;
}

describe("EngineFormulaEvaluationService", () => {
  it("re-evaluates JS indirection spills through the service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-indirect" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "B1", 10);
    engine.setCellValue("Sheet1", "B2", 20);
    engine.setCellFormula("Sheet1", "G1", 'INDIRECT("B1:B2")');

    const g1Index = engine.workbook.getCellIndex("Sheet1", "G1");
    expect(g1Index).toBeDefined();

    Effect.runSync(getMutationSupportService(engine).clearOwnedSpill(g1Index!));

    expect(engine.getCellValue("Sheet1", "G2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined();

    Effect.runSync(getEvaluationService(engine).evaluateUnsupportedFormula(g1Index!));

    expect(engine.getCellValue("Sheet1", "G1")).toEqual({ tag: ValueTag.Number, value: 10 });
    expect(engine.getCellValue("Sheet1", "G2")).toEqual({ tag: ValueTag.Number, value: 20 });
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([
      { sheetName: "Sheet1", address: "G1", rows: 2, cols: 1 },
    ]);
    expect(
      Effect.runSync(getEvaluationService(engine).resolveSpillReference("Sheet1", undefined, "G1")),
    ).toEqual({
      kind: "RangeRef",
      refKind: "cells",
      sheetName: "Sheet1",
      start: "G1",
      end: "G2",
    } satisfies FormulaNode);
  });

  it("resolves structured references to table body rows through the service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-structured-ref" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setTable({
      name: "Sales",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Amount", "Total"],
      headerRow: true,
      totalsRow: false,
    });

    const resolved = Effect.runSync(
      getEvaluationService(engine).resolveStructuredReference("Sales", "Amount"),
    );

    expect(resolved).toEqual({
      kind: "RangeRef",
      refKind: "cells",
      sheetName: "Sheet1",
      start: "A2",
      end: "A3",
    } satisfies FormulaNode);

    expect(
      Effect.runSync(getEvaluationService(engine).resolveStructuredReference("Missing", "Amount")),
    ).toBeUndefined();
    expect(
      Effect.runSync(getEvaluationService(engine).resolveStructuredReference("Sales", "Missing")),
    ).toBeUndefined();

    engine.setTable({
      name: "HeaderOnly",
      sheetName: "Sheet1",
      startAddress: "D1",
      endAddress: "D1",
      columnNames: ["Amount"],
      headerRow: true,
      totalsRow: false,
    });
    expect(
      Effect.runSync(
        getEvaluationService(engine).resolveStructuredReference("HeaderOnly", "Amount"),
      ),
    ).toEqual({
      kind: "ErrorLiteral",
      code: ErrorCode.Ref,
    } satisfies FormulaNode);

    expect(
      Effect.runSync(getEvaluationService(engine).resolveSpillReference("Sheet1", undefined, "Z1")),
    ).toBeUndefined();
  });

  it("resolves MULTIPLE.OPERATIONS through reference replacements and missing formula cells", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-multiple-operations" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "B1", 2);
    engine.setCellValue("Sheet1", "A2", 10);
    engine.setCellValue("Sheet1", "B2", 20);
    engine.setCellFormula("Sheet1", "C1", "A1+B1");

    const evaluation = getEvaluationService(engine);
    expect(
      Effect.runSync(
        evaluation.resolveMultipleOperations({
          formulaSheetName: "Sheet1",
          formulaAddress: "C1",
          rowCellSheetName: "Sheet1",
          rowCellAddress: "A1",
          rowReplacementSheetName: "Sheet1",
          rowReplacementAddress: "A2",
          columnCellSheetName: "Sheet1",
          columnCellAddress: "B1",
          columnReplacementSheetName: "Sheet1",
          columnReplacementAddress: "B2",
        }),
      ),
    ).toEqual({ tag: ValueTag.Number, value: 30 });

    expect(
      Effect.runSync(
        evaluation.resolveMultipleOperations({
          formulaSheetName: "Sheet1",
          formulaAddress: "Z99",
          rowCellSheetName: "Sheet1",
          rowCellAddress: "A1",
          rowReplacementSheetName: "Sheet1",
          rowReplacementAddress: "A2",
        }),
      ),
    ).toEqual({ tag: ValueTag.Empty });
  });

  it("returns empty results for non-formula cells and evaluates literal MATCH through the lookup resolver", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "evaluation-lookup-resolver",
      useColumnIndex: true,
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", "apple");
    engine.setCellValue("Sheet1", "A2", "pear");
    engine.setCellValue("Sheet1", "A3", "plum");
    engine.setCellFormula("Sheet1", "B1", 'MATCH("pear",A1:A3,0)');

    const evaluation = getEvaluationService(engine);
    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    expect(a1Index).toBeDefined();
    expect(b1Index).toBeDefined();

    expect(Effect.runSync(evaluation.evaluateDirectLookupFormula(a1Index!))).toBeUndefined();
    expect(Effect.runSync(evaluation.evaluateUnsupportedFormula(a1Index!))).toEqual([]);

    Effect.runSync(evaluation.evaluateUnsupportedFormula(b1Index!));
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 2 });
  });

  it("wraps workbook access failures from structured, spill, and multiple-operations helpers", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-error-wrappers" });
    await engine.ready();
    engine.createSheet("Sheet1");

    const evaluation = getEvaluationService(engine);

    const getTableSpy = vi.spyOn(engine.workbook, "getTable").mockImplementation(() => {
      throw new Error("structured explode");
    });
    const structured = Effect.runSync(
      Effect.either(evaluation.resolveStructuredReference("Sales", "Amount")),
    );
    expect(structured._tag).toBe("Left");
    expect(structured.left).toBeInstanceOf(EngineFormulaEvaluationError);
    expect(structured.left.message).toContain("structured explode");
    getTableSpy.mockRestore();

    const getSpillSpy = vi.spyOn(engine.workbook, "getSpill").mockImplementation(() => {
      throw new Error("spill explode");
    });
    const spill = Effect.runSync(
      Effect.either(evaluation.resolveSpillReference("Sheet1", undefined, "A1")),
    );
    expect(spill._tag).toBe("Left");
    expect(spill.left).toBeInstanceOf(EngineFormulaEvaluationError);
    expect(spill.left.message).toContain("spill explode");
    getSpillSpy.mockRestore();

    const getCellIndexSpy = vi.spyOn(engine.workbook, "getCellIndex").mockImplementation(() => {
      throw new Error("multiple operations explode");
    });
    const multipleOperations = Effect.runSync(
      Effect.either(
        evaluation.resolveMultipleOperations({
          formulaSheetName: "Sheet1",
          formulaAddress: "A1",
          rowCellSheetName: "Sheet1",
          rowCellAddress: "A1",
          rowReplacementSheetName: "Sheet1",
          rowReplacementAddress: "A2",
        }),
      ),
    );
    expect(multipleOperations._tag).toBe("Left");
    expect(multipleOperations.left).toBeInstanceOf(EngineFormulaEvaluationError);
    expect(multipleOperations.left.message).toContain("multiple operations explode");
    getCellIndexSpy.mockRestore();
  });

  it("evaluates direct exact lookup formulas across uniform, text, and mixed columns", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "evaluation-direct-exact-service",
      useColumnIndex: true,
    });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "A2", 2);
    engine.setCellValue("Sheet1", "A3", 3);
    engine.setCellValue("Sheet1", "B1", 3);
    engine.setCellValue("Sheet1", "B2", 2);
    engine.setCellValue("Sheet1", "B3", 1);
    engine.setCellValue("Sheet1", "C1", "pear");
    engine.setCellValue("Sheet1", "C2", "apple");
    engine.setCellValue("Sheet1", "C3", "pear");
    engine.setCellValue("Sheet1", "E2", "pear");
    engine.setCellValue("Sheet1", "E3", false);
    engine.setCellValue("Sheet1", "D1", 2.5);
    engine.setCellValue("Sheet1", "D2", 2);
    engine.setCellValue("Sheet1", "D3", false);
    engine.setCellValue("Sheet1", "D4", false);

    engine.setCellFormula("Sheet1", "F1", "MATCH(D1,A1:A3,0)");
    engine.setCellFormula("Sheet1", "F2", "MATCH(D2,B1:B3,0)");
    engine.setCellFormula("Sheet1", "F3", "MATCH(D3,C1:C3,0)");
    engine.setCellFormula("Sheet1", "F4", "MATCH(D4,E1:E3,0)");

    const evaluation = getEvaluationService(engine);
    const f1Index = engine.workbook.getCellIndex("Sheet1", "F1");
    const f2Index = engine.workbook.getCellIndex("Sheet1", "F2");
    const f3Index = engine.workbook.getCellIndex("Sheet1", "F3");
    const f4Index = engine.workbook.getCellIndex("Sheet1", "F4");
    expect(f1Index).toBeDefined();
    expect(f2Index).toBeDefined();
    expect(f3Index).toBeDefined();
    expect(f4Index).toBeDefined();

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!));
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    engine.setCellValue("Sheet1", "D1", 2);
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!));
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setRangeValues({ sheetName: "Sheet1", startAddress: "A1", endAddress: "A3" }, [
      [4],
      [5],
      [6],
    ]);
    engine.setCellValue("Sheet1", "D1", 5);
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!));
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({ tag: ValueTag.Number, value: 2 });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!));
    expect(engine.getCellValue("Sheet1", "F2")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setRangeValues({ sheetName: "Sheet1", startAddress: "B1", endAddress: "B3" }, [
      [6],
      [5],
      [4],
    ]);
    engine.setCellValue("Sheet1", "D2", 5);
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!));
    expect(engine.getCellValue("Sheet1", "F2")).toEqual({ tag: ValueTag.Number, value: 2 });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!));
    expect(engine.getCellValue("Sheet1", "F3")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f4Index!));
    expect(engine.getCellValue("Sheet1", "F4")).toEqual({ tag: ValueTag.Number, value: 3 });
  });

  it("evaluates direct approximate lookup formulas across uniform, refreshed, and text columns", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "evaluation-direct-approx-service" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "A2", 2);
    engine.setCellValue("Sheet1", "A3", 3);
    engine.setCellValue("Sheet1", "B1", 3);
    engine.setCellValue("Sheet1", "B2", 2);
    engine.setCellValue("Sheet1", "B3", 1);
    engine.setCellValue("Sheet1", "C1", "apple");
    engine.setCellValue("Sheet1", "C2", "banana");
    engine.setCellValue("Sheet1", "C3", "pear");
    engine.setCellValue("Sheet1", "D1", true);
    engine.setCellValue("Sheet1", "D2", 2.5);
    engine.setCellValue("Sheet1", "D3", "peach");
    engine.setCellValue("Sheet1", "D4", 5);

    engine.setCellFormula("Sheet1", "F1", "MATCH(D1,A1:A3,1)");
    engine.setCellFormula("Sheet1", "F2", "MATCH(D2,B1:B3,-1)");
    engine.setCellFormula("Sheet1", "F3", "MATCH(D3,C1:C3,1)");
    engine.setCellFormula("Sheet1", "F4", "MATCH(D4,C1:C3,1)");

    const evaluation = getEvaluationService(engine);
    const f1Index = engine.workbook.getCellIndex("Sheet1", "F1");
    const f2Index = engine.workbook.getCellIndex("Sheet1", "F2");
    const f3Index = engine.workbook.getCellIndex("Sheet1", "F3");
    const f4Index = engine.workbook.getCellIndex("Sheet1", "F4");
    expect(f1Index).toBeDefined();
    expect(f2Index).toBeDefined();
    expect(f3Index).toBeDefined();
    expect(f4Index).toBeDefined();

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!));
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({ tag: ValueTag.Number, value: 1 });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!));
    expect(engine.getCellValue("Sheet1", "F2")).toEqual({ tag: ValueTag.Number, value: 1 });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!));
    expect(engine.getCellValue("Sheet1", "F3")).toEqual({ tag: ValueTag.Number, value: 2 });

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f4Index!));
    expect(engine.getCellValue("Sheet1", "F4")).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    engine.setCellValue("Sheet1", "A2", 4);
    engine.setCellValue("Sheet1", "A3", 5);
    engine.setCellValue("Sheet1", "D1", 4.5);
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!));
    expect(engine.getCellValue("Sheet1", "F1")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setRangeValues({ sheetName: "Sheet1", startAddress: "B1", endAddress: "B3" }, [
      [6],
      [5],
      [4],
    ]);
    engine.setCellValue("Sheet1", "D2", 4.5);
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!));
    expect(engine.getCellValue("Sheet1", "F2")).toEqual({ tag: ValueTag.Number, value: 2 });

    engine.setCellValue("Sheet1", "C2", "blueberry");
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!));
    expect(engine.getCellValue("Sheet1", "F3")).toEqual({ tag: ValueTag.Number, value: 2 });
  });
});
