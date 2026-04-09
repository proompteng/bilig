import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import type { FormulaNode } from "@bilig/formula";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineFormulaEvaluationService } from "../engine/services/formula-evaluation-service.js";
import type { EngineMutationSupportService } from "../engine/services/mutation-support-service.js";

function isEngineFormulaEvaluationService(
  value: unknown,
): value is EngineFormulaEvaluationService {
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
  });
});
