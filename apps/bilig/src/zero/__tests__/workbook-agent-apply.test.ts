import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import { ValueTag } from "@bilig/protocol";
import { applyWorkbookAgentCommandBundleWithUndoCapture } from "../workbook-agent-apply.js";

function createBundle(
  overrides: Partial<WorkbookAgentCommandBundle> = {},
): WorkbookAgentCommandBundle {
  return {
    id: "bundle-1",
    documentId: "doc-1",
    threadId: "thr-1",
    turnId: "turn-1",
    goalText: "Populate prepaid expense template",
    summary: "Write cells in prepaid expenses!A1:I10",
    scope: "sheet",
    riskClass: "medium",
    approvalMode: "preview",
    baseRevision: 0,
    createdAtUnixMs: 1,
    context: null,
    commands: [],
    affectedRanges: [],
    estimatedAffectedCells: null,
    ...overrides,
  };
}

describe("workbook agent apply", () => {
  it("captures one undo bundle for multi-cell writeRange commands", async () => {
    const engine = new SpreadsheetEngine();
    await engine.ready();
    engine.createSheet("prepaid expenses");

    const bundle = createBundle({
      commands: [
        {
          kind: "writeRange",
          sheetName: "prepaid expenses",
          startAddress: "A1",
          values: [
            ["Expense", "Vendor"],
            ["Insurance", "Acme"],
          ],
        },
      ],
    });

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle);

    expect(engine.getCell("prepaid expenses", "A1").value).toMatchObject({
      tag: ValueTag.String,
      value: "Expense",
    });
    expect(engine.getCell("prepaid expenses", "B2").value).toMatchObject({
      tag: ValueTag.String,
      value: "Acme",
    });
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: "engineOps",
      }),
    );
    if (!undoBundle || undoBundle.kind !== "engineOps") {
      throw new Error("Expected engineOps undo bundle");
    }

    engine.applyOps(undoBundle.ops, { trusted: true });

    expect(engine.getCell("prepaid expenses", "A1").value).toEqual({
      tag: ValueTag.Empty,
    });
    expect(engine.getCell("prepaid expenses", "B2").value).toEqual({
      tag: ValueTag.Empty,
    });
  });

  it("captures one undo bundle when formatRange stages style and number format together", async () => {
    const engine = new SpreadsheetEngine();
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 123);

    const bundle = createBundle({
      commands: [
        {
          kind: "formatRange",
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A1",
          },
          patch: {
            font: {
              bold: true,
            },
          },
          numberFormat: "0.00",
        },
      ],
    });

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle);

    expect(engine.getCell("Sheet1", "A1").format).toBe("0.00");
    expect(engine.getCellStyle(engine.getCell("Sheet1", "A1").styleId)?.font?.bold).toBe(true);
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: "engineOps",
      }),
    );
  });

  it("captures undo for structural and cell commands in the same bundle", async () => {
    const engine = new SpreadsheetEngine();
    await engine.ready();

    const bundle = createBundle({
      commands: [
        {
          kind: "createSheet",
          name: "prepaid expenses",
        },
        {
          kind: "writeRange",
          sheetName: "prepaid expenses",
          startAddress: "A1",
          values: [["Expense"]],
        },
      ],
    });

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle);

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toContain("prepaid expenses");
    expect(engine.getCell("prepaid expenses", "A1").value).toMatchObject({
      tag: ValueTag.String,
      value: "Expense",
    });
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: "engineOps",
      }),
    );
    if (!undoBundle || undoBundle.kind !== "engineOps") {
      throw new Error("Expected engineOps undo bundle");
    }

    engine.applyOps(undoBundle.ops, { trusted: true });

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).not.toContain(
      "prepaid expenses",
    );
  });

  it("captures undo for row and column structural commands", async () => {
    const engine = new SpreadsheetEngine();
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", "Header");
    engine.setCellValue("Sheet1", "A2", "Value");
    engine.setCellValue("Sheet1", "B1", "Extra");

    const bundle = createBundle({
      commands: [
        {
          kind: "insertRows",
          sheetName: "Sheet1",
          start: 1,
          count: 1,
        },
        {
          kind: "deleteColumns",
          sheetName: "Sheet1",
          start: 1,
          count: 1,
        },
      ],
    });

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle);

    expect(engine.getCell("Sheet1", "A3").value).toMatchObject({
      tag: ValueTag.String,
      value: "Value",
    });
    expect(engine.getCell("Sheet1", "B1").value).toEqual({
      tag: ValueTag.Empty,
    });
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: "engineOps",
      }),
    );
    if (!undoBundle || undoBundle.kind !== "engineOps") {
      throw new Error("Expected engineOps undo bundle");
    }

    engine.applyOps(undoBundle.ops, { trusted: true });

    expect(engine.getCell("Sheet1", "A2").value).toMatchObject({
      tag: ValueTag.String,
      value: "Value",
    });
    expect(engine.getCell("Sheet1", "B1").value).toMatchObject({
      tag: ValueTag.String,
      value: "Extra",
    });
  });
});
