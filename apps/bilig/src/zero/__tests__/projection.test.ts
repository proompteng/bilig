import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { describe, expect, it } from "vitest";
import {
  buildSheetCellSourceRows,
  buildSheetCellSourceRowsFromEngine,
  diffProjectionRows,
  materializeCellEvalProjection,
  sourceProjectionKeys,
} from "../projection.js";

describe("projection helpers", () => {
  it("can diff semantic source rows without treating revision churn as a rewrite", () => {
    const previous = [
      {
        workbookId: "doc-1",
        sheetName: "Sheet1",
        address: "A1",
        rowNum: 0,
        colNum: 0,
        inputValue: 21,
        formula: null,
        format: null,
        styleId: null,
        explicitFormatId: null,
        sourceRevision: 1,
        updatedBy: "user-a",
        updatedAt: "2026-03-28T10:00:00.000Z",
      },
    ];
    const next = [
      {
        ...previous[0],
        sourceRevision: 2,
        updatedBy: "user-b",
        updatedAt: "2026-03-28T10:01:00.000Z",
      },
    ];

    const diff = diffProjectionRows(previous, next, sourceProjectionKeys.cell, (row) =>
      JSON.stringify([
        row.sheetName,
        row.address,
        row.rowNum,
        row.colNum,
        row.inputValue,
        row.formula,
        row.format,
        row.styleId,
        row.explicitFormatId,
      ]),
    );

    expect(diff.upserts).toHaveLength(0);
    expect(diff.deletes).toHaveLength(0);
  });

  it("keys sheet projection rows by stable sheet id instead of the mutable name", () => {
    expect(
      sourceProjectionKeys.sheet({
        workbookId: "doc-1",
        sheetId: 7,
        name: "Sheet1",
        sortOrder: 0,
        freezeRows: 0,
        freezeCols: 0,
        updatedAt: "2026-04-06T10:00:00.000Z",
      }),
    ).toBe(
      sourceProjectionKeys.sheet({
        workbookId: "doc-1",
        sheetId: 7,
        name: "Revenue",
        sortOrder: 0,
        freezeRows: 0,
        freezeCols: 0,
        updatedAt: "2026-04-06T10:01:00.000Z",
      }),
    );
  });

  it("materializes authoritative cell_eval rows from the workbook engine, including styled blanks", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "projection-test",
    });
    await engine.ready();
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*3");
    engine.setRangeStyle(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "C2" },
      { fill: { backgroundColor: "#abcdef" } },
    );
    engine.setRangeNumberFormat(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "C2" },
      { kind: "number", code: "$#,##0.00" },
    );

    const rows = materializeCellEvalProjection(engine, "doc-1", 9, "2026-03-28T10:02:00.000Z");
    const b1 = rows.find((row) => row.sheetName === "Sheet1" && row.address === "B1");
    const c2 = rows.find((row) => row.sheetName === "Sheet1" && row.address === "C2");

    expect(b1).toEqual(
      expect.objectContaining({
        workbookId: "doc-1",
        rowNum: 0,
        colNum: 1,
        styleId: expect.any(String),
        styleJson: expect.objectContaining({
          fill: { backgroundColor: "#abcdef" },
        }),
        formatId: expect.any(String),
        formatCode: expect.any(String),
        calcRevision: 9,
      }),
    );
    expect(b1?.value).toEqual({
      tag: ValueTag.Number,
      value: 21,
    });
    expect(c2).toEqual(
      expect.objectContaining({
        workbookId: "doc-1",
        rowNum: 1,
        colNum: 2,
        value: { tag: ValueTag.Empty },
        styleId: expect.any(String),
        styleJson: expect.objectContaining({
          fill: { backgroundColor: "#abcdef" },
        }),
        formatId: expect.any(String),
        formatCode: expect.any(String),
      }),
    );
  });

  it("flattens style and format ranges into sparse source cell rows", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "projection-test",
    });
    await engine.ready();
    engine.setRangeStyle(
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
      { fill: { backgroundColor: "#abcdef" } },
    );
    engine.setRangeNumberFormat(
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
      { kind: "number", code: "$#,##0.00" },
    );

    const snapshot = engine.exportSnapshot();
    const sourceRows = buildSheetCellSourceRows("doc-a", snapshot, "Sheet1", {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-a",
      updatedBy: "user-a",
      updatedAt: "2026-03-29T10:00:00.000Z",
    });
    const b2 = sourceRows.find((row) => row.address === "B2");
    const c3 = sourceRows.find((row) => row.address === "C3");

    expect(b2).toEqual(
      expect.objectContaining({
        sheetName: "Sheet1",
        address: "B2",
        styleId: expect.any(String),
        explicitFormatId: expect.any(String),
        format: expect.any(String),
      }),
    );
    expect(c3).toEqual(
      expect.objectContaining({
        sheetName: "Sheet1",
        address: "C3",
        styleId: expect.any(String),
        explicitFormatId: expect.any(String),
        format: expect.any(String),
      }),
    );
  });

  it("materializes the same sparse source rows directly from the engine", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "projection-test",
    });
    await engine.ready();
    engine.setCellValue("Sheet1", "A1", 99);
    engine.setRangeStyle(
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
      { fill: { backgroundColor: "#abcdef" } },
    );
    engine.setRangeNumberFormat(
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
      { kind: "number", code: "$#,##0.00" },
    );

    const options = {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-a",
      updatedBy: "user-a",
      updatedAt: "2026-03-29T10:00:00.000Z",
    } as const;

    expect(buildSheetCellSourceRowsFromEngine("doc-a", engine, "Sheet1", options)).toEqual(
      buildSheetCellSourceRows("doc-a", engine.exportSnapshot(), "Sheet1", options),
    );
  });
});
