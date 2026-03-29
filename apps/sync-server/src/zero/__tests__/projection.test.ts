import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { describe, expect, it } from "vitest";
import {
  buildSheetFormatRangeRows,
  buildSheetStyleRangeRows,
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
        row.explicitFormatId,
      ]),
    );

    expect(diff.upserts).toHaveLength(0);
    expect(diff.deletes).toHaveLength(0);
  });

  it("materializes authoritative cell_eval rows from the workbook engine", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "projection-test",
    });
    await engine.ready();
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*3");
    engine.setRangeStyle(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
      { fill: { backgroundColor: "#abcdef" } },
    );
    engine.setRangeNumberFormat(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
      { kind: "number", code: "$#,##0.00" },
    );

    const rows = materializeCellEvalProjection(engine, "doc-1", 9, "2026-03-28T10:02:00.000Z");
    const b1 = rows.find((row) => row.sheetName === "Sheet1" && row.address === "B1");

    expect(b1).toEqual(
      expect.objectContaining({
        workbookId: "doc-1",
        rowNum: 0,
        colNum: 1,
        styleId: expect.any(String),
        formatId: expect.any(String),
        formatCode: expect.any(String),
        calcRevision: 9,
      }),
    );
    expect(b1?.value).toEqual({
      tag: ValueTag.Number,
      value: 21,
    });
  });

  it("names style and format range ids per workbook to avoid cross-document collisions", async () => {
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
    const styleRowsA = buildSheetStyleRangeRows("doc-a", snapshot, "Sheet1", {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-a",
      updatedBy: "user-a",
      updatedAt: "2026-03-29T10:00:00.000Z",
    });
    const styleRowsB = buildSheetStyleRangeRows("doc-b", snapshot, "Sheet1", {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-b",
      updatedBy: "user-b",
      updatedAt: "2026-03-29T10:00:00.000Z",
    });
    const formatRowsA = buildSheetFormatRangeRows("doc-a", snapshot, "Sheet1", {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-a",
      updatedBy: "user-a",
      updatedAt: "2026-03-29T10:00:00.000Z",
    });
    const formatRowsB = buildSheetFormatRangeRows("doc-b", snapshot, "Sheet1", {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "owner-b",
      updatedBy: "user-b",
      updatedAt: "2026-03-29T10:00:00.000Z",
    });

    expect(styleRowsA.map((row) => row.id)).not.toEqual(styleRowsB.map((row) => row.id));
    expect(formatRowsA.map((row) => row.id)).not.toEqual(formatRowsB.map((row) => row.id));
    expect(styleRowsA[0]?.id).toContain("doc-a");
    expect(styleRowsB[0]?.id).toContain("doc-b");
    expect(formatRowsA[0]?.id).toContain("doc-a");
    expect(formatRowsB[0]?.id).toContain("doc-b");
  });
});
