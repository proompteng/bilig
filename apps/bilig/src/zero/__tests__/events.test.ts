import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { describe, expect, it } from "vitest";
import { applyWorkbookEvent, deriveDirtyRegions } from "@bilig/zero-sync";

describe("workbook events", () => {
  it("replays workbook mutations onto a warm engine", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "event-test",
    });
    await engine.ready();

    applyWorkbookEvent(engine, {
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A1",
      value: 5,
    });
    applyWorkbookEvent(engine, {
      kind: "setCellFormula",
      sheetName: "Sheet1",
      address: "B1",
      formula: "A1*4",
    });

    expect(engine.getCell("Sheet1", "B1").value).toEqual({
      tag: ValueTag.Number,
      value: 20,
    });
  });

  it("replays revertChange events from persisted undo bundles", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "event-test",
    });
    await engine.ready();

    applyWorkbookEvent(engine, {
      kind: "setCellValue",
      sheetName: "Sheet1",
      address: "A1",
      value: "seed",
    });
    expect(engine.getCellValue("Sheet1", "A1")).toMatchObject({
      tag: ValueTag.String,
      value: "seed",
    });

    applyWorkbookEvent(engine, {
      kind: "revertChange",
      targetRevision: 1,
      targetSummary: "Updated Sheet1!A1",
      sheetName: "Sheet1",
      address: "A1",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A1",
      },
      appliedBundle: {
        kind: "engineOps",
        ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "A1" }],
      },
    });

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });
  });

  it("replays structural metadata events onto a warm engine", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "event-test",
    });
    await engine.ready();

    applyWorkbookEvent(engine, {
      kind: "updateRowMetadata",
      sheetName: "Sheet1",
      startRow: 2,
      count: 1,
      height: 36,
      hidden: true,
    });
    applyWorkbookEvent(engine, {
      kind: "setFreezePane",
      sheetName: "Sheet1",
      rows: 1,
      cols: 2,
    });

    expect(engine.getRowMetadata("Sheet1")).toEqual([
      {
        count: 1,
        hidden: true,
        sheetName: "Sheet1",
        size: 36,
        start: 2,
      },
    ]);
    expect(engine.getFreezePane("Sheet1")).toEqual({ sheetName: "Sheet1", rows: 1, cols: 2 });
    expect(
      deriveDirtyRegions({
        kind: "setFreezePane",
        sheetName: "Sheet1",
        rows: 1,
        cols: 2,
      }),
    ).toBeNull();
  });

  it("derives focused dirty regions for common source edits", () => {
    expect(
      deriveDirtyRegions({
        kind: "setCellValue",
        sheetName: "Sheet1",
        address: "C4",
        value: 1,
      }),
    ).toEqual([
      {
        sheetName: "Sheet1",
        rowStart: 3,
        rowEnd: 3,
        colStart: 2,
        colEnd: 2,
      },
    ]);

    expect(
      deriveDirtyRegions({
        kind: "fillRange",
        source: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A2",
        },
        target: {
          sheetName: "Sheet1",
          startAddress: "B1",
          endAddress: "B4",
        },
      }),
    ).toEqual([
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 0,
      },
      {
        sheetName: "Sheet1",
        rowStart: 0,
        rowEnd: 3,
        colStart: 1,
        colEnd: 1,
      },
    ]);
  });
});
