import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag } from "@bilig/protocol";
import { describe, expect, it } from "vitest";
import { applyWorkbookEvent, deriveDirtyRegions } from "../events.js";

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

  it("restores snapshots through replaceSnapshot events", async () => {
    const source = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "source",
    });
    await source.ready();
    source.setCellValue("Sheet1", "A1", 12);
    const sourceSnapshot = source.exportSnapshot();

    const restored = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "restored",
    });
    await restored.ready();
    applyWorkbookEvent(restored, {
      kind: "replaceSnapshot",
      snapshot: sourceSnapshot,
      replicaSnapshot: source.exportReplicaSnapshot(),
    });

    expect(restored.exportSnapshot()).toEqual(sourceSnapshot);
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
