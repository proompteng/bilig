import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";

describe("SpreadsheetEngine formula initialization", () => {
  it("initializes formula refs without emitting watched events or batches", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "engine-formula-initialize" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 4);
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;

    const events: string[] = [];
    const batches: unknown[] = [];
    const unsubscribeEvents = engine.subscribe((event) => {
      events.push(event.kind);
    });
    const unsubscribeBatches = engine.subscribeBatches((batch) => {
      batches.push(batch);
    });
    events.length = 0;
    batches.length = 0;

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: "setCellFormula", row: 0, col: 1, formula: "A1*2" } },
        { sheetId, mutation: { kind: "setCellFormula", row: 0, col: 2, formula: "B1+1" } },
      ],
      2,
    );

    unsubscribeEvents();
    unsubscribeBatches();

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Number, value: 9 });
    expect(events).toEqual([]);
    expect(batches).toEqual([]);
  });

  it("initializes invalid formulas and propagates their errors through dependent formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "engine-formula-initialize-errors" });
    await engine.ready();
    engine.createSheet("Sheet1");
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: "setCellFormula", row: 0, col: 1, formula: "SUM(" } },
        { sheetId, mutation: { kind: "setCellFormula", row: 0, col: 2, formula: "B1+1" } },
      ],
      2,
    );

    expect(engine.getCellValue("Sheet1", "B1")).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    });
    expect(engine.getCellValue("Sheet1", "C1")).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    });
  });
});
