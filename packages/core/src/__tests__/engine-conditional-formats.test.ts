import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";

describe("SpreadsheetEngine conditional formats", () => {
  it("roundtrips conditional format metadata through snapshots", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "conditional-format-roundtrip" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "cellIs",
        operator: "greaterThan",
        values: [10],
      },
      style: {
        fill: { backgroundColor: "#ff0000" },
      },
      stopIfTrue: true,
      priority: 1,
    });

    const snapshot = engine.exportSnapshot();
    expect(
      snapshot.sheets.find((sheet) => sheet.name === "Sheet1")?.metadata?.conditionalFormats,
    ).toEqual([
      {
        id: "cf-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
        rule: {
          kind: "cellIs",
          operator: "greaterThan",
          values: [10],
        },
        style: {
          fill: { backgroundColor: "#ff0000" },
        },
        stopIfTrue: true,
        priority: 1,
      },
    ]);

    const restored = new SpreadsheetEngine({ workbookName: "conditional-format-restored" });
    await restored.ready();
    restored.importSnapshot(snapshot);

    expect(restored.getConditionalFormats("Sheet1")).toEqual([
      {
        id: "cf-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
        rule: {
          kind: "cellIs",
          operator: "greaterThan",
          values: [10],
        },
        style: {
          fill: { backgroundColor: "#ff0000" },
        },
        stopIfTrue: true,
        priority: 1,
      },
    ]);
  });

  it("rewrites conditional format ranges across structural edits", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "conditional-format-structural" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "textContains",
        text: "urgent",
      },
      style: {
        font: { bold: true },
      },
    });

    engine.insertRows("Sheet1", 1, 1);
    expect(engine.getConditionalFormats("Sheet1")).toEqual([
      {
        id: "cf-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "B3",
          endAddress: "B5",
        },
        rule: {
          kind: "textContains",
          text: "urgent",
        },
        style: {
          font: { bold: true },
        },
      },
    ]);

    engine.deleteColumns("Sheet1", 0, 1);
    expect(engine.getConditionalFormats("Sheet1")).toEqual([
      {
        id: "cf-1",
        range: {
          sheetName: "Sheet1",
          startAddress: "A3",
          endAddress: "A5",
        },
        rule: {
          kind: "textContains",
          text: "urgent",
        },
        style: {
          font: { bold: true },
        },
      },
    ]);

    expect(engine.undo()).toBe(true);
    expect(engine.getConditionalFormats("Sheet1")[0]?.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "B3",
      endAddress: "B5",
    });
  });

  it("deletes conditional formats through the public API and reports missing ids", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "conditional-format-delete" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A3",
      },
      rule: {
        kind: "cellIs",
        operator: "greaterThan",
        values: [5],
      },
      style: {
        fill: { backgroundColor: "#00ff00" },
      },
    });

    expect(engine.getConditionalFormat("cf-1")).toMatchObject({ id: "cf-1" });
    expect(engine.deleteConditionalFormat("cf-1")).toBe(true);
    expect(engine.getConditionalFormat("cf-1")).toBeUndefined();
    expect(engine.getConditionalFormats("Sheet1")).toEqual([]);
    expect(engine.deleteConditionalFormat("cf-1")).toBe(false);
  });
});
