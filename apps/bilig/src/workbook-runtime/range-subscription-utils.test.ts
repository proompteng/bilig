import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { ValueTag, type EngineEvent } from "@bilig/protocol";
import {
  collectChangedAddressesForEvent,
  getRangeBounds,
  iterateRangeBounds,
} from "./range-subscription-utils.js";

function createMetrics() {
  return {
    batchId: 1,
    changedInputCount: 1,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    compileMs: 0,
  };
}

describe("range-subscription-utils", () => {
  it("collects changed addresses from direct cell indices and invalidated intersections", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setRangeValues({ sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" }, [
      [11, 12],
      [13, 14],
    ]);

    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 11 });

    const b2 = engine.workbook.getCellIndex("Sheet1", "B2");
    const c3 = engine.workbook.getCellIndex("Sheet1", "C3");
    if (b2 === undefined || c3 === undefined) {
      throw new Error("Expected seeded cells to exist");
    }

    const event: EngineEvent = {
      kind: "batch",
      invalidation: "cells",
      changedCellIndices: [b2, c3],
      invalidatedRanges: [
        { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        { sheetName: "Sheet1", startAddress: "D4", endAddress: "E5" },
      ],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: createMetrics(),
    };

    expect(
      collectChangedAddressesForEvent(
        engine,
        { sheetName: "Sheet1", startAddress: "B2", endAddress: "D4" },
        getRangeBounds({ sheetName: "Sheet1", startAddress: "B2", endAddress: "D4" }),
        event,
      ),
    ).toEqual(["C3", "B2", "D4"]);
  });

  it("expands full invalidation from precomputed bounds", () => {
    expect(
      iterateRangeBounds({
        startCol: 2,
        endCol: 3,
        startRow: 4,
        endRow: 5,
      }),
    ).toEqual(["B4", "C4", "B5", "C5"]);
  });
});
