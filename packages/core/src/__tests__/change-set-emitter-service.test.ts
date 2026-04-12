import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import { createEngineChangeSetEmitterService } from "../engine/services/change-set-emitter-service.js";

describe("EngineChangeSetEmitterService", () => {
  it("captures tiny same-sheet change sets without requiring per-cell sheet-name lookups", () => {
    const engine = new SpreadsheetEngine({ workbookName: "change-set-emitter" });
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "B1", 2);

    const emitter = createEngineChangeSetEmitterService({
      state: {
        workbook: engine.workbook,
        strings: engine.strings,
      },
    });

    const a1 = engine.workbook.getCellIndex("Sheet1", "A1");
    const b1 = engine.workbook.getCellIndex("Sheet1", "B1");
    expect(a1).toBeDefined();
    expect(b1).toBeDefined();

    const changes = emitter.captureChangedCells([a1!, b1!]);

    expect(changes).toHaveLength(2);
    expect(changes[0]?.sheetName).toBe("Sheet1");
    expect(changes[1]?.sheetName).toBe("Sheet1");
    expect(changes[0]?.newValue).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(changes[1]?.newValue).toEqual({ tag: ValueTag.Number, value: 2 });
  });
});
