import { describe, expect, it } from "vitest";

import { SpreadsheetEngine } from "../engine.js";

describe("sheet ids", () => {
  it("preserves stable sheet ids across export, import, and rename", () => {
    const engine = new SpreadsheetEngine({ workbookName: "sheet-id-doc" });
    engine.createSheet("Sheet1");
    engine.createSheet("Summary");
    engine.setCellValue("Sheet1", "A1", 11);
    engine.setCellValue("Summary", "B2", 22);
    engine.renameSheet("Sheet1", "Revenue");

    const exported = engine.exportSnapshot();
    expect(exported.sheets).toEqual([
      expect.objectContaining({ id: 1, name: "Revenue", order: 0 }),
      expect.objectContaining({ id: 2, name: "Summary", order: 1 }),
    ]);

    const restored = new SpreadsheetEngine({ workbookName: "restored-sheet-id-doc" });
    restored.importSnapshot(exported);

    expect(restored.exportSnapshot()).toEqual(exported);
  });
});
