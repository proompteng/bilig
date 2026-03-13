import type { SpreadsheetEngine } from "@bilig/core";

export function seedWorkbook(engine: SpreadsheetEngine, cellCount = 1000): void {
  engine.createSheet("Sheet1");
  for (let index = 1; index <= cellCount; index += 1) {
    engine.setCellValue("Sheet1", `A${index}`, index);
    engine.setCellFormula("Sheet1", `B${index}`, `A${index}*2`);
  }
}
