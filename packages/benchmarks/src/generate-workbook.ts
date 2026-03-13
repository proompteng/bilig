import type { CommitOp, SpreadsheetEngine } from "@bilig/core";
import type { WorkbookSnapshot } from "@bilig/protocol";

export function seedWorkbook(engine: SpreadsheetEngine, materializedCells = 1000): void {
  seedLoadWorkbook(engine, materializedCells);
}

export function seedLoadWorkbook(engine: SpreadsheetEngine, materializedCells = 1000): void {
  engine.importSnapshot(buildWorkbookSnapshot(materializedCells));
}

export function buildWorkbookSnapshot(materializedCells = 1000): WorkbookSnapshot {
  const literalCount = Math.max(1, Math.ceil(materializedCells / 2));
  const formulaCount = Math.max(0, materializedCells - literalCount);

  return {
    version: 1,
    workbook: { name: "benchmark-load" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [
          ...Array.from({ length: literalCount }, (_, index) => ({
            address: `A${index + 1}`,
            value: index + 1
          })),
          ...Array.from({ length: formulaCount }, (_, index) => ({
            address: `B${index + 1}`,
            formula: `A${index + 1}*2`
          }))
        ]
      }
    ]
  };
}

export function seedDownstreamWorkbook(engine: SpreadsheetEngine, downstreamCount = 1000): void {
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 1);

  for (let index = 1; index <= downstreamCount; index += 1) {
    engine.setCellFormula("Sheet1", `B${index}`, `A1*2+${index}`);
  }
}

export function buildRenderCommitOps(cellCount = 1000): CommitOp[] {
  const ops: CommitOp[] = [
    { kind: "upsertWorkbook", name: "benchmark-renderer" },
    { kind: "upsertSheet", name: "Sheet1", order: 0 }
  ];

  for (let index = 1; index <= cellCount; index += 1) {
    ops.push({ kind: "upsertCell", sheetName: "Sheet1", addr: `A${index}`, value: index });
  }

  return ops;
}
