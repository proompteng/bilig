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
  engine.importSnapshot(buildDownstreamSnapshot(downstreamCount));
}

export function buildDownstreamSnapshot(downstreamCount = 1000): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "benchmark-edit" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [
          { address: "A1", value: 1 },
          ...Array.from({ length: downstreamCount }, (_, index) => ({
            address: `B${index + 1}`,
            formula: `A1*2+${index + 1}`
          }))
        ]
      }
    ]
  };
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

export function seedRangeAggregateWorkbook(
  engine: SpreadsheetEngine,
  sourceCount = 1_024,
  aggregateCount = 10_000
): void {
  engine.importSnapshot(buildRangeAggregateSnapshot(sourceCount, aggregateCount));
}

export function buildRangeAggregateSnapshot(sourceCount = 1_024, aggregateCount = 10_000): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "benchmark-range-aggregates" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [
          ...Array.from({ length: sourceCount }, (_, index) => ({
            address: `A${index + 1}`,
            value: index + 1
          })),
          ...Array.from({ length: aggregateCount }, (_, index) => ({
            address: `B${index + 1}`,
            formula: `SUM(A1:A${sourceCount})+${index + 1}`
          }))
        ]
      }
    ]
  };
}

export function seedTopologyEditWorkbook(engine: SpreadsheetEngine, chainLength = 10_000): void {
  engine.importSnapshot(buildTopologyEditSnapshot(chainLength));
}

export function buildTopologyEditSnapshot(chainLength = 10_000): WorkbookSnapshot {
  const cells: WorkbookSnapshot["sheets"][number]["cells"] = [
    { address: "A1", value: 1 },
    { address: "A2", value: 2 },
    { address: "B1", formula: "A1*2" }
  ];

  for (let index = 2; index <= chainLength; index += 1) {
    cells.push({
      address: `B${index}`,
      formula: `B${index - 1}+1`
    });
  }

  return {
    version: 1,
    workbook: { name: "benchmark-topology-edit" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells
      }
    ]
  };
}
