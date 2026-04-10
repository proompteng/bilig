import { performance } from "node:perf_hooks";
import { WorkPaper, type WorkPaperCellAddress, type WorkPaperSheet } from "@bilig/headless";
import type { RecalcMetrics } from "@bilig/protocol";
import { measureMemory, sampleMemory, type MemoryMeasurement } from "./metrics.js";

export type WorkPaperBenchmarkScenario =
  | "workpaper-build"
  | "workpaper-single-edit"
  | "workpaper-batch-edit"
  | "workpaper-range-read"
  | "workpaper-lookup"
  | "workpaper-dynamic-array";

export interface WorkPaperBenchmarkResult {
  scenario: WorkPaperBenchmarkScenario;
  elapsedMs: number;
  memory: MemoryMeasurement;
  details: Record<string, number | boolean | string>;
  metrics?: RecalcMetrics;
}

export async function runWorkPaperBenchmarkSuite(): Promise<WorkPaperBenchmarkResult[]> {
  return [
    runWorkPaperBuildBenchmark(),
    runWorkPaperSingleEditBenchmark(),
    runWorkPaperBatchEditBenchmark(),
    runWorkPaperRangeReadBenchmark(),
    runWorkPaperLookupBenchmark(),
    runWorkPaperDynamicArrayBenchmark(),
  ];
}

export function runWorkPaperBuildBenchmark(rows = 160, cols = 24): WorkPaperBenchmarkResult {
  const sheet = buildDenseLiteralSheet(rows, cols);
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const workbook = WorkPaper.buildFromSheets({ Bench: sheet });
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  const sheetId = workbook.getSheetId("Bench")!;

  return {
    scenario: "workpaper-build",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      width: workbook.getSheetDimensions(sheetId).width,
      height: workbook.getSheetDimensions(sheetId).height,
    },
    metrics: workbook.getStats().lastMetrics,
  };
}

export function runWorkPaperSingleEditBenchmark(downstreamCount = 2_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: [buildFormulaChainRow(downstreamCount)],
  });
  const sheetId = workbook.getSheetId("Bench")!;
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const changes = workbook.setCellContents(address(sheetId, 0, 0), 99);
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "workpaper-single-edit",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      downstreamCount,
      changeCount: changes.length,
      terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? "",
    },
    metrics: workbook.getStats().lastMetrics,
  };
}

export function runWorkPaperBatchEditBenchmark(editCount = 500): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildValueFormulaRows(editCount),
  });
  const sheetId = workbook.getSheetId("Bench")!;
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const changes = workbook.batch(() => {
    for (let row = 0; row < editCount; row += 1) {
      workbook.setCellContents(address(sheetId, row, 0), row * 3);
    }
  });
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "workpaper-batch-edit",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      editCount,
      changeCount: changes.length,
      sampleFormulaValue: JSON.stringify(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    },
    metrics: workbook.getStats().lastMetrics,
  };
}

export function runWorkPaperRangeReadBenchmark(rows = 240, cols = 24): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDenseLiteralSheet(rows, cols),
  });
  const sheetId = workbook.getSheetId("Bench")!;
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const values = workbook.getRangeValues({
    start: address(sheetId, 0, 0),
    end: address(sheetId, rows - 1, cols - 1),
  });
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "workpaper-range-read",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      readRows: values.length,
      readCols: values[0]?.length ?? 0,
    },
  };
}

export function runWorkPaperLookupBenchmark(rowCount = 5_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets(
    {
      Bench: buildLookupSheet(rowCount),
    },
    {
      useColumnIndex: true,
    },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const changes = workbook.setCellContents(targetAddress, rowCount);
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "workpaper-lookup",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount,
      changeCount: changes.length,
      formulaValue: JSON.stringify(workbook.getCellValue(formulaAddress)),
    },
    metrics: workbook.getStats().lastMetrics,
  };
}

export function runWorkPaperDynamicArrayBenchmark(rowCount = 750): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDynamicArraySheet(rowCount),
  });
  const sheetId = workbook.getSheetId("Bench")!;
  const thresholdAddress = address(sheetId, 0, 1);
  const spillAnchor = address(sheetId, 0, 2);
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const changes = workbook.setCellContents(thresholdAddress, rowCount - 10);
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();

  return {
    scenario: "workpaper-dynamic-array",
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount,
      changeCount: changes.length,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillHeight: workbook.getSheetDimensions(sheetId).height,
    },
    metrics: workbook.getStats().lastMetrics,
  };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  console.log(JSON.stringify(await runWorkPaperBenchmarkSuite(), null, 2));
}

function buildDenseLiteralSheet(rows: number, cols: number): WorkPaperSheet {
  return Array.from({ length: rows }, (_rowValue, rowIndex) =>
    Array.from({ length: cols }, (_colValue, colIndex) => rowIndex * cols + colIndex + 1),
  );
}

function buildFormulaChainRow(downstreamCount: number): readonly (number | string)[] {
  const row: Array<number | string> = [1];
  for (let col = 1; col <= downstreamCount; col += 1) {
    row.push(`=${columnLabel(col - 1)}1+1`);
  }
  return row;
}

function buildValueFormulaRows(rowCount: number): WorkPaperSheet {
  return Array.from({ length: rowCount }, (_, index) => [index + 1, `=A${index + 1}*2`]);
}

function buildLookupSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    ["Key", "Value", "", Math.floor(rowCount / 2), "=MATCH(D1,A2:A20001,0)"],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index, index * 10]);
  }
  const headerRow = rows[0];
  if (!headerRow) {
    throw new Error("lookup benchmark header row is missing");
  }
  headerRow[4] = `=MATCH(D1,A2:A${rowCount + 1},0)`;
  return rows;
}

function buildDynamicArraySheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<number | string>> = [
    [0, Math.floor(rowCount / 2), `=FILTER(A2:A${rowCount + 1},A2:A${rowCount + 1}>B1)`],
  ];
  for (let index = 1; index <= rowCount; index += 1) {
    rows.push([index]);
  }
  return rows;
}

function address(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col };
}

function columnLabel(index: number): string {
  let current = index;
  let label = "";
  while (current >= 0) {
    label = String.fromCharCode(65 + (current % 26)) + label;
    current = Math.floor(current / 26) - 1;
  }
  return label;
}
