import { WorkPaper } from "@bilig/headless";
import { address } from "./workpaper-benchmark-fixtures.js";
import {
  buildConditionalAggregationSheet,
  buildLookupSheet,
  buildApproxLookupSheet,
  buildMixedFrontierSheet,
  buildOverlappingAggregateSheet,
  buildParserCacheTemplateSheet,
  buildValueFormulaRows,
  buildBatchMultiColumnRows,
} from "./workpaper-benchmark-fixtures.js";
import {
  HyperFormula,
  measureHyperFormulaBuildFromSheets,
  measureHyperFormulaMutationSample,
  measureMutationSample,
  measureWorkPaperBuildFromSheets,
  normalizeHyperFormulaValue,
  normalizeWorkPaperValue,
  toHyperFormulaSheet,
  type BenchmarkSample,
} from "./benchmark-workpaper-vs-hyperformula-expanded-support.js";
import { HYPERFORMULA_LICENSE_KEY } from "./benchmark-workpaper-vs-hyperformula.js";

export function measureWorkPaperSuspendedBatchSingleColumnEditSample(
  editCount: number,
): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation();
      for (let row = 0; row < editCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 7);
      }
      return workbook.resumeEvaluation();
    },
    () => ({
      sampleFormulaValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, editCount - 1, 1)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

export function measureHyperFormulaSuspendedBatchSingleColumnEditSample(
  editCount: number,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation();
      for (let row = 0; row < editCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 7);
      }
      return workbook.resumeEvaluation();
    },
    () => ({
      sampleFormulaValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, editCount - 1, 1)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

export function measureWorkPaperSuspendedBatchMultiColumnEditSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildBatchMultiColumnRows(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation();
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 3);
        workbook.setCellContents(address(sheetId, row, 1), row * 5);
      }
      return workbook.resumeEvaluation();
    },
    () => ({
      sampleSumValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 2)),
      ),
      sampleProductValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 3)),
      ),
    }),
  );
}

export function measureHyperFormulaSuspendedBatchMultiColumnEditSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildBatchMultiColumnRows(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation();
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 3);
        workbook.setCellContents(address(sheetId, row, 1), row * 5);
      }
      return workbook.resumeEvaluation();
    },
    () => ({
      sampleSumValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 2)),
      ),
      sampleProductValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 3)),
      ),
    }),
  );
}

export function measureWorkPaperRebuildAndRecalculateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildParserCacheTemplateSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.rebuildAndRecalculate(),
    () => ({
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 4)),
      ),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  );
}

export function measureHyperFormulaRebuildAndRecalculateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildParserCacheTemplateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.rebuildAndRecalculate(),
    () => ({
      terminalValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 4)),
      ),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  );
}

export function measureWorkPaperConfigToggleSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(
    { Bench: buildLookupSheet(rowCount) },
    { useColumnIndex: false },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.updateConfig({ useColumnIndex: true }),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}

export function measureHyperFormulaConfigToggleSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: false },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.updateConfig({ useColumnIndex: true }),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}

export function measureWorkPaperStructuralInsertRowsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.addRows(sheetId, Math.floor(rowCount / 2), 1),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount, 1))),
    }),
  );
}

export function measureHyperFormulaStructuralInsertRowsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.addRows(sheetId, [Math.floor(rowCount / 2), 1]),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount, 1))),
    }),
  );
}

export function measureWorkPaperOverlappingAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 1)),
      ),
    }),
  );
}

export function measureHyperFormulaOverlappingAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 1)),
      ),
    }),
  );
}

export function measureWorkPaperConditionalAggregationSample(
  rowCount: number,
  formulaCopies: number,
): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildConditionalAggregationSheet(rowCount, formulaCopies),
  });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 1), rowCount * 2),
    () => ({
      sumifValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies)),
      ),
    }),
  );
}

export function measureHyperFormulaConditionalAggregationSample(
  rowCount: number,
  formulaCopies: number,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildConditionalAggregationSheet(rowCount, formulaCopies)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 1), rowCount * 2),
    () => ({
      sumifValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies)),
      ),
    }),
  );
}

export function measureWorkPaperParserCacheTemplateSample(rowCount: number): BenchmarkSample {
  return measureWorkPaperBuildFromSheets(
    { Bench: buildParserCacheTemplateSheet(rowCount) },
    (workbook) => {
      const sheetId = workbook.getSheetId("Bench")!;
      return {
        dimensions: workbook.getSheetDimensions(sheetId),
        terminalValue: normalizeWorkPaperValue(
          workbook.getCellValue(address(sheetId, rowCount - 1, 4)),
        ),
      };
    },
  );
}

export function measureHyperFormulaParserCacheTemplateSample(rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets(
    { Bench: toHyperFormulaSheet(buildParserCacheTemplateSheet(rowCount)) },
    (workbook) => {
      const sheetId = workbook.getSheetId("Bench")!;
      return {
        dimensions: workbook.getSheetDimensions(sheetId),
        terminalValue: normalizeHyperFormulaValue(
          workbook.getCellValue(address(sheetId, rowCount - 1, 4)),
        ),
      };
    },
  );
}

export function measureWorkPaperMixedFrontierSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildMixedFrontierSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalAggregate: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 2)),
      ),
      terminalFanout: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 1)),
      ),
    }),
  );
}

export function measureHyperFormulaMixedFrontierSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildMixedFrontierSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalAggregate: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 2)),
      ),
      terminalFanout: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 1)),
      ),
    }),
  );
}

export function measureWorkPaperIndexedLookupAfterColumnWriteSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(
    { Bench: buildLookupSheet(rowCount) },
    { useColumnIndex: true },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}

export function measureHyperFormulaIndexedLookupAfterColumnWriteSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: true },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}

export function measureWorkPaperApproximateLookupAfterColumnWriteSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}

export function measureHyperFormulaApproximateLookupAfterColumnWriteSample(
  rowCount: number,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  );
}
