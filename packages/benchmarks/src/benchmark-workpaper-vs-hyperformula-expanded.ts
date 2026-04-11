import { performance } from "node:perf_hooks";
import { WorkPaper } from "@bilig/headless";
import { ValueTag } from "@bilig/protocol";
import type {
  RawCellContent as HyperFormulaRawCellContent,
  Sheet as HyperFormulaSheet,
} from "hyperformula";
import type {
  ComparativeBenchmarkSuiteOptions,
  ComparativeMeasuredEngineResult,
  ComparativeMemorySummary,
  ComparativeUnsupportedEngineResult,
} from "./benchmark-workpaper-vs-hyperformula.js";
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  HYPERFORMULA_LICENSE_KEY,
} from "./benchmark-workpaper-vs-hyperformula.js";
import type { MemoryMeasurement } from "./metrics.js";
import { measureMemory, sampleMemory } from "./metrics.js";
import { summarizeNumbers } from "./stats.js";
import {
  address,
  buildApproxLookupSheet,
  buildBatchMultiColumnRows,
  buildDenseLiteralSheet,
  buildDynamicArraySheet,
  buildFormulaChainRow,
  buildFormulaEditChainRow,
  buildFormulaFanoutRow,
  buildLookupSheet,
  buildMixedContentSheet,
  buildMultiSheetLiteralSheets,
  buildTextLookupSheet,
  buildValueFormulaRows,
  range,
} from "./workpaper-benchmark-fixtures.js";

const { HyperFormula } = await import("hyperformula");
type HyperFormulaInstance = ReturnType<typeof HyperFormula.buildFromSheets>;

export type ExpandedComparativeBenchmarkWorkload =
  | "build-dense-literals"
  | "build-mixed-content"
  | "build-many-sheets"
  | "single-edit-chain"
  | "single-edit-fanout"
  | "single-formula-edit-recalc"
  | "batch-edit-single-column"
  | "batch-edit-multi-column"
  | "range-read-dense"
  | "lookup-no-column-index"
  | "lookup-with-column-index"
  | "lookup-approximate-sorted"
  | "lookup-text-exact"
  | "dynamic-array-filter";

export interface ExpandedComparativeComparableResult {
  workload: ExpandedComparativeBenchmarkWorkload;
  category: "directly-comparable";
  comparable: true;
  fixture: Record<string, unknown>;
  comparison: {
    fasterEngine: "workpaper" | "hyperformula";
    meanSpeedup: number;
    verificationEquivalent: true;
  };
  engines: {
    workpaper: ComparativeMeasuredEngineResult;
    hyperformula: ComparativeMeasuredEngineResult;
  };
}

export interface ExpandedComparativeLeadershipResult {
  workload: ExpandedComparativeBenchmarkWorkload;
  category: "leadership";
  comparable: false;
  fixture: Record<string, unknown>;
  note: string;
  engines: {
    workpaper: ComparativeMeasuredEngineResult;
    hyperformula: ComparativeUnsupportedEngineResult;
  };
}

interface BenchmarkSample {
  elapsedMs: number;
  memory: MemoryMeasurement;
  verification: Record<string, unknown>;
}

export type ExpandedComparativeBenchmarkResult =
  | ExpandedComparativeComparableResult
  | ExpandedComparativeLeadershipResult;

export function runWorkPaperVsHyperFormulaExpandedBenchmarkSuite(
  options: ComparativeBenchmarkSuiteOptions = {},
): ExpandedComparativeBenchmarkResult[] {
  const runtimeOptions = resolveSuiteOptions(options);
  return [
    runComparableScenario(
      "build-dense-literals",
      { cols: 24, rows: 160, materializedCells: 160 * 24 },
      runtimeOptions,
      () => measureWorkPaperDenseBuildSample(160, 24),
      () => measureHyperFormulaDenseBuildSample(160, 24),
    ),
    runComparableScenario(
      "build-mixed-content",
      { cols: 6, rows: 750 },
      runtimeOptions,
      () => measureWorkPaperMixedBuildSample(750),
      () => measureHyperFormulaMixedBuildSample(750),
    ),
    runComparableScenario(
      "build-many-sheets",
      { sheetCount: 8, rowsPerSheet: 120, colsPerSheet: 12 },
      runtimeOptions,
      () => measureWorkPaperManySheetsBuildSample(8, 120, 12),
      () => measureHyperFormulaManySheetsBuildSample(8, 120, 12),
    ),
    runComparableScenario(
      "single-edit-chain",
      { downstreamCount: 2_000 },
      runtimeOptions,
      () => measureWorkPaperSingleChainEditSample(2_000),
      () => measureHyperFormulaSingleChainEditSample(2_000),
    ),
    runComparableScenario(
      "single-edit-fanout",
      { downstreamCount: 2_000 },
      runtimeOptions,
      () => measureWorkPaperSingleFanoutEditSample(2_000),
      () => measureHyperFormulaSingleFanoutEditSample(2_000),
    ),
    runComparableScenario(
      "single-formula-edit-recalc",
      { downstreamCount: 1_500 },
      runtimeOptions,
      () => measureWorkPaperFormulaEditSample(1_500),
      () => measureHyperFormulaFormulaEditSample(1_500),
    ),
    runComparableScenario(
      "batch-edit-single-column",
      { editCount: 500 },
      runtimeOptions,
      () => measureWorkPaperBatchSingleColumnEditSample(500),
      () => measureHyperFormulaBatchSingleColumnEditSample(500),
    ),
    runComparableScenario(
      "batch-edit-multi-column",
      { rowCount: 250, editsPerRow: 2 },
      runtimeOptions,
      () => measureWorkPaperBatchMultiColumnEditSample(250),
      () => measureHyperFormulaBatchMultiColumnEditSample(250),
    ),
    runComparableScenario(
      "range-read-dense",
      { cols: 24, rows: 240, materializedCells: 240 * 24 },
      runtimeOptions,
      () => measureWorkPaperRangeReadSample(240, 24),
      () => measureHyperFormulaRangeReadSample(240, 24),
    ),
    runComparableScenario(
      "lookup-no-column-index",
      { rowCount: 5_000, useColumnIndex: false },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, false),
      () => measureHyperFormulaLookupSample(5_000, false),
    ),
    runComparableScenario(
      "lookup-with-column-index",
      { rowCount: 5_000, useColumnIndex: true },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, true),
      () => measureHyperFormulaLookupSample(5_000, true),
    ),
    runComparableScenario(
      "lookup-approximate-sorted",
      { rowCount: 5_000 },
      runtimeOptions,
      () => measureWorkPaperApproximateLookupSample(5_000),
      () => measureHyperFormulaApproximateLookupSample(5_000),
    ),
    runComparableScenario(
      "lookup-text-exact",
      { rowCount: 5_000 },
      runtimeOptions,
      () => measureWorkPaperTextLookupSample(5_000),
      () => measureHyperFormulaTextLookupSample(5_000),
    ),
    runLeadershipScenario(
      "dynamic-array-filter",
      { rowCount: 750, formula: "=FILTER(A2:A751,A2:A751>B1)" },
      runtimeOptions,
      () => measureWorkPaperDynamicArraySample(750),
      {
        status: "unsupported",
        evidence: [
          "/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md",
          "/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts",
        ],
        reason: "HyperFormula 3.2.0 documents dynamic arrays as unsupported.",
      },
    ),
  ];
}

if (import.meta.url === `file://${process.argv[1]}`) {
  console.log(
    JSON.stringify(
      runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
        sampleCount: DEFAULT_COMPETITIVE_SAMPLE_COUNT,
        warmupCount: DEFAULT_COMPETITIVE_WARMUP_COUNT,
      }),
      null,
      2,
    ),
  );
}

function runComparableScenario(
  workload: ExpandedComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  runHyperFormulaSample: () => BenchmarkSample,
): ExpandedComparativeComparableResult {
  const workpaper = benchmarkSupportedEngine(runWorkPaperSample, options);
  const hyperformula = benchmarkSupportedEngine(runHyperFormulaSample, options);
  const workPaperVerification = JSON.stringify(workpaper.verification);
  const hyperFormulaVerification = JSON.stringify(hyperformula.verification);
  if (workPaperVerification !== hyperFormulaVerification) {
    throw new Error(
      `Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== HyperFormula ${hyperFormulaVerification}`,
    );
  }

  const fasterEngine =
    workpaper.elapsedMs.mean <= hyperformula.elapsedMs.mean ? "workpaper" : "hyperformula";
  const fasterMean =
    fasterEngine === "workpaper" ? workpaper.elapsedMs.mean : hyperformula.elapsedMs.mean;
  const slowerMean =
    fasterEngine === "workpaper" ? hyperformula.elapsedMs.mean : workpaper.elapsedMs.mean;

  return {
    workload,
    category: "directly-comparable",
    comparable: true,
    fixture,
    comparison: {
      fasterEngine,
      meanSpeedup: slowerMean / fasterMean,
      verificationEquivalent: true,
    },
    engines: {
      workpaper,
      hyperformula,
    },
  };
}

function runLeadershipScenario(
  workload: ExpandedComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  hyperformula: ComparativeUnsupportedEngineResult,
): ExpandedComparativeLeadershipResult {
  return {
    workload,
    category: "leadership",
    comparable: false,
    fixture,
    note: "This workload demonstrates capability leadership and is not an apples-to-apples speed comparison.",
    engines: {
      workpaper: benchmarkSupportedEngine(runWorkPaperSample, options),
      hyperformula,
    },
  };
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: Required<ComparativeBenchmarkSuiteOptions>,
): ComparativeMeasuredEngineResult {
  for (let warmup = 0; warmup < options.warmupCount; warmup += 1) {
    runSample();
  }

  const samples: BenchmarkSample[] = [];
  for (let sample = 0; sample < options.sampleCount; sample += 1) {
    samples.push(runSample());
  }

  const verificationStrings = new Set(samples.map((sample) => JSON.stringify(sample.verification)));
  if (verificationStrings.size !== 1) {
    throw new Error("Benchmark verification drifted across samples");
  }

  return {
    status: "supported",
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    verification: samples[0]?.verification ?? {},
  };
}

function measureWorkPaperDenseBuildSample(rows: number, cols: number): BenchmarkSample {
  const sheet = buildDenseLiteralSheet(rows, cols);
  return measureWorkPaperBuildFromSheets({ Bench: sheet }, (workbook) => {
    const sheetId = workbook.getSheetId("Bench")!;
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rows - 1, cols - 1)),
      ),
    };
  });
}

function measureHyperFormulaDenseBuildSample(rows: number, cols: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets(
    { Bench: toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols)) },
    (workbook) => {
      const sheetId = workbook.getSheetId("Bench")!;
      return {
        dimensions: workbook.getSheetDimensions(sheetId),
        terminalValue: normalizeHyperFormulaValue(
          workbook.getCellValue(address(sheetId, rows - 1, cols - 1)),
        ),
      };
    },
  );
}

function measureWorkPaperMixedBuildSample(rowCount: number): BenchmarkSample {
  const sheet = buildMixedContentSheet(rowCount);
  return measureWorkPaperBuildFromSheets({ Bench: sheet }, (workbook) => {
    const sheetId = workbook.getSheetId("Bench")!;
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalFormulaValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rowCount - 1, 5)),
      ),
    };
  });
}

function measureHyperFormulaMixedBuildSample(rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets(
    { Bench: toHyperFormulaSheet(buildMixedContentSheet(rowCount)) },
    (workbook) => {
      const sheetId = workbook.getSheetId("Bench")!;
      return {
        dimensions: workbook.getSheetDimensions(sheetId),
        terminalFormulaValue: normalizeHyperFormulaValue(
          workbook.getCellValue(address(sheetId, rowCount - 1, 5)),
        ),
      };
    },
  );
}

function measureWorkPaperManySheetsBuildSample(
  sheetCount: number,
  rows: number,
  cols: number,
): BenchmarkSample {
  const sheets = buildMultiSheetLiteralSheets(sheetCount, rows, cols);
  return measureWorkPaperBuildFromSheets(sheets, (workbook) => {
    const sheetId = workbook.getSheetId(`Sheet${sheetCount}`)!;
    return {
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, rows - 1, cols - 1)),
      ),
    };
  });
}

function measureHyperFormulaManySheetsBuildSample(
  sheetCount: number,
  rows: number,
  cols: number,
): BenchmarkSample {
  const sheets = Object.fromEntries(
    Object.entries(buildMultiSheetLiteralSheets(sheetCount, rows, cols)).map(
      ([sheetName, sheet]) => [sheetName, toHyperFormulaSheet(sheet)],
    ),
  );
  return measureHyperFormulaBuildFromSheets(sheets, (workbook) => {
    const sheetId = workbook.getSheetId(`Sheet${sheetCount}`)!;
    return {
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, rows - 1, cols - 1)),
      ),
    };
  });
}

function measureWorkPaperSingleChainEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: [buildFormulaChainRow(downstreamCount)] });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, 0, downstreamCount)),
      ),
    }),
  );
}

function measureHyperFormulaSingleChainEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaChainRow(downstreamCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, 0, downstreamCount)),
      ),
    }),
  );
}

function measureWorkPaperSingleFanoutEditSample(fanoutCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: [buildFormulaFanoutRow(fanoutCount)] });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, 0, fanoutCount)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

function measureHyperFormulaSingleFanoutEditSample(fanoutCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaFanoutRow(fanoutCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, 0, fanoutCount)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

function measureWorkPaperFormulaEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: [buildFormulaEditChainRow(downstreamCount)],
  });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 2), "=A1*B1"),
    () => ({
      editedFormula: workbook.getCellFormula(address(sheetId, 0, 2)) ?? null,
      terminalValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, 0, downstreamCount + 2)),
      ),
    }),
  );
}

function measureHyperFormulaFormulaEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaEditChainRow(downstreamCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 2), "=A1*B1"),
    () => ({
      editedFormula: workbook.getCellFormula(address(sheetId, 0, 2)) ?? null,
      terminalValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, 0, downstreamCount + 2)),
      ),
    }),
  );
}

function measureWorkPaperBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3);
        }
      }),
    () => ({
      sampleFormulaValue: normalizeWorkPaperValue(
        workbook.getCellValue(address(sheetId, editCount - 1, 1)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

function measureHyperFormulaBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3);
        }
      }),
    () => ({
      sampleFormulaValue: normalizeHyperFormulaValue(
        workbook.getCellValue(address(sheetId, editCount - 1, 1)),
      ),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  );
}

function measureWorkPaperBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildBatchMultiColumnRows(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3);
          workbook.setCellContents(address(sheetId, row, 1), row * 5);
        }
      }),
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

function measureHyperFormulaBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildBatchMultiColumnRows(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3);
          workbook.setCellContents(address(sheetId, row, 1), row * 5);
        }
      }),
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

function measureWorkPaperRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseLiteralSheet(rows, cols) });
  const sheetId = workbook.getSheetId("Bench")!;
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1);
  return measureMutationSample(
    workbook,
    () => workbook.getRangeValues(targetRange),
    (values) => {
      const lastRow = values.at(-1);
      return {
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeWorkPaperValue(lastRow?.at(-1)),
        topLeftValue: normalizeWorkPaperValue(values[0]?.[0]),
      };
    },
  );
}

function measureHyperFormulaRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1);
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.getRangeValues(targetRange),
    (values) => {
      const lastRow = values.at(-1);
      return {
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeHyperFormulaValue(lastRow?.at(-1)),
        topLeftValue: normalizeHyperFormulaValue(values[0]?.[0]),
      };
    },
  );
}

function measureWorkPaperLookupSample(rowCount: number, useColumnIndex: boolean): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(
    { Bench: buildLookupSheet(rowCount) },
    { useColumnIndex },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureHyperFormulaLookupSample(
  rowCount: number,
  useColumnIndex: boolean,
): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureWorkPaperApproximateLookupSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount - 0.5),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureHyperFormulaApproximateLookupSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount - 0.5),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureWorkPaperTextLookupSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildTextLookupSheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, "KEY-04999"),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureHyperFormulaTextLookupSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildTextLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  );
  const sheetId = workbook.getSheetId("Bench")!;
  const targetAddress = address(sheetId, 0, 3);
  const formulaAddress = address(sheetId, 0, 4);
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, "KEY-04999"),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  );
}

function measureWorkPaperDynamicArraySample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDynamicArraySheet(rowCount) });
  const sheetId = workbook.getSheetId("Bench")!;
  const thresholdAddress = address(sheetId, 0, 1);
  const spillAnchor = address(sheetId, 0, 2);
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(thresholdAddress, rowCount - 10),
    () => ({
      spillHeight: workbook.getSheetDimensions(sheetId).height,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillValue: normalizeWorkPaperValue(workbook.getCellValue(spillAnchor)),
    }),
  );
}

function measureWorkPaperBuildFromSheets(
  sheets: Record<string, ReturnType<typeof buildDenseLiteralSheet>>,
  verification: (workbook: WorkPaper) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const workbook = WorkPaper.buildFromSheets(sheets);
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  const result = verification(workbook);
  workbook.dispose();
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: result,
  };
}

function measureHyperFormulaBuildFromSheets(
  sheets: Record<string, HyperFormulaSheet>,
  verification: (workbook: HyperFormulaInstance) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const workbook = HyperFormula.buildFromSheets(sheets, {
    licenseKey: HYPERFORMULA_LICENSE_KEY,
  });
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  const result = verification(workbook);
  workbook.destroy();
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: result,
  };
}

function measureMutationSample<Result>(
  workbook: WorkPaper,
  execute: () => Result,
  verification: (result: Result) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const result = execute();
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  const resolvedVerification = verification(result);
  workbook.dispose();
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: resolvedVerification,
  };
}

function measureHyperFormulaMutationSample<Result>(
  workbook: HyperFormulaInstance,
  execute: () => Result,
  verification: (result: Result) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory();
  const started = performance.now();
  const result = execute();
  const elapsedMs = performance.now() - started;
  const memoryAfter = sampleMemory();
  const resolvedVerification = verification(result);
  workbook.destroy();
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: resolvedVerification,
  };
}

function summarizeMemory(samples: readonly MemoryMeasurement[]): ComparativeMemorySummary {
  return {
    rssBytes: summarizeNumbers(samples.map((sample) => sample.delta.rssBytes)),
    heapUsedBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapUsedBytes)),
    heapTotalBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapTotalBytes)),
    externalBytes: summarizeNumbers(samples.map((sample) => sample.delta.externalBytes)),
    arrayBuffersBytes: summarizeNumbers(samples.map((sample) => sample.delta.arrayBuffersBytes)),
  };
}

function normalizeWorkPaperValue(
  value: unknown,
): boolean | number | string | null | { error: unknown } {
  if (!isProtocolValueLike(value)) {
    return null;
  }

  switch (value.tag) {
    case ValueTag.Empty:
      return null;
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null;
    case ValueTag.Error:
      return { error: value.code ?? "ERROR" };
    default:
      return { error: `UNKNOWN_TAG_${String(value.tag)}` };
  }
}

function normalizeHyperFormulaValue(
  value: unknown,
): boolean | number | string | null | { error: unknown } {
  if (
    value === null ||
    typeof value === "boolean" ||
    typeof value === "number" ||
    typeof value === "string"
  ) {
    return value;
  }
  if (isHyperFormulaErrorLike(value)) {
    return { error: value.value };
  }
  return { error: "UNKNOWN_VALUE" };
}

function resolveSuiteOptions(
  options: ComparativeBenchmarkSuiteOptions,
): Required<ComparativeBenchmarkSuiteOptions> {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  };
}

function isProtocolValueLike(
  value: unknown,
): value is { code?: unknown; tag: ValueTag; value?: boolean | number | string } {
  if (!value || typeof value !== "object") {
    return false;
  }
  const tag = Reflect.get(value, "tag");
  return (
    tag === ValueTag.Empty ||
    tag === ValueTag.Number ||
    tag === ValueTag.Boolean ||
    tag === ValueTag.String ||
    tag === ValueTag.Error
  );
}

function isHyperFormulaErrorLike(value: unknown): value is { value: unknown } {
  return value !== null && typeof value === "object" && "value" in value;
}

function toHyperFormulaSheet(sheet: ReadonlyArray<ReadonlyArray<unknown>>): HyperFormulaSheet {
  return sheet.map((row) => row.map((cell) => toHyperFormulaCell(cell)));
}

function toHyperFormulaCell(cell: unknown): HyperFormulaRawCellContent {
  if (
    cell === null ||
    typeof cell === "boolean" ||
    typeof cell === "number" ||
    typeof cell === "string"
  ) {
    return cell;
  }
  throw new Error(`Unsupported HyperFormula benchmark cell type: ${typeof cell}`);
}
