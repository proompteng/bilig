import { performance } from 'node:perf_hooks'
import xlsxCalc from 'xlsx-calc'
import { WorkPaper, type WorkPaperSheet } from '../../headless/src/work-paper.js'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  type ComparativeBenchmarkSuiteOptions,
  type ComparativeMeasuredEngineResult,
  type ComparativeMemorySummary,
} from './benchmark-workpaper-vs-hyperformula.js'
import {
  measureMutationSample,
  normalizeWorkPaperValue,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'
import { summarizeNumbers } from './stats.js'
export type {
  WorkPaperXlsxCalcBenchmarkReport,
  WorkPaperXlsxCalcBenchmarkResult,
  WorkPaperXlsxCalcComparison,
  WorkPaperXlsxCalcFixture,
  WorkPaperXlsxCalcScorecard,
  WorkPaperXlsxCalcWorkload,
  WorkPaperXlsxCalcWorkloadFamily,
  XlsxCalcEditableValue,
} from './benchmark-workpaper-vs-xlsx-calc-types.js'
import type {
  WorkPaperXlsxCalcBenchmarkReport,
  WorkPaperXlsxCalcBenchmarkResult,
  WorkPaperXlsxCalcFixture,
  WorkPaperXlsxCalcWorkload,
  XlsxCalcEditableValue,
} from './benchmark-workpaper-vs-xlsx-calc-types.js'

interface XlsxCell {
  f?: string
  t?: 'b' | 'e' | 'n' | 's'
  v?: boolean | number | string
  w?: string
}

interface XlsxSheet {
  [address: string]: string | XlsxCell | undefined
  '!ref': string
}

interface XlsxWorkbook {
  SheetNames: string[]
  Sheets: Record<string, XlsxSheet>
}

interface WorkPaperXlsxCalcScenario {
  readonly fixture: WorkPaperXlsxCalcFixture
  readonly buildWorkPaperSheets: () => Record<string, WorkPaperSheet>
  readonly buildXlsxWorkbook: () => XlsxWorkbook
}

interface ResolvedBenchmarkSuiteOptions {
  readonly sampleCount: number
  readonly warmupCount: number
}

export const WORKPAPER_XLSX_CALC_WORKLOADS = [
  'xlsx-calc-large-sum-recalc',
  'xlsx-calc-2d-sum-recalc',
  'xlsx-calc-overlapping-sum-recalc',
  'xlsx-calc-exact-match-recalc',
  'xlsx-calc-approximate-match-recalc',
  'xlsx-calc-formula-chain-recalc',
  'xlsx-calc-scalar-fanout-recalc',
  'xlsx-calc-deep-chain-recalc',
  'xlsx-calc-cross-sheet-sum-recalc',
  'xlsx-calc-cross-sheet-chain-recalc',
  'xlsx-calc-cross-sheet-scalar-fanout-recalc',
  'xlsx-calc-index-match-exact-text-recalc',
  'xlsx-calc-vlookup-text-exact-recalc',
  'xlsx-calc-vlookup-approximate-duplicates-recalc',
  'xlsx-calc-hlookup-exact-numeric-recalc',
  'xlsx-calc-range-stats-recalc',
] as const satisfies readonly WorkPaperXlsxCalcWorkload[]

export function runWorkPaperVsXlsxCalcBenchmarkSuite(options: ComparativeBenchmarkSuiteOptions = {}): WorkPaperXlsxCalcBenchmarkResult[] {
  const resolvedOptions = resolveSuiteOptions(options)
  return WORKPAPER_XLSX_CALC_WORKLOADS.map((workload) => runXlsxCalcScenario(workload, xlsxCalcScenario(workload), resolvedOptions))
}

export function buildWorkPaperVsXlsxCalcBenchmarkReport(
  results: readonly WorkPaperXlsxCalcBenchmarkResult[],
): WorkPaperXlsxCalcBenchmarkReport {
  if (results.length === 0) {
    throw new Error('Cannot build a WorkPaper vs xlsx-calc scorecard without benchmark results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToXlsxCalcMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToXlsxCalcP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToXlsxCalcMeanRatio < 1 && result.comparison.workpaperToXlsxCalcP95Ratio < 1,
  ).length

  return {
    suite: 'workpaper-vs-xlsx-calc',
    scorecard: {
      comparableWorkloadCount: results.length,
      coverageNote:
        'xlsx-calc is covered as a workbook-wide-limited SheetJS-style recalculation engine for formulas it supports. This lane is intentionally limited to equivalent aggregate, lookup, formula-chain, fanout, range-stat, overlapping-range, and cross-sheet recalculation workloads.',
      coverageTier: 'workbook-wide-limited',
      directionalMeanRatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToXlsxCalcMeanRatio)),
      directionalP95RatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToXlsxCalcP95Ratio)),
      meanAndP95WinCount,
      meanWinCount,
      p95WinCount,
      workloadFamilies: orderedUnique(results.map((result) => result.fixture.family)),
      worstMeanRatioWorkload: maxComparableRatioWorkload(results, 'workpaperToXlsxCalcMeanRatio'),
      worstP95RatioWorkload: maxComparableRatioWorkload(results, 'workpaperToXlsxCalcP95Ratio'),
      worstWorkpaperToXlsxCalcMeanRatio: maxComparableRatio(results, 'workpaperToXlsxCalcMeanRatio'),
      worstWorkpaperToXlsxCalcP95Ratio: maxComparableRatio(results, 'workpaperToXlsxCalcP95Ratio'),
      xlsxCalcMeanWinCount: results.length - meanWinCount,
      xlsxCalcP95WinCount: results.length - p95WinCount,
    },
    results,
  }
}

function runXlsxCalcScenario(
  workload: WorkPaperXlsxCalcWorkload,
  scenario: WorkPaperXlsxCalcScenario,
  options: ResolvedBenchmarkSuiteOptions,
): WorkPaperXlsxCalcBenchmarkResult {
  const workpaper = benchmarkSupportedEngine(() => measureWorkPaperRecalcSample(scenario), options)
  const xlsxCalcResult = benchmarkSupportedEngine(() => measureXlsxCalcRecalcSample(scenario), options)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const xlsxCalcVerification = JSON.stringify(xlsxCalcResult.verification)
  if (workPaperVerification !== xlsxCalcVerification) {
    throw new Error(`Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== xlsx-calc ${xlsxCalcVerification}`)
  }

  const fasterEngine = workpaper.elapsedMs.mean <= xlsxCalcResult.elapsedMs.mean ? 'workpaper' : 'xlsx-calc'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : xlsxCalcResult.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? xlsxCalcResult.elapsedMs.mean : workpaper.elapsedMs.mean

  return {
    workload,
    category: 'workbook-wide-limited',
    comparable: true,
    fixture: scenario.fixture,
    comparison: {
      confidenceIntervalOverlaps:
        workpaper.elapsedMs.confidence95.low <= xlsxCalcResult.elapsedMs.confidence95.high &&
        xlsxCalcResult.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high,
      fasterEngine,
      maxRelativeNoise: Math.max(workpaper.elapsedMs.relativeStandardDeviation, xlsxCalcResult.elapsedMs.relativeStandardDeviation),
      meanSpeedup: slowerMean / fasterMean,
      verificationEquivalent: true,
      workpaperToXlsxCalcMeanRatio: workpaper.elapsedMs.mean / xlsxCalcResult.elapsedMs.mean,
      workpaperToXlsxCalcMedianRatio: workpaper.elapsedMs.median / xlsxCalcResult.elapsedMs.median,
      workpaperToXlsxCalcP95Ratio: workpaper.elapsedMs.p95 / xlsxCalcResult.elapsedMs.p95,
    },
    engines: {
      workpaper,
      xlsxCalc: xlsxCalcResult,
    },
  }
}

function measureWorkPaperRecalcSample(scenario: WorkPaperXlsxCalcScenario): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(scenario.buildWorkPaperSheets())
  const editSheetId = workbook.getSheetId(scenario.fixture.edit.sheetName)
  const resultSheetId = workbook.getSheetId(scenario.fixture.result.sheetName)
  if (editSheetId === undefined) {
    workbook.dispose()
    throw new Error(`WorkPaper xlsx-calc benchmark fixture did not create ${scenario.fixture.edit.sheetName}`)
  }
  if (resultSheetId === undefined) {
    workbook.dispose()
    throw new Error(`WorkPaper xlsx-calc benchmark fixture did not create ${scenario.fixture.result.sheetName}`)
  }
  return measureMutationSample(
    workbook,
    () =>
      workbook.setCellContents(
        { sheet: editSheetId, row: scenario.fixture.edit.row, col: scenario.fixture.edit.col },
        scenario.fixture.edit.value,
      ),
    () => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: resultSheetId, row: scenario.fixture.result.row, col: scenario.fixture.result.col }),
        ),
      ),
    }),
  )
}

function measureXlsxCalcRecalcSample(scenario: WorkPaperXlsxCalcScenario): BenchmarkSample {
  const workbook = scenario.buildXlsxWorkbook()
  xlsxCalc(workbook)
  const editSheet = workbook.Sheets[scenario.fixture.edit.sheetName]
  const resultSheet = workbook.Sheets[scenario.fixture.result.sheetName]
  if (!editSheet) {
    throw new Error(`xlsx-calc benchmark fixture did not create ${scenario.fixture.edit.sheetName}`)
  }
  if (!resultSheet) {
    throw new Error(`xlsx-calc benchmark fixture did not create ${scenario.fixture.result.sheetName}`)
  }

  const memoryBefore = sampleMemory()
  const started = performance.now()
  setXlsxCell(editSheet, scenario.fixture.edit.address, scenario.fixture.edit.value)
  xlsxCalc(workbook)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: {
      value: normalizeBenchmarkValue(normalizeXlsxCell(resultSheet[scenario.fixture.result.address])),
    },
  }
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: ResolvedBenchmarkSuiteOptions,
): ComparativeMeasuredEngineResult {
  for (let warmup = 0; warmup < options.warmupCount; warmup += 1) {
    runSample()
  }

  const samples: BenchmarkSample[] = []
  for (let sample = 0; sample < options.sampleCount; sample += 1) {
    samples.push(runSample())
  }

  const verificationStrings = new Set(samples.map((sample) => JSON.stringify(sample.verification)))
  if (verificationStrings.size !== 1) {
    throw new Error('Benchmark verification drifted across samples')
  }

  return {
    status: 'supported',
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    verification: samples[0]?.verification ?? {},
  }
}

function summarizeMemory(samples: readonly MemoryMeasurement[]): ComparativeMemorySummary {
  return {
    rssBytes: summarizeNumbers(samples.map((sample) => sample.delta.rssBytes)),
    heapUsedBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapUsedBytes)),
    heapTotalBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapTotalBytes)),
    externalBytes: summarizeNumbers(samples.map((sample) => sample.delta.externalBytes)),
    arrayBuffersBytes: summarizeNumbers(samples.map((sample) => sample.delta.arrayBuffersBytes)),
  }
}

function xlsxCalcScenario(workload: WorkPaperXlsxCalcWorkload): WorkPaperXlsxCalcScenario {
  switch (workload) {
    case 'xlsx-calc-large-sum-recalc':
      return aggregateScenario(5_000)
    case 'xlsx-calc-2d-sum-recalc':
      return twoDimensionalAggregateScenario(1_000, 8)
    case 'xlsx-calc-overlapping-sum-recalc':
      return overlappingAggregateScenario(1_500)
    case 'xlsx-calc-exact-match-recalc':
      return vlookupScenario({
        editValue: 6_200,
        family: 'lookup-exact',
        formula: '=VLOOKUP(D1,A1:B5000,2,FALSE)',
        initialLookupValue: 2_400,
        rowCount: 5_000,
      })
    case 'xlsx-calc-approximate-match-recalc':
      return vlookupScenario({
        editValue: 6_201,
        family: 'lookup-approximate',
        formula: '=VLOOKUP(D1,A1:B5000,2,TRUE)',
        initialLookupValue: 2_401,
        rowCount: 5_000,
      })
    case 'xlsx-calc-formula-chain-recalc':
      return formulaChainScenario(2_000)
    case 'xlsx-calc-scalar-fanout-recalc':
      return scalarFanoutScenario(1_000)
    case 'xlsx-calc-deep-chain-recalc':
      return deepChainScenario(1_000)
    case 'xlsx-calc-cross-sheet-sum-recalc':
      return crossSheetAggregateScenario(3_000)
    case 'xlsx-calc-cross-sheet-chain-recalc':
      return crossSheetFormulaChainScenario(2_000)
    case 'xlsx-calc-cross-sheet-scalar-fanout-recalc':
      return crossSheetScalarFanoutScenario(1_000)
    case 'xlsx-calc-index-match-exact-text-recalc':
      return indexMatchExactTextScenario(5_000)
    case 'xlsx-calc-vlookup-text-exact-recalc':
      return vlookupTextExactScenario(5_000)
    case 'xlsx-calc-vlookup-approximate-duplicates-recalc':
      return vlookupApproximateDuplicatesScenario(5_000)
    case 'xlsx-calc-hlookup-exact-numeric-recalc':
      return hlookupExactNumericScenario(3_000)
    case 'xlsx-calc-range-stats-recalc':
      return rangeStatsScenario(1_000)
  }
}

function aggregateScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=SUM(A1:A${String(rowCount)})`
  const fixture = {
    edit: { address: 'A2500', col: 0, row: 2_499, sheetName: 'Sheet1', value: 10_000 },
    family: 'aggregate',
    formula,
    result: { address: 'C1', col: 2, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet = numberColumnSheet(rowCount)
      sheet[0] = [1, null, formula]
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:C${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        setXlsxCell(sheet, address(row, 0), row + 1)
      }
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function twoDimensionalAggregateScenario(rowCount: number, colCount: number): WorkPaperXlsxCalcScenario {
  const lastCol = columnName(colCount - 1)
  const formula = `=SUM(A1:${lastCol}${String(rowCount)})`
  const resultAddress = `${columnName(colCount + 1)}1`
  const fixture = {
    edit: { address: 'D500', col: 3, row: 499, sheetName: 'Sheet1', value: 10_000 },
    family: 'aggregate-2d',
    formula,
    result: { address: resultAddress, col: colCount + 1, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowValues: Array<number | string | null> = []
        for (let col = 0; col < colCount; col += 1) {
          rowValues.push((row + 1) * (col + 1))
        }
        sheet.push(rowValues)
      }
      sheet[0]![colCount + 1] = formula
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:${resultAddress.slice(0, -1)}${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        for (let col = 0; col < colCount; col += 1) {
          setXlsxCell(sheet, address(row, col), (row + 1) * (col + 1))
        }
      }
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function overlappingAggregateScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=B${String(rowCount)}`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 99 },
    family: 'overlapping-aggregate',
    formula,
    result: { address: 'C1', col: 2, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([rowNumber, `=SUM(A1:A${String(rowNumber)})`])
      }
      sheet[0]![2] = formula
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:C${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(sheet, address(row, 0), rowNumber)
        setXlsxFormulaCell(sheet, address(row, 1), `=SUM(A1:A${String(rowNumber)})`)
      }
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function vlookupScenario(args: {
  readonly editValue: number
  readonly family: 'lookup-approximate' | 'lookup-exact'
  readonly formula: string
  readonly initialLookupValue: number
  readonly rowCount: number
}): WorkPaperXlsxCalcScenario {
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: args.editValue },
    family: args.family,
    formula: args.formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount: args.rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < args.rowCount; row += 1) {
        sheet.push([(row + 1) * 2, (row + 1) * 10])
      }
      sheet[0] = [2, 10, null, args.initialLookupValue, args.formula]
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:E${String(args.rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < args.rowCount; row += 1) {
        setXlsxCell(sheet, address(row, 0), (row + 1) * 2)
        setXlsxCell(sheet, address(row, 1), (row + 1) * 10)
      }
      setXlsxCell(sheet, 'D1', args.initialLookupValue)
      setXlsxFormulaCell(sheet, 'E1', args.formula)
      return workbook
    },
  }
}

function formulaChainScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = '=B1000+B999'
  const fixture = {
    edit: { address: 'A1000', col: 0, row: 999, sheetName: 'Sheet1', value: 5_000 },
    family: 'formula-chain',
    formula,
    result: { address: 'C1', col: 2, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([rowNumber, `=A${String(rowNumber)}*2`])
      }
      sheet[0] = [1, '=A1*2', formula]
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:C${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(sheet, address(row, 0), rowNumber)
        setXlsxFormulaCell(sheet, address(row, 1), `=A${String(rowNumber)}*2`)
      }
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function scalarFanoutScenario(formulaCount: number): WorkPaperXlsxCalcScenario {
  const resultCol = formulaCount
  const resultAddress = `${columnName(resultCol)}1`
  const formula = `=$A$1+${String(formulaCount)}`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 5_000 },
    family: 'formula-fanout',
    formula,
    result: { address: resultAddress, col: resultCol, row: 0, sheetName: 'Sheet1' },
    rowCount: formulaCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const row: Array<number | string | null> = [1]
      for (let index = 1; index <= formulaCount; index += 1) {
        row[index] = `=$A$1+${String(index)}`
      }
      return { Sheet1: [row] }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:${resultAddress}`)
      const sheet = workbook.Sheets['Sheet1']!
      setXlsxCell(sheet, 'A1', 1)
      for (let index = 1; index <= formulaCount; index += 1) {
        setXlsxFormulaCell(sheet, `${columnName(index)}1`, `=$A$1+${String(index)}`)
      }
      return workbook
    },
  }
}

function deepChainScenario(chainLength: number): WorkPaperXlsxCalcScenario {
  const resultCol = chainLength
  const resultAddress = `${columnName(resultCol)}1`
  const formula = `=${columnName(resultCol - 1)}1+1`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Sheet1', value: 5_000 },
    family: 'formula-chain',
    formula,
    result: { address: resultAddress, col: resultCol, row: 0, sheetName: 'Sheet1' },
    rowCount: chainLength,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const row: Array<number | string | null> = [1]
      for (let index = 1; index <= chainLength; index += 1) {
        row[index] = `=${columnName(index - 1)}1+1`
      }
      return { Sheet1: [row] }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:${resultAddress}`)
      const sheet = workbook.Sheets['Sheet1']!
      setXlsxCell(sheet, 'A1', 1)
      for (let index = 1; index <= chainLength; index += 1) {
        setXlsxFormulaCell(sheet, `${columnName(index)}1`, `=${columnName(index - 1)}1+1`)
      }
      return workbook
    },
  }
}

function crossSheetAggregateScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=SUM(Data!A1:A${String(rowCount)})`
  const fixture = {
    edit: { address: 'A1500', col: 0, row: 1_499, sheetName: 'Data', value: 10_000 },
    family: 'cross-sheet',
    formula,
    result: { address: 'A1', col: 0, row: 0, sheetName: 'Summary' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => ({
      Data: numberColumnSheet(rowCount),
      Summary: [[formula]],
    }),
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbookWithSheets({
        Data: `A1:A${String(rowCount)}`,
        Summary: 'A1:A1',
      })
      const data = workbook.Sheets['Data']!
      const summary = workbook.Sheets['Summary']!
      for (let row = 0; row < rowCount; row += 1) {
        setXlsxCell(data, address(row, 0), row + 1)
      }
      setXlsxFormulaCell(summary, fixture.result.address, formula)
      return workbook
    },
  }
}

function crossSheetFormulaChainScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = '=Data!B1000+Data!B999'
  const fixture = {
    edit: { address: 'A1000', col: 0, row: 999, sheetName: 'Data', value: 5_000 },
    family: 'cross-sheet',
    formula,
    result: { address: 'A1', col: 0, row: 0, sheetName: 'Summary' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const data: Array<Array<number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        data.push([rowNumber, `=A${String(rowNumber)}*2`])
      }
      return {
        Data: data,
        Summary: [[formula]],
      }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbookWithSheets({
        Data: `A1:B${String(rowCount)}`,
        Summary: 'A1:A1',
      })
      const data = workbook.Sheets['Data']!
      const summary = workbook.Sheets['Summary']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(data, address(row, 0), rowNumber)
        setXlsxFormulaCell(data, address(row, 1), `=A${String(rowNumber)}*2`)
      }
      setXlsxFormulaCell(summary, fixture.result.address, formula)
      return workbook
    },
  }
}

function crossSheetScalarFanoutScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=Data!$A$1+Data!B${String(rowCount)}`
  const fixture = {
    edit: { address: 'A1', col: 0, row: 0, sheetName: 'Data', value: 7 },
    family: 'cross-sheet',
    formula,
    result: { address: `A${String(rowCount)}`, col: 0, row: rowCount - 1, sheetName: 'Summary' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const data: Array<Array<number | string | null>> = []
      const summary: Array<Array<number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        data.push([3, rowNumber * 10])
        summary.push([`=Data!$A$1+Data!B${String(rowNumber)}`])
      }
      return { Data: data, Summary: summary }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbookWithSheets({
        Data: `A1:B${String(rowCount)}`,
        Summary: `A1:A${String(rowCount)}`,
      })
      const data = workbook.Sheets['Data']!
      const summary = workbook.Sheets['Summary']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(data, address(row, 0), 3)
        setXlsxCell(data, address(row, 1), rowNumber * 10)
        setXlsxFormulaCell(summary, address(row, 0), `=Data!$A$1+Data!B${String(rowNumber)}`)
      }
      return workbook
    },
  }
}

function indexMatchExactTextScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=INDEX(B1:B${String(rowCount)},MATCH(D1,A1:A${String(rowCount)},0),1)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: textLookupKey(3_100) },
    family: 'lookup-exact',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return textLookupScenario({ formula, initialLookupValue: textLookupKey(2_400), rowCount, fixture })
}

function vlookupTextExactScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,FALSE)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: textLookupKey(3_100) },
    family: 'lookup-exact',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return textLookupScenario({ formula, initialLookupValue: textLookupKey(2_400), rowCount, fixture })
}

function textLookupScenario(args: {
  readonly fixture: WorkPaperXlsxCalcFixture
  readonly formula: string
  readonly initialLookupValue: string
  readonly rowCount: number
}): WorkPaperXlsxCalcScenario {
  return {
    fixture: args.fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < args.rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([textLookupKey(rowNumber), rowNumber * 10])
      }
      sheet[0] = [textLookupKey(1), 10, null, args.initialLookupValue, args.formula]
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:E${String(args.rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < args.rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(sheet, address(row, 0), textLookupKey(rowNumber))
        setXlsxCell(sheet, address(row, 1), rowNumber * 10)
      }
      setXlsxCell(sheet, 'D1', args.initialLookupValue)
      setXlsxFormulaCell(sheet, 'E1', args.formula)
      return workbook
    },
  }
}

function vlookupApproximateDuplicatesScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=VLOOKUP(D1,A1:B${String(rowCount)},2,TRUE)`
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1', value: 1_550.5 },
    family: 'lookup-approximate',
    formula,
    result: { address: 'E1', col: 4, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([Math.ceil(rowNumber / 2), rowNumber])
      }
      sheet[0] = [1, 1, null, 1_200.5, formula]
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:E${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(sheet, address(row, 0), Math.ceil(rowNumber / 2))
        setXlsxCell(sheet, address(row, 1), rowNumber)
      }
      setXlsxCell(sheet, 'D1', 1_200.5)
      setXlsxFormulaCell(sheet, 'E1', formula)
      return workbook
    },
  }
}

function hlookupExactNumericScenario(colCount: number): WorkPaperXlsxCalcScenario {
  const lastCol = columnName(colCount - 1)
  const formula = `=HLOOKUP(A4,A1:${lastCol}2,2,FALSE)`
  const fixture = {
    edit: { address: 'A4', col: 0, row: 3, sheetName: 'Sheet1', value: 620 },
    family: 'lookup-exact',
    formula,
    result: { address: 'B4', col: 1, row: 3, sheetName: 'Sheet1' },
    rowCount: colCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const keys: Array<number | string | null> = []
      const values: Array<number | string | null> = []
      for (let col = 0; col < colCount; col += 1) {
        keys[col] = col + 1
        values[col] = (col + 1) * 5
      }
      return { Sheet1: [keys, values, [], [200, formula]] }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:${lastCol}4`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let col = 0; col < colCount; col += 1) {
        setXlsxCell(sheet, address(0, col), col + 1)
        setXlsxCell(sheet, address(1, col), (col + 1) * 5)
      }
      setXlsxCell(sheet, 'A4', 200)
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function rangeStatsScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=AVERAGE(A1:A${String(rowCount)})+MAX(B1:B${String(rowCount)})-MIN(C1:C${String(rowCount)})`
  const fixture = {
    edit: { address: 'B1000', col: 1, row: 999, sheetName: 'Sheet1', value: 9_999 },
    family: 'range-stats',
    formula,
    result: { address: 'D1', col: 3, row: 0, sheetName: 'Sheet1' },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheets: () => {
      const sheet: Array<Array<number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([rowNumber, rowNumber, rowNumber])
      }
      sheet[0]![3] = formula
      return { Sheet1: sheet }
    },
    buildXlsxWorkbook: () => {
      const workbook = createXlsxWorkbook(`A1:D${String(rowCount)}`)
      const sheet = workbook.Sheets['Sheet1']!
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        setXlsxCell(sheet, address(row, 0), rowNumber)
        setXlsxCell(sheet, address(row, 1), rowNumber)
        setXlsxCell(sheet, address(row, 2), rowNumber)
      }
      setXlsxFormulaCell(sheet, fixture.result.address, formula)
      return workbook
    },
  }
}

function textLookupKey(rowNumber: number): string {
  return `KEY-${String(rowNumber).padStart(5, '0')}`
}

function numberColumnSheet(rowCount: number): Array<Array<boolean | number | string | null>> {
  const sheet: Array<Array<boolean | number | string | null>> = []
  for (let row = 0; row < rowCount; row += 1) {
    sheet.push([row + 1])
  }
  return sheet
}

function createXlsxWorkbook(ref: string): XlsxWorkbook {
  return createXlsxWorkbookWithSheets({ Sheet1: ref })
}

function createXlsxWorkbookWithSheets(refsBySheetName: Record<string, string>): XlsxWorkbook {
  const sheets: Record<string, XlsxSheet> = {}
  for (const [sheetName, ref] of Object.entries(refsBySheetName)) {
    sheets[sheetName] = { '!ref': ref }
  }
  return {
    SheetNames: Object.keys(refsBySheetName),
    Sheets: sheets,
  }
}

function setXlsxCell(sheet: XlsxSheet, cellAddress: string, value: XlsxCalcEditableValue): void {
  if (value === null) {
    delete sheet[cellAddress]
    return
  }
  if (typeof value === 'number') {
    sheet[cellAddress] = { t: 'n', v: value }
    return
  }
  if (typeof value === 'boolean') {
    sheet[cellAddress] = { t: 'b', v: value }
    return
  }
  sheet[cellAddress] = { t: 's', v: value }
}

function setXlsxFormulaCell(sheet: XlsxSheet, cellAddress: string, formula: string): void {
  sheet[cellAddress] = { f: formula.startsWith('=') ? formula.slice(1) : formula }
}

function normalizeXlsxCell(cell: string | XlsxCell | undefined): boolean | number | string | null | { error: unknown } {
  if (cell === undefined || typeof cell === 'string') {
    return cell ?? null
  }
  if (cell.t === 'e') {
    return { error: cell.w ?? cell.v ?? 'ERROR' }
  }
  return cell.v ?? null
}

function normalizeBenchmarkValue(
  value: boolean | number | string | null | { error: unknown },
): boolean | number | string | null | { error: string } {
  if (typeof value === 'number') {
    return Number(value.toPrecision(12))
  }
  if (isErrorRecord(value)) {
    return { error: String(value.error) }
  }
  return value
}

function isErrorRecord(value: unknown): value is { error: unknown } {
  return typeof value === 'object' && value !== null && 'error' in value
}

function address(row: number, col: number): string {
  return `${columnName(col)}${String(row + 1)}`
}

function columnName(col: number): string {
  let value = col + 1
  let name = ''
  while (value > 0) {
    const modulo = (value - 1) % 26
    name = String.fromCharCode(65 + modulo) + name
    value = Math.floor((value - modulo) / 26)
  }
  return name
}

function geometricMean(values: readonly number[]): number {
  if (values.length === 0) {
    return Number.NaN
  }
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}

function maxComparableRatio(
  results: readonly WorkPaperXlsxCalcBenchmarkResult[],
  ratioKey: 'workpaperToXlsxCalcMeanRatio' | 'workpaperToXlsxCalcP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly WorkPaperXlsxCalcBenchmarkResult[],
  ratioKey: 'workpaperToXlsxCalcMeanRatio' | 'workpaperToXlsxCalcP95Ratio',
): WorkPaperXlsxCalcWorkload {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique<T extends string>(values: readonly T[]): T[] {
  return [...new Set(values)]
}

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): ResolvedBenchmarkSuiteOptions {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
}
