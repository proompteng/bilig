import { performance } from 'node:perf_hooks'
import { WorkPaper, type WorkPaperSheet } from '@bilig/headless'
import xlsxCalc from 'xlsx-calc'
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

export type WorkPaperXlsxCalcWorkload =
  | 'xlsx-calc-large-sum-recalc'
  | 'xlsx-calc-exact-match-recalc'
  | 'xlsx-calc-approximate-match-recalc'
  | 'xlsx-calc-formula-chain-recalc'

export type WorkPaperXlsxCalcWorkloadFamily = 'aggregate' | 'lookup-exact' | 'lookup-approximate' | 'formula-chain'

export type XlsxCalcEditableValue = boolean | number | string | null

export interface WorkPaperXlsxCalcFixture {
  readonly edit: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly value: XlsxCalcEditableValue
  }
  readonly family: WorkPaperXlsxCalcWorkloadFamily
  readonly formula: string
  readonly result: {
    readonly address: string
    readonly col: number
    readonly row: number
  }
  readonly rowCount: number
}

export interface WorkPaperXlsxCalcComparison {
  readonly confidenceIntervalOverlaps: boolean
  readonly fasterEngine: 'workpaper' | 'xlsx-calc'
  readonly maxRelativeNoise: number
  readonly meanSpeedup: number
  readonly verificationEquivalent: true
  readonly workpaperToXlsxCalcMeanRatio: number
  readonly workpaperToXlsxCalcMedianRatio: number
  readonly workpaperToXlsxCalcP95Ratio: number
}

export interface WorkPaperXlsxCalcBenchmarkResult {
  readonly workload: WorkPaperXlsxCalcWorkload
  readonly category: 'workbook-wide-limited'
  readonly comparable: true
  readonly fixture: WorkPaperXlsxCalcFixture
  readonly comparison: WorkPaperXlsxCalcComparison
  readonly engines: {
    readonly workpaper: ComparativeMeasuredEngineResult
    readonly xlsxCalc: ComparativeMeasuredEngineResult
  }
}

export interface WorkPaperXlsxCalcScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: 'workbook-wide'
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly workloadFamilies: readonly WorkPaperXlsxCalcWorkloadFamily[]
  readonly worstMeanRatioWorkload: WorkPaperXlsxCalcWorkload
  readonly worstP95RatioWorkload: WorkPaperXlsxCalcWorkload
  readonly worstWorkpaperToXlsxCalcMeanRatio: number
  readonly worstWorkpaperToXlsxCalcP95Ratio: number
  readonly xlsxCalcMeanWinCount: number
  readonly xlsxCalcP95WinCount: number
}

export interface WorkPaperXlsxCalcBenchmarkReport {
  readonly suite: 'workpaper-vs-xlsx-calc'
  readonly scorecard: WorkPaperXlsxCalcScorecard
  readonly results: readonly WorkPaperXlsxCalcBenchmarkResult[]
}

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
  readonly buildWorkPaperSheet: () => WorkPaperSheet
  readonly buildXlsxWorkbook: () => XlsxWorkbook
}

interface ResolvedBenchmarkSuiteOptions {
  readonly sampleCount: number
  readonly warmupCount: number
}

export const WORKPAPER_XLSX_CALC_WORKLOADS = [
  'xlsx-calc-large-sum-recalc',
  'xlsx-calc-exact-match-recalc',
  'xlsx-calc-approximate-match-recalc',
  'xlsx-calc-formula-chain-recalc',
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
        'xlsx-calc is covered as a workbook-wide SheetJS-style recalculation engine for formulas it supports. This lane is intentionally limited to equivalent aggregate, VLOOKUP, and formula-chain recalculation workloads.',
      coverageTier: 'workbook-wide',
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
  const workbook = WorkPaper.buildFromSheets({ Sheet1: scenario.buildWorkPaperSheet() })
  const sheetId = workbook.getSheetId('Sheet1')
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error('WorkPaper xlsx-calc benchmark fixture did not create Sheet1')
  }
  return measureMutationSample(
    workbook,
    () =>
      workbook.setCellContents(
        { sheet: sheetId, row: scenario.fixture.edit.row, col: scenario.fixture.edit.col },
        scenario.fixture.edit.value,
      ),
    () => ({
      value: normalizeBenchmarkValue(
        normalizeWorkPaperValue(
          workbook.getCellValue({ sheet: sheetId, row: scenario.fixture.result.row, col: scenario.fixture.result.col }),
        ),
      ),
    }),
  )
}

function measureXlsxCalcRecalcSample(scenario: WorkPaperXlsxCalcScenario): BenchmarkSample {
  const workbook = scenario.buildXlsxWorkbook()
  xlsxCalc(workbook)
  const sheet = workbook.Sheets['Sheet1']
  if (!sheet) {
    throw new Error('xlsx-calc benchmark fixture did not create Sheet1')
  }

  const memoryBefore = sampleMemory()
  const started = performance.now()
  setXlsxCell(sheet, scenario.fixture.edit.address, scenario.fixture.edit.value)
  xlsxCalc(workbook)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: {
      value: normalizeBenchmarkValue(normalizeXlsxCell(sheet[scenario.fixture.result.address])),
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
  }
}

function aggregateScenario(rowCount: number): WorkPaperXlsxCalcScenario {
  const formula = `=SUM(A1:A${String(rowCount)})`
  const fixture = {
    edit: { address: 'A2500', col: 0, row: 2_499, value: 10_000 },
    family: 'aggregate',
    formula,
    result: { address: 'C1', col: 2, row: 0 },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheet: () => {
      const sheet = numberColumnSheet(rowCount)
      sheet[0] = [1, null, formula]
      return sheet
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

function vlookupScenario(args: {
  readonly editValue: number
  readonly family: 'lookup-approximate' | 'lookup-exact'
  readonly formula: string
  readonly initialLookupValue: number
  readonly rowCount: number
}): WorkPaperXlsxCalcScenario {
  const fixture = {
    edit: { address: 'D1', col: 3, row: 0, value: args.editValue },
    family: args.family,
    formula: args.formula,
    result: { address: 'E1', col: 4, row: 0 },
    rowCount: args.rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheet: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < args.rowCount; row += 1) {
        sheet.push([(row + 1) * 2, (row + 1) * 10])
      }
      sheet[0] = [2, 10, null, args.initialLookupValue, args.formula]
      return sheet
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
    edit: { address: 'A1000', col: 0, row: 999, value: 5_000 },
    family: 'formula-chain',
    formula,
    result: { address: 'C1', col: 2, row: 0 },
    rowCount,
  } as const satisfies WorkPaperXlsxCalcFixture
  return {
    fixture,
    buildWorkPaperSheet: () => {
      const sheet: Array<Array<boolean | number | string | null>> = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        sheet.push([rowNumber, `=A${String(rowNumber)}*2`])
      }
      sheet[0] = [1, '=A1*2', formula]
      return sheet
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

function numberColumnSheet(rowCount: number): Array<Array<boolean | number | string | null>> {
  const sheet: Array<Array<boolean | number | string | null>> = []
  for (let row = 0; row < rowCount; row += 1) {
    sheet.push([row + 1])
  }
  return sheet
}

function createXlsxWorkbook(ref: string): XlsxWorkbook {
  return {
    SheetNames: ['Sheet1'],
    Sheets: {
      Sheet1: { '!ref': ref },
    },
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
