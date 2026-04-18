import { performance } from 'node:perf_hooks'
import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import type { RawCellContent as HyperFormulaRawCellContent, Sheet as HyperFormulaSheet } from 'hyperformula'
import type { MemoryMeasurement } from './metrics.js'
import { measureMemory, sampleMemory } from './metrics.js'
import type { NumericSummary } from './stats.js'
import { summarizeNumbers } from './stats.js'
import {
  address,
  buildDenseLiteralSheet,
  buildDynamicArraySheet,
  buildFormulaChainRow,
  buildLookupSheet,
  buildValueFormulaRows,
  range,
} from './workpaper-benchmark-fixtures.js'

const { HyperFormula } = await import('hyperformula')

export const DEFAULT_COMPETITIVE_WARMUP_COUNT = 2
export const DEFAULT_COMPETITIVE_SAMPLE_COUNT = 5
export const HYPERFORMULA_LICENSE_KEY = 'gpl-v3'

export type ComparativeBenchmarkWorkload =
  | 'build-from-sheets'
  | 'single-edit-recalc'
  | 'batch-edit-recalc'
  | 'range-read'
  | 'lookup-no-column-index'
  | 'lookup-with-column-index'
  | 'dynamic-array-filter'

export type ComparativeBenchmarkCategory = 'directly-comparable' | 'leadership'

export interface ComparativeBenchmarkSuiteOptions {
  sampleCount?: number
  warmupCount?: number
}

export interface ComparativeMeasuredEngineResult {
  status: 'supported'
  elapsedMs: NumericSummary
  memoryDeltaBytes: ComparativeMemorySummary
  engineCounters?: Record<string, NumericSummary>
  verification: Record<string, unknown>
}

export interface ComparativeUnsupportedEngineResult {
  status: 'unsupported'
  reason: string
  evidence: readonly string[]
}

export interface ComparativeComparableResult {
  workload: ComparativeBenchmarkWorkload
  category: 'directly-comparable'
  comparable: true
  fixture: Record<string, unknown>
  comparison: {
    fasterEngine: 'workpaper' | 'hyperformula'
    meanSpeedup: number
    verificationEquivalent: true
  }
  engines: {
    workpaper: ComparativeMeasuredEngineResult
    hyperformula: ComparativeMeasuredEngineResult
  }
}

export interface ComparativeLeadershipResult {
  workload: ComparativeBenchmarkWorkload
  category: 'leadership'
  comparable: false
  fixture: Record<string, unknown>
  note: string
  engines: {
    workpaper: ComparativeMeasuredEngineResult
    hyperformula: ComparativeUnsupportedEngineResult
  }
}

export interface ComparativeMemorySummary {
  rssBytes: NumericSummary
  heapUsedBytes: NumericSummary
  heapTotalBytes: NumericSummary
  externalBytes: NumericSummary
  arrayBuffersBytes: NumericSummary
}

interface BenchmarkSample {
  elapsedMs: number
  memory: MemoryMeasurement
  verification: Record<string, unknown>
}

export type ComparativeBenchmarkResult = ComparativeComparableResult | ComparativeLeadershipResult

export function runWorkPaperVsHyperFormulaBenchmarkSuite(options: ComparativeBenchmarkSuiteOptions = {}): ComparativeBenchmarkResult[] {
  const runtimeOptions = resolveSuiteOptions(options)
  return [
    runComparableScenario(
      'build-from-sheets',
      {
        cols: 24,
        materializedCells: 160 * 24,
        rows: 160,
      },
      runtimeOptions,
      () => measureWorkPaperBuildSample(160, 24),
      () => measureHyperFormulaBuildSample(160, 24),
    ),
    runComparableScenario(
      'single-edit-recalc',
      {
        downstreamCount: 2_000,
      },
      runtimeOptions,
      () => measureWorkPaperSingleEditSample(2_000),
      () => measureHyperFormulaSingleEditSample(2_000),
    ),
    runComparableScenario(
      'batch-edit-recalc',
      {
        editCount: 500,
      },
      runtimeOptions,
      () => measureWorkPaperBatchEditSample(500),
      () => measureHyperFormulaBatchEditSample(500),
    ),
    runComparableScenario(
      'range-read',
      {
        cols: 24,
        materializedCells: 240 * 24,
        rows: 240,
      },
      runtimeOptions,
      () => measureWorkPaperRangeReadSample(240, 24),
      () => measureHyperFormulaRangeReadSample(240, 24),
    ),
    runComparableScenario(
      'lookup-no-column-index',
      {
        rowCount: 5_000,
        useColumnIndex: false,
      },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, false),
      () => measureHyperFormulaLookupSample(5_000, false),
    ),
    runComparableScenario(
      'lookup-with-column-index',
      {
        rowCount: 5_000,
        useColumnIndex: true,
      },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, true),
      () => measureHyperFormulaLookupSample(5_000, true),
    ),
    runLeadershipScenario(
      'dynamic-array-filter',
      {
        rowCount: 750,
        formula: '=FILTER(A2:A751,A2:A751>B1)',
      },
      runtimeOptions,
      () => measureWorkPaperDynamicArraySample(750),
      {
        status: 'unsupported',
        evidence: [
          '/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md',
          '/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts',
        ],
        reason: 'HyperFormula 3.2.0 documents dynamic arrays as unsupported.',
      },
    ),
  ]
}

if (import.meta.url === `file://${process.argv[1]}`) {
  console.log(
    JSON.stringify(
      runWorkPaperVsHyperFormulaBenchmarkSuite({
        sampleCount: DEFAULT_COMPETITIVE_SAMPLE_COUNT,
        warmupCount: DEFAULT_COMPETITIVE_WARMUP_COUNT,
      }),
      null,
      2,
    ),
  )
}

function runComparableScenario(
  workload: ComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  runHyperFormulaSample: () => BenchmarkSample,
): ComparativeComparableResult {
  const workpaper = benchmarkSupportedEngine(runWorkPaperSample, options)
  const hyperformula = benchmarkSupportedEngine(runHyperFormulaSample, options)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const hyperFormulaVerification = JSON.stringify(hyperformula.verification)
  if (workPaperVerification !== hyperFormulaVerification) {
    throw new Error(
      `Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== HyperFormula ${hyperFormulaVerification}`,
    )
  }

  const fasterEngine = workpaper.elapsedMs.mean <= hyperformula.elapsedMs.mean ? 'workpaper' : 'hyperformula'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : hyperformula.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? hyperformula.elapsedMs.mean : workpaper.elapsedMs.mean

  return {
    workload,
    category: 'directly-comparable',
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
  }
}

function runLeadershipScenario(
  workload: ComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  hyperformula: ComparativeUnsupportedEngineResult,
): ComparativeLeadershipResult {
  return {
    workload,
    category: 'leadership',
    comparable: false,
    fixture,
    note: 'This workload demonstrates capability leadership and is not an apples-to-apples speed comparison.',
    engines: {
      workpaper: benchmarkSupportedEngine(runWorkPaperSample, options),
      hyperformula,
    },
  }
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: Required<ComparativeBenchmarkSuiteOptions>,
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

function measureWorkPaperBuildSample(rows: number, cols: number): BenchmarkSample {
  const sheet = buildDenseLiteralSheet(rows, cols)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const workbook = WorkPaper.buildFromSheets({ Bench: sheet })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const sheetId = workbook.getSheetId('Bench')!
  const verification = {
    dimensions: workbook.getSheetDimensions(sheetId),
    terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureHyperFormulaBuildSample(rows: number, cols: number): BenchmarkSample {
  const sheet = toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols))
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const workbook = HyperFormula.buildFromSheets({ Bench: sheet }, { licenseKey: HYPERFORMULA_LICENSE_KEY })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const sheetId = workbook.getSheetId('Bench')!
  const verification = {
    dimensions: workbook.getSheetDimensions(sheetId),
    terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
  }
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureWorkPaperSingleEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: [buildFormulaChainRow(downstreamCount)],
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(address(sheetId, 0, 0), 99)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? null,
    terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureHyperFormulaSingleEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    {
      Bench: toHyperFormulaSheet([buildFormulaChainRow(downstreamCount)]),
    },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(address(sheetId, 0, 0), 99)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? null,
    terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
  }
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureWorkPaperBatchEditSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildValueFormulaRows(editCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.batch(() => {
    for (let row = 0; row < editCount; row += 1) {
      workbook.setCellContents(address(sheetId, row, 0), row * 3)
    }
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    sampleFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureHyperFormulaBatchEditSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    {
      Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)),
    },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.batch(() => {
    for (let row = 0; row < editCount; row += 1) {
      workbook.setCellContents(address(sheetId, row, 0), row * 3)
    }
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    sampleFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
  }
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureWorkPaperRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDenseLiteralSheet(rows, cols),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const values = workbook.getRangeValues(targetRange)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const lastRow = values.at(-1)
  const verification = {
    readCols: values[0]?.length ?? 0,
    readRows: values.length,
    terminalValue: normalizeWorkPaperValue(lastRow?.at(-1)),
    topLeftValue: normalizeWorkPaperValue(values[0]?.[0]),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureHyperFormulaRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    {
      Bench: toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols)),
    },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const values = workbook.getRangeValues(targetRange)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const lastRow = values.at(-1)
  const verification = {
    readCols: values[0]?.length ?? 0,
    readRows: values.length,
    terminalValue: normalizeHyperFormulaValue(lastRow?.at(-1)),
    topLeftValue: normalizeHyperFormulaValue(values[0]?.[0]),
  }
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureWorkPaperLookupSample(rowCount: number, useColumnIndex: boolean): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(
    {
      Bench: buildLookupSheet(rowCount),
    },
    {
      useColumnIndex,
    },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(targetAddress, rowCount)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureHyperFormulaLookupSample(rowCount: number, useColumnIndex: boolean): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    {
      Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)),
    },
    {
      licenseKey: HYPERFORMULA_LICENSE_KEY,
      useColumnIndex,
    },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(targetAddress, rowCount)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
  }
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
  }
}

function measureWorkPaperDynamicArraySample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDynamicArraySheet(rowCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const thresholdAddress = address(sheetId, 0, 1)
  const spillAnchor = address(sheetId, 0, 2)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(thresholdAddress, rowCount - 10)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const verification = {
    changeCount: changes.length,
    spillHeight: workbook.getSheetDimensions(sheetId).height,
    spillIsArray: workbook.isCellPartOfArray(spillAnchor),
    spillValue: normalizeWorkPaperValue(workbook.getCellValue(spillAnchor)),
  }
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification,
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

function normalizeWorkPaperValue(value: unknown): boolean | number | string | null | { error: unknown } {
  if (!isProtocolValueLike(value)) {
    return null
  }

  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return { error: value.code ?? 'ERROR' }
    default:
      return { error: `UNKNOWN_TAG_${String(value.tag)}` }
  }
}

function normalizeHyperFormulaValue(value: unknown): boolean | number | string | null | { error: unknown } {
  if (value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
    return value
  }
  if (isHyperFormulaErrorLike(value)) {
    return { error: value.value }
  }
  return { error: 'UNKNOWN_VALUE' }
}

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): Required<ComparativeBenchmarkSuiteOptions> {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
}

function isProtocolValueLike(value: unknown): value is { code?: unknown; tag: ValueTag; value?: boolean | number | string } {
  if (!value || typeof value !== 'object') {
    return false
  }

  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function isHyperFormulaErrorLike(value: unknown): value is { value: unknown } {
  return value !== null && typeof value === 'object' && 'value' in value
}

function toHyperFormulaSheet(sheet: ReadonlyArray<ReadonlyArray<unknown>>): HyperFormulaSheet {
  return sheet.map((row) => row.map((cell) => toHyperFormulaCell(cell)))
}

function toHyperFormulaCell(cell: unknown): HyperFormulaRawCellContent {
  if (cell === null || typeof cell === 'boolean' || typeof cell === 'number' || typeof cell === 'string') {
    return cell
  }
  throw new Error(`Unsupported HyperFormula benchmark cell type: ${typeof cell}`)
}
