#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { performance } from 'node:perf_hooks'

import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import * as XLSX from 'xlsx'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  buildConditionalAggregationSheet,
  buildFormulaFanoutRow,
  buildParserCacheTemplateSheet,
  buildValueFormulaRows,
} from '../packages/benchmarks/src/workpaper-benchmark-fixtures.js'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export type MicrosoftExcelLiveRecalculationWorkload =
  | 'dirty-fanout-edit'
  | 'suspended-batch-single-column-edit'
  | 'conditional-aggregation-criteria-edit'
  | 'full-rebuild-recalculate'

export interface MicrosoftExcelLiveRecalculationCase {
  readonly id: string
  readonly workload: MicrosoftExcelLiveRecalculationWorkload
  readonly fixture: {
    readonly rowCount: number
    readonly formulaCount: number
    readonly materializedCells: number
  }
  readonly sampleCount: number
  readonly workpaperElapsedMs: NumericSummary
  readonly microsoftExcelElapsedMs: NumericSummary
  readonly workpaperToMicrosoftExcelMeanRatio: number
  readonly workpaperToMicrosoftExcelP95Ratio: number
  readonly tenXMeanAndP95: boolean
  readonly verification: {
    readonly workpaper: Record<string, boolean | number | string | null>
    readonly microsoftExcel: Record<string, boolean | number | string | null>
    readonly equivalent: boolean
  }
  readonly passed: boolean
}

export interface MicrosoftExcelLiveRecalculationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'microsoft-excel-live-recalculation-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-local-microsoft-excel-automation'
    readonly appleScriptTransport: 'osascript'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly warmupCount: number
    readonly screenUpdating: false
    readonly calculationMode: 'manual-during-measurement'
  }
  readonly microsoftExcel: {
    readonly appPath: '/Applications/Microsoft Excel.app'
    readonly version: string
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly workpaperWins: number
    readonly coveredWorkloads: MicrosoftExcelLiveRecalculationWorkload[]
    readonly googleSheetsEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: MicrosoftExcelLiveRecalculationCase[]
}

type RecalculationSheet = ReadonlyArray<ReadonlyArray<boolean | number | string | null>>

interface RecalculationCaseSpec {
  readonly id: string
  readonly workload: MicrosoftExcelLiveRecalculationWorkload
  readonly rowCount: number
  readonly formulaCount: number
  readonly materializedCells: number
}

interface ExcelSampleResult {
  readonly elapsedMs: number
  readonly excelVersion: string
  readonly verification: Record<string, boolean | number | string | null>
}

interface RecalculationCaseRun {
  readonly caseResult: MicrosoftExcelLiveRecalculationCase
  readonly excelVersion: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'microsoft-excel-live-recalculation-scorecard.json')
const excelAppPath = '/Applications/Microsoft Excel.app' as const
const worksheetName = 'Bench'
const sampleCount = 5
const warmupCount = 1
const TEN_X_RATIO = 0.1
const fanoutCount = 1_000
const singleColumnRowCount = 1_000
const aggregationRowCount = 1_000
const aggregationFormulaCopies = 8
const rebuildRowCount = 1_000
const caseSpecs = [
  {
    id: 'excel-live-recalculation-dirty-fanout-edit',
    workload: 'dirty-fanout-edit',
    rowCount: 1,
    formulaCount: fanoutCount,
    materializedCells: fanoutCount + 1,
  },
  {
    id: 'excel-live-recalculation-suspended-batch-single-column-edit',
    workload: 'suspended-batch-single-column-edit',
    rowCount: singleColumnRowCount,
    formulaCount: singleColumnRowCount,
    materializedCells: singleColumnRowCount * 2,
  },
  {
    id: 'excel-live-recalculation-conditional-aggregation-criteria-edit',
    workload: 'conditional-aggregation-criteria-edit',
    rowCount: aggregationRowCount,
    formulaCount: aggregationFormulaCopies * 2,
    materializedCells: aggregationRowCount * 2 + aggregationFormulaCopies * 2 + 4,
  },
  {
    id: 'excel-live-recalculation-full-rebuild-recalculate',
    workload: 'full-rebuild-recalculate',
    rowCount: rebuildRowCount,
    formulaCount: rebuildRowCount * 4,
    materializedCells: rebuildRowCount * 6,
  },
] as const satisfies readonly RecalculationCaseSpec[]

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Microsoft Excel live recalculation scorecard is missing. Run: bun scripts/gen-microsoft-excel-live-recalculation-scorecard.ts`,
      )
    }
    const scorecard = parseMicrosoftExcelLiveRecalculationScorecard(readJsonObject(outputPath))
    validateMicrosoftExcelLiveRecalculationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = buildMicrosoftExcelLiveRecalculationScorecard(new Date().toISOString())
  validateMicrosoftExcelLiveRecalculationScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function buildMicrosoftExcelLiveRecalculationScorecard(generatedAt: string): MicrosoftExcelLiveRecalculationScorecard {
  if (!existsSync(excelAppPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${excelAppPath}`)
  }

  const caseRuns = caseSpecs.map(runRecalculationCase)
  const cases = caseRuns.map((entry) => entry.caseResult)
  const excelVersions = new Set(caseRuns.map((entry) => entry.excelVersion))
  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during recalculation benchmark: ${[...excelVersions].join(', ')}`)
  }
  const excelVersion = [...excelVersions][0] ?? ''
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const workpaperWins = cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'microsoft-excel-live-recalculation-performance',
    generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-local-microsoft-excel-automation',
      appleScriptTransport: 'osascript',
    },
    benchmark: {
      sampleCount,
      warmupCount,
      screenUpdating: false,
      calculationMode: 'manual-during-measurement',
    },
    microsoftExcel: {
      appPath: excelAppPath,
      version: excelVersion,
    },
    summary: {
      allRequiredCasesPassed: cases.every((entry) => entry.passed),
      requiredCaseCount: cases.length,
      tenXMeanAndP95CaseCount,
      workpaperWins,
      coveredWorkloads: caseSpecs.map((entry) => entry.workload),
      googleSheetsEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseMicrosoftExcelLiveRecalculationScorecard(value: Record<string, unknown>): MicrosoftExcelLiveRecalculationScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const microsoftExcel = objectField(value, 'microsoftExcel')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'microsoft-excel-live-recalculation-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-local-microsoft-excel-automation'),
      appleScriptTransport: literalField(source, 'appleScriptTransport', 'osascript'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      warmupCount: numberField(benchmark, 'warmupCount'),
      screenUpdating: literalField(benchmark, 'screenUpdating', false),
      calculationMode: literalField(benchmark, 'calculationMode', 'manual-during-measurement'),
    },
    microsoftExcel: {
      appPath: literalField(microsoftExcel, 'appPath', excelAppPath),
      version: stringField(microsoftExcel, 'version'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      tenXMeanAndP95CaseCount: numberField(summary, 'tenXMeanAndP95CaseCount'),
      workpaperWins: numberField(summary, 'workpaperWins'),
      coveredWorkloads: stringArrayField(summary, 'coveredWorkloads').map(parseWorkload),
      googleSheetsEvidence: literalField(summary, 'googleSheetsEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseRecalculationCase),
  }
}

export function validateMicrosoftExcelLiveRecalculationScorecard(scorecard: MicrosoftExcelLiveRecalculationScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedWorkloads = caseSpecs.map((entry) => entry.workload)
  if (scorecard.microsoftExcel.version.trim().length === 0) {
    throw new Error('Microsoft Excel live recalculation scorecard must record an Excel version')
  }
  if (scorecard.benchmark.sampleCount !== sampleCount || scorecard.benchmark.warmupCount !== warmupCount) {
    throw new Error('Microsoft Excel live recalculation scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Microsoft Excel live recalculation scorecard required cases are stale')
  }
  if (JSON.stringify(scorecard.summary.coveredWorkloads) !== JSON.stringify(expectedWorkloads)) {
    throw new Error('Microsoft Excel live recalculation scorecard covered workloads are stale')
  }
  if (
    scorecard.summary.workpaperWins !==
    scorecard.cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length
  ) {
    throw new Error('Microsoft Excel live recalculation scorecard workpaper win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Microsoft Excel live recalculation scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Microsoft Excel live recalculation scorecard has failing required cases: ${failingCases
        .map(
          (entry) =>
            `${entry.id} WorkPaper=${JSON.stringify(entry.verification.workpaper)} Excel=${JSON.stringify(entry.verification.microsoftExcel)}`,
        )
        .join(', ')}`,
    )
  }
  for (const entry of scorecard.cases) {
    validateNumericSummary(entry.workpaperElapsedMs, `${entry.id} WorkPaper elapsedMs`)
    validateNumericSummary(entry.microsoftExcelElapsedMs, `${entry.id} Microsoft Excel elapsedMs`)
    if (!entry.verification.equivalent) {
      throw new Error(`Microsoft Excel live recalculation scorecard verification mismatch: ${entry.id}`)
    }
  }
}

function runRecalculationCase(caseSpec: RecalculationCaseSpec): RecalculationCaseRun {
  const workpaperSamples: number[] = []
  const excelSamples: number[] = []
  const workpaperVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVersions = new Set<string>()

  for (let index = 0; index < warmupCount; index += 1) {
    runWorkPaperSample(caseSpec.workload)
    excelVersions.add(runExcelSample(caseSpec.workload).excelVersion)
  }

  for (let index = 0; index < sampleCount; index += 1) {
    const workpaperSample = runWorkPaperSample(caseSpec.workload)
    const excelSample = runExcelSample(caseSpec.workload)
    workpaperSamples.push(workpaperSample.elapsedMs)
    excelSamples.push(excelSample.elapsedMs)
    workpaperVerifications.push(workpaperSample.verification)
    excelVerifications.push(excelSample.verification)
    excelVersions.add(excelSample.excelVersion)
  }

  const workpaperElapsedMs = summarizeNumbers(workpaperSamples)
  const microsoftExcelElapsedMs = summarizeNumbers(excelSamples)
  const workpaperToMicrosoftExcelMeanRatio = workpaperElapsedMs.mean / microsoftExcelElapsedMs.mean
  const workpaperToMicrosoftExcelP95Ratio = workpaperElapsedMs.p95 / microsoftExcelElapsedMs.p95
  const verification = {
    workpaper: workpaperVerifications[0] ?? {},
    microsoftExcel: excelVerifications[0] ?? {},
    equivalent: verificationsEquivalent(workpaperVerifications, excelVerifications),
  }

  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during recalculation benchmark: ${[...excelVersions].join(', ')}`)
  }

  return {
    excelVersion: [...excelVersions][0] ?? '',
    caseResult: {
      id: caseSpec.id,
      workload: caseSpec.workload,
      fixture: {
        rowCount: caseSpec.rowCount,
        formulaCount: caseSpec.formulaCount,
        materializedCells: caseSpec.materializedCells,
      },
      sampleCount,
      workpaperElapsedMs,
      microsoftExcelElapsedMs,
      workpaperToMicrosoftExcelMeanRatio,
      workpaperToMicrosoftExcelP95Ratio,
      tenXMeanAndP95: workpaperToMicrosoftExcelMeanRatio <= TEN_X_RATIO && workpaperToMicrosoftExcelP95Ratio <= TEN_X_RATIO,
      verification,
      passed:
        verification.equivalent &&
        workpaperElapsedMs.samples.length === sampleCount &&
        microsoftExcelElapsedMs.samples.length === sampleCount,
    },
  }
}

function runWorkPaperSample(workload: MicrosoftExcelLiveRecalculationWorkload): {
  elapsedMs: number
  verification: Record<string, boolean | number | string | null>
} {
  const workbook = WorkPaper.buildFromSheets({ [worksheetName]: sheetForWorkload(workload) })
  const sheetId = workbook.getSheetId(worksheetName)
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error(`Missing WorkPaper recalculation worksheet: ${worksheetName}`)
  }

  const startedAt = performance.now()
  switch (workload) {
    case 'dirty-fanout-edit':
      workbook.setCellContents({ sheet: sheetId, row: 0, col: 0 }, 7)
      break
    case 'suspended-batch-single-column-edit':
      workbook.suspendEvaluation()
      for (let row = 0; row < singleColumnRowCount; row += 1) {
        workbook.setCellContents({ sheet: sheetId, row, col: 0 }, row * 7)
      }
      workbook.resumeEvaluation()
      break
    case 'conditional-aggregation-criteria-edit':
      workbook.setCellContents({ sheet: sheetId, row: 0, col: 3 }, 'B')
      break
    case 'full-rebuild-recalculate':
      workbook.rebuildAndRecalculate()
      break
  }
  const elapsedMs = performance.now() - startedAt
  const verification = workpaperVerification(workbook, sheetId, workload)
  workbook.dispose()
  return { elapsedMs, verification }
}

function runExcelSample(workload: MicrosoftExcelLiveRecalculationWorkload): ExcelSampleResult {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-live-recalculation-'))
  const workbookPath = join(tempDir, 'recalculation.xlsx')
  const scriptPath = join(tempDir, 'run-recalculation.scpt')
  try {
    writeFileSync(workbookPath, createExcelWorkbookBytes(workload))
    writeFileSync(scriptPath, createRecalculationAppleScript(workload))
    return parseExcelSampleOutput(execFileSync('osascript', [scriptPath, workbookPath], { encoding: 'utf8' }).trim())
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function createExcelWorkbookBytes(workload: MicrosoftExcelLiveRecalculationWorkload): Uint8Array {
  const worksheet = aoaToFormulaWorksheet(sheetForWorkload(workload))
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName)
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

function createRecalculationAppleScript(workload: MicrosoftExcelLiveRecalculationWorkload): string {
  return `use framework "Foundation"

on run argv
  set workbookPath to POSIX file (item 1 of argv)
  set output to ""
  tell application "Microsoft Excel"
    set display alerts to false
    set screen updating to false
    set previousCalculation to calculation
    try
      set calculation to calculation manual
      open workbook workbook file name workbookPath
      calculate full rebuild
${excelPreparationStatement(workload)}
      set startedAt to current application's NSDate's timeIntervalSinceReferenceDate()
${excelOperationStatement(workload)}
${excelCalculationStatement(workload)}
      set elapsedMs to ((current application's NSDate's timeIntervalSinceReferenceDate()) - startedAt) * 1000
      set output to "version=" & (version as string)
      set output to output & linefeed & "elapsedMs=" & (elapsedMs as string)
${excelVerificationStatements(workload)}
      close active workbook saving no
      set calculation to previousCalculation
      set screen updating to true
    on error errMsg number errNum
      try
        close active workbook saving no
      end try
      set calculation to previousCalculation
      set screen updating to true
      error errMsg number errNum
    end try
  end tell
  return output
end run
`
}

function excelPreparationStatement(workload: MicrosoftExcelLiveRecalculationWorkload): string {
  if (workload !== 'suspended-batch-single-column-edit') {
    return ''
  }
  return `      set batchValues to {}
      repeat with rowIndex from 1 to ${String(singleColumnRowCount)}
        set end of batchValues to {((rowIndex - 1) * 7)}
      end repeat`
}

function excelOperationStatement(workload: MicrosoftExcelLiveRecalculationWorkload): string {
  switch (workload) {
    case 'dirty-fanout-edit':
      return '      set value of range "A1" of worksheet "Bench" of active workbook to 7'
    case 'suspended-batch-single-column-edit':
      return `      set value of range "A1:A${String(singleColumnRowCount)}" of worksheet "Bench" of active workbook to batchValues`
    case 'conditional-aggregation-criteria-edit':
      return '      set value of range "D1" of worksheet "Bench" of active workbook to "B"'
    case 'full-rebuild-recalculate':
      return '      set value of range "A1" of worksheet "Bench" of active workbook to 1'
  }
}

function excelCalculationStatement(workload: MicrosoftExcelLiveRecalculationWorkload): string {
  return workload === 'full-rebuild-recalculate' ? '      calculate full rebuild' : '      calculate'
}

function excelVerificationStatements(workload: MicrosoftExcelLiveRecalculationWorkload): string {
  switch (workload) {
    case 'dirty-fanout-edit':
      return `      set output to output & linefeed & "terminalValue=" & ((value of range "${columnName(fanoutCount)}1" of worksheet "Bench" of active workbook) as string)
      set output to output & linefeed & "width=" & ((count of columns of used range of worksheet "Bench" of active workbook) as string)`
    case 'suspended-batch-single-column-edit':
      return `      set output to output & linefeed & "sampleFormulaValue=" & ((value of range "B${String(
        singleColumnRowCount,
      )}" of worksheet "Bench" of active workbook) as string)
      set output to output & linefeed & "height=" & ((count of rows of used range of worksheet "Bench" of active workbook) as string)`
    case 'conditional-aggregation-criteria-edit':
      return `      set output to output & linefeed & "sampleSumValue=" & ((value of range "E1" of worksheet "Bench" of active workbook) as string)
      set output to output & linefeed & "sampleCountValue=" & ((value of range "${columnName(
        4 + aggregationFormulaCopies,
      )}1" of worksheet "Bench" of active workbook) as string)`
    case 'full-rebuild-recalculate':
      return `      set output to output & linefeed & "terminalAggregateValue=" & ((value of range "E${String(
        rebuildRowCount,
      )}" of worksheet "Bench" of active workbook) as string)
      set output to output & linefeed & "terminalChainValue=" & ((value of range "F${String(
        rebuildRowCount,
      )}" of worksheet "Bench" of active workbook) as string)`
  }
}

function sheetForWorkload(workload: MicrosoftExcelLiveRecalculationWorkload): RecalculationSheet {
  switch (workload) {
    case 'dirty-fanout-edit':
      return [buildFormulaFanoutRow(fanoutCount)]
    case 'suspended-batch-single-column-edit':
      return buildValueFormulaRows(singleColumnRowCount)
    case 'conditional-aggregation-criteria-edit':
      return buildConditionalAggregationSheet(aggregationRowCount, aggregationFormulaCopies)
    case 'full-rebuild-recalculate':
      return buildParserCacheTemplateSheet(rebuildRowCount)
  }
}

function workpaperVerification(
  workbook: WorkPaper,
  sheetId: number,
  workload: MicrosoftExcelLiveRecalculationWorkload,
): Record<string, boolean | number | string | null> {
  const dimensions = workbook.getSheetDimensions(sheetId)
  switch (workload) {
    case 'dirty-fanout-edit':
      return {
        terminalValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: fanoutCount })),
        width: dimensions.width,
      }
    case 'suspended-batch-single-column-edit':
      return {
        sampleFormulaValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: singleColumnRowCount - 1, col: 1 })),
        height: dimensions.height,
      }
    case 'conditional-aggregation-criteria-edit':
      return {
        sampleSumValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: 4 })),
        sampleCountValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: 4 + aggregationFormulaCopies })),
      }
    case 'full-rebuild-recalculate':
      return {
        terminalAggregateValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: rebuildRowCount - 1, col: 4 })),
        terminalChainValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: rebuildRowCount - 1, col: 5 })),
      }
  }
}

function aoaToFormulaWorksheet(sheet: RecalculationSheet): XLSX.WorkSheet {
  const worksheet = XLSX.utils.aoa_to_sheet(
    sheet.map((row) => row.map((cell) => (typeof cell === 'string' && cell.startsWith('=') ? null : cell))),
  )
  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex] ?? []
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = row[colIndex]
      if (typeof cell === 'string' && cell.startsWith('=')) {
        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })] = { t: 'n', f: cell.slice(1) }
      }
    }
  }
  return worksheet
}

function parseExcelSampleOutput(rawOutput: string): ExcelSampleResult {
  const values = new Map<string, string>()
  for (const line of rawOutput.split(/\r?\n/u)) {
    const separatorIndex = line.indexOf('=')
    if (separatorIndex > 0) {
      values.set(line.slice(0, separatorIndex), line.slice(separatorIndex + 1))
    }
  }
  return {
    excelVersion: requiredString(values, 'version'),
    elapsedMs: Number(requiredString(values, 'elapsedMs')),
    verification: Object.fromEntries(
      [...values.entries()]
        .filter(([key]) => key !== 'version' && key !== 'elapsedMs')
        .map(([key, value]) => [key, parseExcelVerificationValue(value)]),
    ),
  }
}

function parseExcelVerificationValue(value: string): number | string | null {
  if (value.length === 0 || value === 'missing value') {
    return null
  }
  const numericValue = Number(value)
  return Number.isFinite(numericValue) ? numericValue : value
}

function normalizeProtocolValue(value: unknown): boolean | number | string | null {
  if (!isProtocolValueLike(value)) {
    throw new Error('Unexpected non-protocol WorkPaper value in recalculation benchmark')
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return null
    default:
      return null
  }
}

function parseRecalculationCase(value: unknown): MicrosoftExcelLiveRecalculationCase {
  const record = asObject(value, 'Microsoft Excel live recalculation case')
  const fixture = objectField(record, 'fixture')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    fixture: {
      rowCount: numberField(fixture, 'rowCount'),
      formulaCount: numberField(fixture, 'formulaCount'),
      materializedCells: numberField(fixture, 'materializedCells'),
    },
    sampleCount: numberField(record, 'sampleCount'),
    workpaperElapsedMs: parseNumericSummary(objectField(record, 'workpaperElapsedMs')),
    microsoftExcelElapsedMs: parseNumericSummary(objectField(record, 'microsoftExcelElapsedMs')),
    workpaperToMicrosoftExcelMeanRatio: numberField(record, 'workpaperToMicrosoftExcelMeanRatio'),
    workpaperToMicrosoftExcelP95Ratio: numberField(record, 'workpaperToMicrosoftExcelP95Ratio'),
    tenXMeanAndP95: booleanField(record, 'tenXMeanAndP95'),
    verification: {
      workpaper: parseVerificationRecord(objectField(verification, 'workpaper')),
      microsoftExcel: parseVerificationRecord(objectField(verification, 'microsoftExcel')),
      equivalent: booleanField(verification, 'equivalent'),
    },
    passed: booleanField(record, 'passed'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  const confidence95 = objectField(value, 'confidence95')
  return {
    samples: arrayField(value, 'samples').map((entry) => {
      if (typeof entry !== 'number' || !Number.isFinite(entry)) {
        throw new Error('Expected numeric summary samples to be finite numbers')
      }
      return entry
    }),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
    standardDeviation: numberField(value, 'standardDeviation'),
    relativeStandardDeviation: numberField(value, 'relativeStandardDeviation'),
    standardError: numberField(value, 'standardError'),
    confidence95: {
      low: numberField(confidence95, 'low'),
      high: numberField(confidence95, 'high'),
    },
  }
}

function parseVerificationRecord(value: Record<string, unknown>): Record<string, boolean | number | string | null> {
  const result: Record<string, boolean | number | string | null> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (entry === null || typeof entry === 'boolean' || typeof entry === 'number' || typeof entry === 'string') {
      result[key] = entry
      continue
    }
    throw new Error(`Unsupported verification value for ${key}`)
  }
  return result
}

function parseWorkload(value: string): MicrosoftExcelLiveRecalculationWorkload {
  if (
    value === 'dirty-fanout-edit' ||
    value === 'suspended-batch-single-column-edit' ||
    value === 'conditional-aggregation-criteria-edit' ||
    value === 'full-rebuild-recalculate'
  ) {
    return value
  }
  throw new Error(`Unexpected Microsoft Excel live recalculation workload: ${value}`)
}

function verificationsEquivalent(
  workpaperVerifications: readonly Record<string, boolean | number | string | null>[],
  excelVerifications: readonly Record<string, boolean | number | string | null>[],
): boolean {
  if (workpaperVerifications.length !== excelVerifications.length) {
    return false
  }
  return workpaperVerifications.every((workpaperSnapshot, index) => {
    const excelVerification = excelVerifications[index]
    return excelVerification !== undefined && JSON.stringify(workpaperSnapshot) === JSON.stringify(excelVerification)
  })
}

function requiredString(values: ReadonlyMap<string, string>, key: string): string {
  const value = values.get(key)
  if (value === undefined || value.length === 0) {
    throw new Error(`Missing Microsoft Excel recalculation output field: ${key}`)
  }
  return value
}

function validateNumericSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length !== sampleCount || summary.samples.some((entry) => !Number.isFinite(entry) || entry < 0)) {
    throw new Error(`${label} has invalid samples`)
  }
}

function isProtocolValueLike(value: unknown): value is { tag: ValueTag; value?: boolean | number | string } {
  if (value === null || typeof value !== 'object') {
    return false
  }
  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function columnName(columnIndex: number): string {
  let index = columnIndex + 1
  let name = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    index = Math.floor((index - 1) / 26)
  }
  return name
}

function logResult(mode: 'check' | 'write', scorecard: MicrosoftExcelLiveRecalculationScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        excelVersion: scorecard.microsoftExcel.version,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        workpaperWins: scorecard.summary.workpaperWins,
        tenXMeanAndP95CaseCount: scorecard.summary.tenXMeanAndP95CaseCount,
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
