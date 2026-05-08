#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import * as XLSX from 'xlsx'
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

export type CalculationScalarValue = boolean | number | string | null | { readonly error: string }

export interface MicrosoftExcelLiveCalculationCase {
  readonly id: string
  readonly formula: string
  readonly formulaCell: string
  readonly coveredFeature: string
  readonly biligValue: CalculationScalarValue
  readonly microsoftExcelRawValue: string
  readonly microsoftExcelValue: CalculationScalarValue
  readonly passed: boolean
}

export interface MicrosoftExcelLiveCalculationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'microsoft-excel-live-calculation-correctness'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-local-microsoft-excel-automation'
    readonly appleScriptTransport: 'osascript'
  }
  readonly microsoftExcel: {
    readonly appPath: '/Applications/Microsoft Excel.app'
    readonly version: string
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly matchingCaseCount: number
    readonly coveredFeatures: string[]
    readonly googleSheetsEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: MicrosoftExcelLiveCalculationCase[]
}

interface CalculationCaseSpec {
  readonly id: string
  readonly formula: string
  readonly formulaRowIndex: number
  readonly coveredFeature: string
  readonly inputValues: readonly (boolean | number | string | null)[]
}

interface LiveExcelEvaluation {
  readonly excelVersion: string
  readonly rawValuesByCaseId: ReadonlyMap<string, string>
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'microsoft-excel-live-calculation-scorecard.json')
const excelAppPath = '/Applications/Microsoft Excel.app' as const
const formulaColumnIndex = 3
const worksheetName = 'Cases'
const requiredCaseSpecs = [
  {
    id: 'arithmetic-precedence',
    formula: '=A2+B2*2',
    formulaRowIndex: 1,
    coveredFeature: 'excelLive.arithmeticPrecedence',
    inputValues: [10, 7, null],
  },
  {
    id: 'aggregate-sum-range',
    formula: '=SUM(A3:C3)',
    formulaRowIndex: 2,
    coveredFeature: 'excelLive.aggregateSumRange',
    inputValues: [1, 2, 3],
  },
  {
    id: 'conditional-if-comparison',
    formula: '=IF(A4>B4,"over","under")',
    formulaRowIndex: 3,
    coveredFeature: 'excelLive.conditionalIfComparison',
    inputValues: [12, 10, null],
  },
  {
    id: 'numeric-rounding',
    formula: '=ROUND(A5/B5,2)',
    formulaRowIndex: 4,
    coveredFeature: 'excelLive.numericRounding',
    inputValues: [10, 3, null],
  },
  {
    id: 'aggregate-count-range',
    formula: '=COUNT(A6:C6)',
    formulaRowIndex: 5,
    coveredFeature: 'excelLive.aggregateCountRange',
    inputValues: [1, 'x', 2],
  },
  {
    id: 'text-concat',
    formula: '=CONCAT(A7,B7)',
    formulaRowIndex: 6,
    coveredFeature: 'excelLive.textConcat',
    inputValues: ['Bi', 'lig', null],
  },
  {
    id: 'textjoin-ignore-empty-range',
    formula: '=TEXTJOIN("-",TRUE,A8:C8)',
    formulaRowIndex: 7,
    coveredFeature: 'excelLive.textJoinIgnoreEmptyRange',
    inputValues: ['a', '', 'c'],
  },
  {
    id: 'lookup-xlookup-exact',
    formula: '=XLOOKUP("b",$A$20:$A$22,$B$20:$B$22)',
    formulaRowIndex: 8,
    coveredFeature: 'excelLive.lookupXlookupExact',
    inputValues: [null, null, null],
  },
  {
    id: 'lookup-match-exact',
    formula: '=MATCH("b",$A$20:$A$22,0)',
    formulaRowIndex: 9,
    coveredFeature: 'excelLive.lookupMatchExact',
    inputValues: [null, null, null],
  },
  {
    id: 'lookup-index-range',
    formula: '=INDEX($B$20:$B$22,2,1)',
    formulaRowIndex: 10,
    coveredFeature: 'excelLive.lookupIndexRange',
    inputValues: [null, null, null],
  },
  {
    id: 'boolean-and',
    formula: '=AND(TRUE,A2>0)',
    formulaRowIndex: 11,
    coveredFeature: 'excelLive.booleanAnd',
    inputValues: [null, null, null],
  },
  {
    id: 'boolean-or',
    formula: '=OR(FALSE,A2=10)',
    formulaRowIndex: 12,
    coveredFeature: 'excelLive.booleanOr',
    inputValues: [null, null, null],
  },
  {
    id: 'date-serial',
    formula: '=N(DATE(2026,5,6))',
    formulaRowIndex: 13,
    coveredFeature: 'excelLive.dateSerial',
    inputValues: [null, null, null],
  },
  {
    id: 'date-year-from-serial',
    formula: '=YEAR(D14)',
    formulaRowIndex: 14,
    coveredFeature: 'excelLive.dateYearFromSerial',
    inputValues: [null, null, null],
  },
  {
    id: 'conditional-sumif-range',
    formula: '=SUMIF($A$20:$A$22,"b",$B$20:$B$22)',
    formulaRowIndex: 15,
    coveredFeature: 'excelLive.conditionalSumifRange',
    inputValues: [null, null, null],
  },
  {
    id: 'math-abs-sqrt',
    formula: '=ABS(A18)+SQRT(16)',
    formulaRowIndex: 17,
    coveredFeature: 'excelLive.mathAbsSqrt',
    inputValues: [-9, null, null],
  },
  {
    id: 'aggregate-average-range',
    formula: '=AVERAGE(A23:C23)',
    formulaRowIndex: 22,
    coveredFeature: 'excelLive.aggregateAverageRange',
    inputValues: [2, 4, 6],
  },
  {
    id: 'aggregate-min-range',
    formula: '=MIN(A24:C24)',
    formulaRowIndex: 23,
    coveredFeature: 'excelLive.aggregateMinRange',
    inputValues: [8, 2, 5],
  },
  {
    id: 'aggregate-max-range',
    formula: '=MAX(A25:C25)',
    formulaRowIndex: 24,
    coveredFeature: 'excelLive.aggregateMaxRange',
    inputValues: [8, 2, 5],
  },
  {
    id: 'aggregate-counta-range',
    formula: '=COUNTA(A26:C26)',
    formulaRowIndex: 25,
    coveredFeature: 'excelLive.aggregateCountaRange',
    inputValues: [1, null, 'x'],
  },
  {
    id: 'aggregate-countblank-range',
    formula: '=COUNTBLANK(A27:C27)',
    formulaRowIndex: 26,
    coveredFeature: 'excelLive.aggregateCountblankRange',
    inputValues: [1, null, 'x'],
  },
  {
    id: 'math-product-range',
    formula: '=PRODUCT(A28:C28)',
    formulaRowIndex: 27,
    coveredFeature: 'excelLive.mathProductRange',
    inputValues: [2, 3, 4],
  },
  {
    id: 'math-sumsq-range',
    formula: '=SUMSQ(A29:C29)',
    formulaRowIndex: 28,
    coveredFeature: 'excelLive.mathSumsqRange',
    inputValues: [2, 3, 4],
  },
  {
    id: 'math-mod',
    formula: '=MOD(A30,B30)',
    formulaRowIndex: 29,
    coveredFeature: 'excelLive.mathMod',
    inputValues: [10, 3, null],
  },
  {
    id: 'math-power',
    formula: '=POWER(A31,B31)',
    formulaRowIndex: 30,
    coveredFeature: 'excelLive.mathPower',
    inputValues: [2, 3, null],
  },
  {
    id: 'math-gcd-range',
    formula: '=GCD(A32:C32)',
    formulaRowIndex: 31,
    coveredFeature: 'excelLive.mathGcdRange',
    inputValues: [2, 3, 4],
  },
  {
    id: 'text-left',
    formula: '=LEFT(A33,2)',
    formulaRowIndex: 32,
    coveredFeature: 'excelLive.textLeft',
    inputValues: ['Bilig', null, null],
  },
  {
    id: 'text-right',
    formula: '=RIGHT(A34,3)',
    formulaRowIndex: 33,
    coveredFeature: 'excelLive.textRight',
    inputValues: ['Bilig', null, null],
  },
  {
    id: 'text-mid',
    formula: '=MID(A35,2,3)',
    formulaRowIndex: 34,
    coveredFeature: 'excelLive.textMid',
    inputValues: ['Bilig', null, null],
  },
  {
    id: 'text-len',
    formula: '=LEN(A36)',
    formulaRowIndex: 35,
    coveredFeature: 'excelLive.textLen',
    inputValues: ['Bilig', null, null],
  },
  {
    id: 'text-trim-upper',
    formula: '=UPPER(TRIM(A37))',
    formulaRowIndex: 36,
    coveredFeature: 'excelLive.textTrimUpper',
    inputValues: ['  Bilig  ', null, null],
  },
  {
    id: 'text-search-case-insensitive',
    formula: '=SEARCH("LI",A38)',
    formulaRowIndex: 37,
    coveredFeature: 'excelLive.textSearchCaseInsensitive',
    inputValues: ['Bilig', null, null],
  },
  {
    id: 'lookup-vlookup-exact',
    formula: '=VLOOKUP("b",$A$20:$B$22,2,FALSE)',
    formulaRowIndex: 38,
    coveredFeature: 'excelLive.lookupVlookupExact',
    inputValues: [null, null, null],
  },
  {
    id: 'lookup-choose',
    formula: '=CHOOSE(2,"red","blue","green")',
    formulaRowIndex: 39,
    coveredFeature: 'excelLive.lookupChoose',
    inputValues: [null, null, null],
  },
  {
    id: 'statistical-countif',
    formula: '=COUNTIF(A41:C41,">0")',
    formulaRowIndex: 40,
    coveredFeature: 'excelLive.statisticalCountif',
    inputValues: [2, 4, -1],
  },
  {
    id: 'statistical-averageif',
    formula: '=AVERAGEIF(A42:C42,">0")',
    formulaRowIndex: 41,
    coveredFeature: 'excelLive.statisticalAverageif',
    inputValues: [2, 4, -1],
  },
] as const satisfies readonly CalculationCaseSpec[]
const requiredCaseIds = requiredCaseSpecs.map((entry) => entry.id)
const requiredCoveredFeatures = requiredCaseSpecs.map((entry) => entry.coveredFeature)
const sharedRows: ReadonlyMap<number, readonly (boolean | number | string | null)[]> = new Map([
  [19, ['a', 10]],
  [20, ['b', 20]],
  [21, ['c', 30]],
])

export const calculationLiveWorksheetName = worksheetName
export const calculationLiveCaseSpecs = requiredCaseSpecs
export const calculationLiveRequiredCaseIds = requiredCaseIds
export const calculationLiveRequiredCoveredFeatures = requiredCoveredFeatures

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Microsoft Excel live calculation scorecard is missing. Run: bun scripts/gen-microsoft-excel-live-calculation-scorecard.ts`,
      )
    }
    const scorecard = parseMicrosoftExcelLiveCalculationScorecard(readJsonObject(outputPath))
    validateMicrosoftExcelLiveCalculationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = buildMicrosoftExcelLiveCalculationScorecard(new Date().toISOString(), evaluateLiveExcelCases())
  validateMicrosoftExcelLiveCalculationScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function buildMicrosoftExcelLiveCalculationScorecard(
  generatedAt: string,
  liveExcelEvaluation: LiveExcelEvaluation,
): MicrosoftExcelLiveCalculationScorecard {
  const biligValuesByCaseId = evaluateBiligCases()
  const cases = requiredCaseSpecs.map((caseSpec) => {
    const biligValue = requiredMapValue(biligValuesByCaseId, caseSpec.id, 'Bilig calculation value')
    const microsoftExcelRawValue = requiredMapValue(liveExcelEvaluation.rawValuesByCaseId, caseSpec.id, 'Microsoft Excel calculation value')
    const microsoftExcelValue = parseExcelRawValue(microsoftExcelRawValue, biligValue)
    return {
      id: caseSpec.id,
      formula: caseSpec.formula,
      formulaCell: toA1Address(caseSpec.formulaRowIndex, formulaColumnIndex),
      coveredFeature: caseSpec.coveredFeature,
      biligValue,
      microsoftExcelRawValue,
      microsoftExcelValue,
      passed: valuesEquivalent(biligValue, microsoftExcelValue),
    }
  })
  const matchingCaseCount = cases.filter((entry) => entry.passed).length

  return {
    schemaVersion: 1,
    suite: 'microsoft-excel-live-calculation-correctness',
    generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-local-microsoft-excel-automation',
      appleScriptTransport: 'osascript',
    },
    microsoftExcel: {
      appPath: excelAppPath,
      version: liveExcelEvaluation.excelVersion,
    },
    summary: {
      allRequiredCasesPassed: matchingCaseCount === cases.length,
      requiredCaseCount: cases.length,
      matchingCaseCount,
      coveredFeatures: [...requiredCoveredFeatures],
      googleSheetsEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function evaluateBiligCases(): Map<string, CalculationScalarValue> {
  const workbook = WorkPaper.buildFromSheets({ [worksheetName]: buildCaseRowsForWorkPaper() })
  const sheetId = workbook.getSheetId(worksheetName)
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error(`Missing Bilig live calculation worksheet: ${worksheetName}`)
  }
  try {
    return new Map(
      requiredCaseSpecs.map((caseSpec) => [
        caseSpec.id,
        normalizeBiligValue(workbook.getCellValue({ sheet: sheetId, row: caseSpec.formulaRowIndex, col: formulaColumnIndex })),
      ]),
    )
  } finally {
    workbook.dispose()
  }
}

export function parseMicrosoftExcelLiveCalculationScorecard(value: Record<string, unknown>): MicrosoftExcelLiveCalculationScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const microsoftExcel = objectField(value, 'microsoftExcel')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'microsoft-excel-live-calculation-correctness'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-local-microsoft-excel-automation'),
      appleScriptTransport: literalField(source, 'appleScriptTransport', 'osascript'),
    },
    microsoftExcel: {
      appPath: literalField(microsoftExcel, 'appPath', excelAppPath),
      version: stringField(microsoftExcel, 'version'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      matchingCaseCount: numberField(summary, 'matchingCaseCount'),
      coveredFeatures: stringArrayField(summary, 'coveredFeatures'),
      googleSheetsEvidence: literalField(summary, 'googleSheetsEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseCalculationCase),
  }
}

export function validateMicrosoftExcelLiveCalculationScorecard(scorecard: MicrosoftExcelLiveCalculationScorecard): void {
  if (scorecard.microsoftExcel.version.trim().length === 0) {
    throw new Error('Microsoft Excel live calculation scorecard must record an Excel version')
  }
  if (scorecard.summary.requiredCaseCount !== requiredCaseIds.length) {
    throw new Error('Microsoft Excel live calculation scorecard required case count is stale')
  }
  if (JSON.stringify(scorecard.summary.coveredFeatures) !== JSON.stringify(requiredCoveredFeatures)) {
    throw new Error('Microsoft Excel live calculation scorecard covered features are stale')
  }
  if (JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(requiredCaseIds)) {
    throw new Error('Microsoft Excel live calculation scorecard required cases are stale')
  }
  if (scorecard.summary.matchingCaseCount !== scorecard.cases.filter((entry) => entry.passed).length) {
    throw new Error('Microsoft Excel live calculation scorecard matching case count is inconsistent')
  }
  const currentBiligValuesByCaseId = evaluateBiligCases()
  const failingCases: MicrosoftExcelLiveCalculationCase[] = []
  for (const [index, entry] of scorecard.cases.entries()) {
    const caseSpec = requiredCaseSpecs[index]
    if (caseSpec === undefined) {
      throw new Error(`Microsoft Excel live calculation scorecard has an unexpected case: ${entry.id}`)
    }
    if (entry.id !== caseSpec.id) {
      throw new Error(`Microsoft Excel live calculation case id is stale: ${entry.id}`)
    }
    if (entry.formula !== caseSpec.formula) {
      throw new Error(`Microsoft Excel live calculation formula is stale: ${entry.id}`)
    }
    if (entry.formulaCell !== toA1Address(caseSpec.formulaRowIndex, formulaColumnIndex)) {
      throw new Error(`Microsoft Excel live calculation formula cell is stale: ${entry.id}`)
    }
    if (entry.coveredFeature !== caseSpec.coveredFeature) {
      throw new Error(`Microsoft Excel live calculation covered feature is stale: ${entry.id}`)
    }
    const currentBiligValue = requiredMapValue(currentBiligValuesByCaseId, entry.id, 'Current Bilig calculation value')
    if (!valuesEquivalent(entry.biligValue, currentBiligValue)) {
      throw new Error(
        `Microsoft Excel live calculation Bilig value is stale for ${entry.id}: scorecard=${JSON.stringify(
          entry.biligValue,
        )} current=${JSON.stringify(currentBiligValue)}`,
      )
    }
    const parsedExcelValue = parseExcelRawValue(entry.microsoftExcelRawValue, entry.biligValue)
    if (!valuesEquivalent(entry.microsoftExcelValue, parsedExcelValue)) {
      throw new Error(`Microsoft Excel live calculation parsed value is stale: ${entry.id}`)
    }
    if (entry.passed !== valuesEquivalent(entry.biligValue, entry.microsoftExcelValue)) {
      throw new Error(`Microsoft Excel live calculation pass flag is stale: ${entry.id}`)
    }
    if (!entry.passed) {
      failingCases.push(entry)
    }
    if (entry.formula.trim().length === 0 || !entry.formula.startsWith('=')) {
      throw new Error(`Microsoft Excel live calculation case has an invalid formula: ${entry.id}`)
    }
  }
  if (failingCases.length > 0 || !scorecard.summary.allRequiredCasesPassed) {
    throw new Error(
      `Microsoft Excel live calculation scorecard has failing required cases: ${failingCases
        .map((entry) => `${entry.id} Bilig=${JSON.stringify(entry.biligValue)} Excel=${JSON.stringify(entry.microsoftExcelValue)}`)
        .join(', ')}`,
    )
  }
}

function evaluateLiveExcelCases(): LiveExcelEvaluation {
  if (!existsSync(excelAppPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${excelAppPath}`)
  }
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-live-calculation-'))
  const workbookPath = join(tempDir, 'cases.xlsx')
  const scriptPath = join(tempDir, 'read-cases.scpt')
  try {
    writeFileSync(workbookPath, createExcelWorkbookBytes())
    writeFileSync(scriptPath, createReadCasesAppleScript())
    const rawOutput = execFileSync('osascript', [scriptPath, workbookPath], { encoding: 'utf8' }).trim()
    return parseExcelOutput(rawOutput)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function createExcelWorkbookBytes(): Uint8Array {
  const worksheet = XLSX.utils.aoa_to_sheet(buildCaseRowsForXlsx())
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName)
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

export function createCalculationCaseWorkbookBytes(): Uint8Array {
  return createExcelWorkbookBytes()
}

function createReadCasesAppleScript(): string {
  const addresses = requiredCaseSpecs.map((caseSpec) => `"${toA1Address(caseSpec.formulaRowIndex, formulaColumnIndex)}"`).join(', ')
  const formulaTexts = requiredCaseSpecs.map((caseSpec) => toAppleScriptString(caseSpec.formula)).join(', ')
  return `on run argv
  set workbookPath to POSIX file (item 1 of argv)
  set cellAddresses to {${addresses}}
  set formulaTexts to {${formulaTexts}}
  set output to ""
  tell application "Microsoft Excel"
    set display alerts to false
    open workbook workbook file name workbookPath
    repeat with formulaIndex from 1 to count of cellAddresses
      set cellAddress to item formulaIndex of cellAddresses
      set formulaText to item formulaIndex of formulaTexts
      set formula of range (cellAddress as string) of worksheet "${worksheetName}" of active workbook to formulaText
    end repeat
    calculate full rebuild
    set output to "version=" & (version as string)
    repeat with cellAddress in cellAddresses
      set cellValue to value of range (cellAddress as string) of worksheet "${worksheetName}" of active workbook
      set output to output & linefeed & (cellValue as string)
    end repeat
    close active workbook saving no
  end tell
  return output
end run
`
}

function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}

function parseExcelOutput(rawOutput: string): LiveExcelEvaluation {
  const lines = rawOutput.split(/\r?\n/u)
  const versionLine = lines[0]
  if (!versionLine?.startsWith('version=')) {
    throw new Error(`Unexpected Microsoft Excel output header: ${versionLine ?? '<empty>'}`)
  }
  const rawValues = lines.slice(1)
  if (rawValues.length !== requiredCaseSpecs.length) {
    throw new Error(`Expected ${String(requiredCaseSpecs.length)} Excel values, received ${String(rawValues.length)}`)
  }
  return {
    excelVersion: versionLine.slice('version='.length),
    rawValuesByCaseId: new Map(requiredCaseSpecs.map((caseSpec, index) => [caseSpec.id, rawValues[index] ?? ''])),
  }
}

function buildCaseRowsForWorkPaper(): Array<Array<boolean | number | string | null>> {
  return buildCaseRows((formula) => formula)
}

function buildCaseRowsForXlsx(): Array<Array<boolean | number | string | null>> {
  return buildCaseRows(() => null)
}

function buildCaseRows(formulaCellValue: (formula: string) => string | null): Array<Array<boolean | number | string | null>> {
  const height = Math.max(...requiredCaseSpecs.map((entry) => entry.formulaRowIndex), ...sharedRows.keys()) + 1
  const rows = Array.from({ length: height }, () => [null, null, null, null] as Array<boolean | number | string | null>)
  for (const caseSpec of requiredCaseSpecs) {
    const row = rows[caseSpec.formulaRowIndex]
    for (let col = 0; col < caseSpec.inputValues.length; col += 1) {
      row[col] = caseSpec.inputValues[col] ?? null
    }
    row[formulaColumnIndex] = formulaCellValue(caseSpec.formula)
  }
  for (const [rowIndex, values] of sharedRows) {
    const row = rows[rowIndex]
    for (let col = 0; col < values.length; col += 1) {
      row[col] = values[col] ?? null
    }
  }
  return rows
}

function parseCalculationCase(value: unknown): MicrosoftExcelLiveCalculationCase {
  const record = asObject(value, 'Microsoft Excel live calculation case')
  return {
    id: stringField(record, 'id'),
    formula: stringField(record, 'formula'),
    formulaCell: stringField(record, 'formulaCell'),
    coveredFeature: stringField(record, 'coveredFeature'),
    biligValue: parseScalarValue(record['biligValue'], 'biligValue'),
    microsoftExcelRawValue: stringField(record, 'microsoftExcelRawValue'),
    microsoftExcelValue: parseScalarValue(record['microsoftExcelValue'], 'microsoftExcelValue'),
    passed: booleanField(record, 'passed'),
  }
}

function parseScalarValue(value: unknown, name: string): CalculationScalarValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
    return value
  }
  const record = asObject(value, name)
  return { error: stringField(record, 'error') }
}

function parseExcelRawValue(rawValue: string, expectedBiligValue: CalculationScalarValue): CalculationScalarValue {
  if (expectedBiligValue === null) {
    return rawValue.length === 0 ? null : rawValue
  }
  if (typeof expectedBiligValue === 'number') {
    const numericValue = Number(rawValue)
    if (!Number.isFinite(numericValue)) {
      return { error: `NON_NUMERIC_EXCEL_VALUE_${rawValue}` }
    }
    return numericValue
  }
  if (typeof expectedBiligValue === 'boolean') {
    const normalized = rawValue.toLowerCase()
    if (normalized === 'true') {
      return true
    }
    if (normalized === 'false') {
      return false
    }
    return { error: `NON_BOOLEAN_EXCEL_VALUE_${rawValue}` }
  }
  if (typeof expectedBiligValue === 'string') {
    return rawValue
  }
  return { error: rawValue }
}

export function parseCalculationRawValue(rawValue: string, expectedBiligValue: CalculationScalarValue): CalculationScalarValue {
  return parseExcelRawValue(rawValue, expectedBiligValue)
}

function normalizeBiligValue(value: unknown): CalculationScalarValue {
  if (!isProtocolValueLike(value)) {
    return { error: 'UNKNOWN_BILIG_VALUE' }
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return { error: formatUnknownErrorCode(value.code) }
    default:
      return { error: `UNKNOWN_BILIG_TAG_${String(value.tag)}` }
  }
}

function isProtocolValueLike(value: unknown): value is { code?: unknown; tag: ValueTag; value?: boolean | number | string } {
  if (value === null || typeof value !== 'object') {
    return false
  }
  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function formatUnknownErrorCode(value: unknown): string {
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  if (value === null || value === undefined) {
    return 'ERROR'
  }
  return JSON.stringify(value) ?? 'ERROR'
}

function valuesEquivalent(left: CalculationScalarValue, right: CalculationScalarValue): boolean {
  if (typeof left === 'number' && typeof right === 'number') {
    return Math.abs(left - right) <= 1e-9
  }
  return JSON.stringify(left) === JSON.stringify(right)
}

export function calculationValuesEquivalent(left: CalculationScalarValue, right: CalculationScalarValue): boolean {
  return valuesEquivalent(left, right)
}

function requiredMapValue<T>(map: ReadonlyMap<string, T>, key: string, label: string): T {
  const value = map.get(key)
  if (value === undefined) {
    throw new Error(`${label} is missing required case: ${key}`)
  }
  return value
}

function toA1Address(rowIndex: number, columnIndex: number): string {
  return `${columnName(columnIndex)}${String(rowIndex + 1)}`
}

export function calculationLiveFormulaCell(caseSpec: Pick<CalculationCaseSpec, 'formulaRowIndex'>): string {
  return toA1Address(caseSpec.formulaRowIndex, formulaColumnIndex)
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

function logResult(mode: 'check' | 'write', scorecard: MicrosoftExcelLiveCalculationScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        excelVersion: scorecard.microsoftExcel.version,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        matchingCaseCount: scorecard.summary.matchingCaseCount,
        requiredCaseCount: scorecard.summary.requiredCaseCount,
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
