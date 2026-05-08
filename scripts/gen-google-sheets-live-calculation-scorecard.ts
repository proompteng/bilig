#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import * as XLSX from 'xlsx'
import {
  calculationLiveCaseSpecs,
  calculationLiveFormulaCell,
  calculationLiveRequiredCaseIds,
  calculationLiveRequiredCoveredFeatures,
  calculationLiveWorksheetName,
  calculationValuesEquivalent,
  createCalculationCaseWorkbookBytes,
  evaluateBiligCases,
  parseCalculationRawValue,
  type CalculationScalarValue,
} from './gen-microsoft-excel-live-calculation-scorecard.ts'
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

export interface GoogleSheetsLiveCalculationCapture {
  readonly generatedAt: string
  readonly googleSheets: {
    readonly spreadsheetId: string
    readonly spreadsheetUrl: string
    readonly title: string
  }
  readonly capture: {
    readonly transport: 'google-drive-connector'
    readonly sourceWorkbook: 'xlsx-native-google-sheets-conversion'
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly worksheetName: string
  }
  readonly rawValuesByCaseId: Record<string, string>
}

export interface GoogleSheetsLiveCalculationCase {
  readonly id: string
  readonly formula: string
  readonly formulaCell: string
  readonly coveredFeature: string
  readonly biligValue: CalculationScalarValue
  readonly googleSheetsRawValue: string
  readonly googleSheetsValue: CalculationScalarValue
  readonly passed: boolean
}

export interface GoogleSheetsLiveCalculationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'google-sheets-live-calculation-correctness'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-google-sheets-live-calculation-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector'
    readonly captureTransport: 'google-drive-connector'
  }
  readonly googleSheets: {
    readonly spreadsheetId: string
    readonly spreadsheetUrl: string
    readonly title: string
    readonly worksheetName: string
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly matchingCaseCount: number
    readonly coveredFeatures: string[]
    readonly microsoftExcelEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: GoogleSheetsLiveCalculationCase[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'google-sheets-live-calculation-scorecard.json')

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  const emitXlsxIndex = process.argv.indexOf('--emit-xlsx')
  const captureIndex = process.argv.indexOf('--capture')

  if (emitXlsxIndex >= 0) {
    const targetPath = process.argv[emitXlsxIndex + 1]
    if (!targetPath) {
      throw new Error('Missing path after --emit-xlsx')
    }
    writeFileSync(targetPath, createGoogleSheetsFormulaWorkbookBytes())
    console.log(JSON.stringify({ mode: 'emit-xlsx', outputPath: targetPath, worksheetName: calculationLiveWorksheetName }, null, 2))
    return
  }

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Google Sheets live calculation scorecard is missing. Run: bun scripts/gen-google-sheets-live-calculation-scorecard.ts --capture <capture.json>`,
      )
    }
    const scorecard = parseGoogleSheetsLiveCalculationScorecard(readJsonObject(outputPath))
    validateGoogleSheetsLiveCalculationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  if (captureIndex < 0) {
    throw new Error(
      'Missing --capture <capture.json>. First emit an XLSX with --emit-xlsx, import it as native Google Sheets, then capture UNFORMATTED_VALUE formula cells through the Google Drive connector.',
    )
  }
  const capturePath = process.argv[captureIndex + 1]
  if (!capturePath) {
    throw new Error('Missing path after --capture')
  }
  const scorecard = buildGoogleSheetsLiveCalculationScorecard(parseGoogleSheetsLiveCalculationCapture(readJsonObject(capturePath)))
  validateGoogleSheetsLiveCalculationScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

function createGoogleSheetsFormulaWorkbookBytes(): Uint8Array {
  const workbook = XLSX.read(createCalculationCaseWorkbookBytes(), { type: 'buffer' })
  const worksheet = workbook.Sheets[calculationLiveWorksheetName]
  if (!worksheet) {
    throw new Error(`Missing worksheet in Google Sheets calculation workbook: ${calculationLiveWorksheetName}`)
  }
  for (const caseSpec of calculationLiveCaseSpecs) {
    worksheet[calculationLiveFormulaCell(caseSpec)] = {
      t: 'n',
      f: caseSpec.formula.replace(/^=/u, ''),
    }
  }
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

export function buildGoogleSheetsLiveCalculationScorecard(
  capture: GoogleSheetsLiveCalculationCapture,
): GoogleSheetsLiveCalculationScorecard {
  const biligValuesByCaseId = evaluateBiligCases()
  const cases = calculationLiveCaseSpecs.map((caseSpec) => {
    const biligValue = requiredMapValue(biligValuesByCaseId, caseSpec.id, 'Bilig calculation value')
    const googleSheetsRawValue = requiredRecordValue(capture.rawValuesByCaseId, caseSpec.id, 'Google Sheets calculation value')
    const googleSheetsValue = parseCalculationRawValue(googleSheetsRawValue, biligValue)
    return {
      id: caseSpec.id,
      formula: caseSpec.formula,
      formulaCell: calculationLiveFormulaCell(caseSpec),
      coveredFeature: caseSpec.coveredFeature,
      biligValue,
      googleSheetsRawValue,
      googleSheetsValue,
      passed: calculationValuesEquivalent(biligValue, googleSheetsValue),
    }
  })
  const matchingCaseCount = cases.filter((entry) => entry.passed).length

  return {
    schemaVersion: 1,
    suite: 'google-sheets-live-calculation-correctness',
    generatedAt: capture.generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-google-sheets-live-calculation-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
      captureTransport: capture.capture.transport,
    },
    googleSheets: {
      spreadsheetId: capture.googleSheets.spreadsheetId,
      spreadsheetUrl: capture.googleSheets.spreadsheetUrl,
      title: capture.googleSheets.title,
      worksheetName: capture.capture.worksheetName,
      valueRenderOption: capture.capture.valueRenderOption,
    },
    summary: {
      allRequiredCasesPassed: matchingCaseCount === cases.length,
      requiredCaseCount: cases.length,
      matchingCaseCount,
      coveredFeatures: [...calculationLiveRequiredCoveredFeatures],
      microsoftExcelEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseGoogleSheetsLiveCalculationCapture(value: Record<string, unknown>): GoogleSheetsLiveCalculationCapture {
  const googleSheets = objectField(value, 'googleSheets')
  const capture = objectField(value, 'capture')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    googleSheets: {
      spreadsheetId: stringField(googleSheets, 'spreadsheetId'),
      spreadsheetUrl: stringField(googleSheets, 'spreadsheetUrl'),
      title: stringField(googleSheets, 'title'),
    },
    capture: {
      transport: literalField(capture, 'transport', 'google-drive-connector'),
      sourceWorkbook: literalField(capture, 'sourceWorkbook', 'xlsx-native-google-sheets-conversion'),
      valueRenderOption: literalField(capture, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      worksheetName: stringField(capture, 'worksheetName'),
    },
    rawValuesByCaseId: parseRawValuesByCaseId(objectField(value, 'rawValuesByCaseId')),
  }
}

export function parseGoogleSheetsLiveCalculationScorecard(value: Record<string, unknown>): GoogleSheetsLiveCalculationScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const googleSheets = objectField(value, 'googleSheets')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'google-sheets-live-calculation-correctness'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-google-sheets-live-calculation-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-google-sheets-native-conversion-via-google-drive-connector'),
      captureTransport: literalField(source, 'captureTransport', 'google-drive-connector'),
    },
    googleSheets: {
      spreadsheetId: stringField(googleSheets, 'spreadsheetId'),
      spreadsheetUrl: stringField(googleSheets, 'spreadsheetUrl'),
      title: stringField(googleSheets, 'title'),
      worksheetName: stringField(googleSheets, 'worksheetName'),
      valueRenderOption: literalField(googleSheets, 'valueRenderOption', 'UNFORMATTED_VALUE'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      matchingCaseCount: numberField(summary, 'matchingCaseCount'),
      coveredFeatures: stringArrayField(summary, 'coveredFeatures'),
      microsoftExcelEvidence: literalField(summary, 'microsoftExcelEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseCalculationCase),
  }
}

export function validateGoogleSheetsLiveCalculationScorecard(scorecard: GoogleSheetsLiveCalculationScorecard): void {
  if (scorecard.googleSheets.spreadsheetId.trim().length === 0 || scorecard.googleSheets.spreadsheetUrl.trim().length === 0) {
    throw new Error('Google Sheets live calculation scorecard must record a spreadsheet id and URL')
  }
  if (scorecard.googleSheets.worksheetName !== calculationLiveWorksheetName) {
    throw new Error('Google Sheets live calculation scorecard worksheet name is stale')
  }
  if (scorecard.summary.requiredCaseCount !== calculationLiveRequiredCaseIds.length) {
    throw new Error('Google Sheets live calculation scorecard required case count is stale')
  }
  if (JSON.stringify(scorecard.summary.coveredFeatures) !== JSON.stringify(calculationLiveRequiredCoveredFeatures)) {
    throw new Error('Google Sheets live calculation scorecard covered features are stale')
  }
  if (JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(calculationLiveRequiredCaseIds)) {
    throw new Error('Google Sheets live calculation scorecard required cases are stale')
  }
  if (scorecard.summary.matchingCaseCount !== scorecard.cases.filter((entry) => entry.passed).length) {
    throw new Error('Google Sheets live calculation scorecard matching case count is inconsistent')
  }
  const currentBiligValuesByCaseId = evaluateBiligCases()
  const failingCases: GoogleSheetsLiveCalculationCase[] = []
  for (const [index, entry] of scorecard.cases.entries()) {
    const caseSpec = calculationLiveCaseSpecs[index]
    if (caseSpec === undefined) {
      throw new Error(`Google Sheets live calculation scorecard has an unexpected case: ${entry.id}`)
    }
    if (entry.id !== caseSpec.id) {
      throw new Error(`Google Sheets live calculation case id is stale: ${entry.id}`)
    }
    if (entry.formula !== caseSpec.formula) {
      throw new Error(`Google Sheets live calculation formula is stale: ${entry.id}`)
    }
    if (entry.formulaCell !== calculationLiveFormulaCell(caseSpec)) {
      throw new Error(`Google Sheets live calculation formula cell is stale: ${entry.id}`)
    }
    if (entry.coveredFeature !== caseSpec.coveredFeature) {
      throw new Error(`Google Sheets live calculation covered feature is stale: ${entry.id}`)
    }
    const currentBiligValue = requiredMapValue(currentBiligValuesByCaseId, entry.id, 'Current Bilig calculation value')
    if (!calculationValuesEquivalent(entry.biligValue, currentBiligValue)) {
      throw new Error(
        `Google Sheets live calculation Bilig value is stale for ${entry.id}: scorecard=${JSON.stringify(
          entry.biligValue,
        )} current=${JSON.stringify(currentBiligValue)}`,
      )
    }
    const parsedGoogleSheetsValue = parseCalculationRawValue(entry.googleSheetsRawValue, entry.biligValue)
    if (!calculationValuesEquivalent(entry.googleSheetsValue, parsedGoogleSheetsValue)) {
      throw new Error(`Google Sheets live calculation parsed value is stale: ${entry.id}`)
    }
    if (entry.passed !== calculationValuesEquivalent(entry.biligValue, entry.googleSheetsValue)) {
      throw new Error(`Google Sheets live calculation pass flag is stale: ${entry.id}`)
    }
    if (!entry.passed) {
      failingCases.push(entry)
    }
    if (entry.formula.trim().length === 0 || !entry.formula.startsWith('=')) {
      throw new Error(`Google Sheets live calculation case has an invalid formula: ${entry.id}`)
    }
  }
  if (failingCases.length > 0 || !scorecard.summary.allRequiredCasesPassed) {
    throw new Error(
      `Google Sheets live calculation scorecard has failing required cases: ${failingCases
        .map((entry) => `${entry.id} Bilig=${JSON.stringify(entry.biligValue)} GoogleSheets=${JSON.stringify(entry.googleSheetsValue)}`)
        .join(', ')}`,
    )
  }
}

function parseCalculationCase(value: unknown): GoogleSheetsLiveCalculationCase {
  const record = asObject(value, 'Google Sheets live calculation case')
  return {
    id: stringField(record, 'id'),
    formula: stringField(record, 'formula'),
    formulaCell: stringField(record, 'formulaCell'),
    coveredFeature: stringField(record, 'coveredFeature'),
    biligValue: parseScalarValue(record['biligValue'], 'biligValue'),
    googleSheetsRawValue: stringField(record, 'googleSheetsRawValue'),
    googleSheetsValue: parseScalarValue(record['googleSheetsValue'], 'googleSheetsValue'),
    passed: booleanField(record, 'passed'),
  }
}

function parseRawValuesByCaseId(value: Record<string, unknown>): Record<string, string> {
  const result: Record<string, string> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (typeof entry === 'string') {
      result[key] = entry
      continue
    }
    if (typeof entry === 'number' || typeof entry === 'boolean') {
      result[key] = String(entry)
      continue
    }
    if (entry === null) {
      result[key] = ''
      continue
    }
    throw new Error(`Unsupported Google Sheets captured value for ${key}`)
  }
  return result
}

function parseScalarValue(value: unknown, name: string): CalculationScalarValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
    return value
  }
  const record = asObject(value, name)
  return { error: stringField(record, 'error') }
}

function requiredMapValue<T>(map: ReadonlyMap<string, T>, key: string, label: string): T {
  const value = map.get(key)
  if (value === undefined) {
    throw new Error(`${label} is missing required case: ${key}`)
  }
  return value
}

function requiredRecordValue(record: Record<string, string>, key: string, label: string): string {
  const value = record[key]
  if (value === undefined) {
    throw new Error(`${label} is missing required case: ${key}`)
  }
  return value
}

function logResult(mode: 'check' | 'write', scorecard: GoogleSheetsLiveCalculationScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        spreadsheetId: scorecard.googleSheets.spreadsheetId,
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
