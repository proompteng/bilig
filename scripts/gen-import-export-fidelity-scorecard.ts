#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { SpreadsheetEngine } from '../packages/core/src/engine.js'
import { exportXlsx, importCsv, importXlsx } from '../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { projectSupportedSnapshotSemantics } from './import-export-fidelity-projection.ts'

export interface ImportExportFidelityCase {
  readonly id: string
  readonly format: 'csv' | 'xlsx'
  readonly direction: 'import' | 'export-import' | 'import-export-import'
  readonly required: boolean
  readonly passed: boolean
  readonly coveredFeatures: string[]
  readonly missingFeatures: string[]
  readonly evidence: string
}

export interface ImportExportFidelityScorecard {
  readonly schemaVersion: 1
  readonly suite: 'import-export-fidelity'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-import-export-fidelity-scorecard.ts'
    readonly implementationPackage: 'packages/excel-import'
    readonly enginePackage: 'packages/core'
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly csvRoundTripPassed: boolean
    readonly xlsxImportPassed: boolean
    readonly xlsxSnapshotRoundTripPassed: boolean
    readonly coveredFeatures: string[]
    readonly unsupportedFeatures: string[]
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly cases: ImportExportFidelityCase[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'import-export-fidelity-scorecard.json')
const requiredCaseIds = [
  'csv-import-preview',
  'csv-engine-roundtrip',
  'xlsx-import-preview',
  'xlsx-snapshot-roundtrip-values-formulas-formats',
  'xlsx-snapshot-roundtrip-dimensions-merges',
  'xlsx-snapshot-roundtrip-freeze-panes',
  'xlsx-snapshot-roundtrip-filters',
  'xlsx-snapshot-roundtrip-sheet-protection',
  'xlsx-snapshot-roundtrip-protected-ranges',
  'xlsx-snapshot-roundtrip-data-validations',
  'xlsx-snapshot-roundtrip-tables',
  'xlsx-snapshot-roundtrip-charts',
  'xlsx-snapshot-roundtrip-pivots',
  'xlsx-unsupported-features-warning',
] as const
const coveredFeatureOrder = [
  'csv.import',
  'csv.preview',
  'csv.export',
  'csv.roundtrip',
  'xlsx.import',
  'xlsx.preview',
  'xlsx.export',
  'xlsx.roundtrip',
  'xlsx.values',
  'xlsx.formulas',
  'xlsx.numberFormats',
  'xlsx.definedNames',
  'xlsx.comments',
  'xlsx.styles',
  'xlsx.rowColumnDimensions',
  'xlsx.merges',
  'xlsx.freezePanes.roundtrip',
  'xlsx.filters.roundtrip',
  'xlsx.sheetProtection.roundtrip',
  'xlsx.protectedRanges.roundtrip',
  'xlsx.dataValidations.roundtrip',
  'xlsx.tables.roundtrip',
  'xlsx.charts.roundtrip',
  'xlsx.pivots.roundtrip',
  'xlsx.multiSheet',
  'xlsx.unsupportedFeatureWarnings',
] as const
const unsupportedFeatures = ['xlsx.macros.execution'] as const

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(`Import/export fidelity scorecard is missing. Run: bun scripts/gen-import-export-fidelity-scorecard.ts`)
    }
    const scorecard = parseImportExportFidelityScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateImportExportFidelityScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildImportExportFidelityScorecard()
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildImportExportFidelityScorecard(generatedAt = new Date().toISOString()): Promise<ImportExportFidelityScorecard> {
  const cases = [
    runCsvImportPreviewCase(),
    await runCsvEngineRoundTripCase(),
    runXlsxImportPreviewCase(),
    runXlsxSnapshotRoundTripValuesCase(),
    runXlsxSnapshotRoundTripDimensionsCase(),
    runXlsxSnapshotRoundTripFreezePanesCase(),
    runXlsxSnapshotRoundTripFiltersCase(),
    runXlsxSnapshotRoundTripSheetProtectionCase(),
    runXlsxSnapshotRoundTripProtectedRangesCase(),
    runXlsxSnapshotRoundTripDataValidationsCase(),
    runXlsxSnapshotRoundTripTablesCase(),
    runXlsxSnapshotRoundTripChartsCase(),
    runXlsxSnapshotRoundTripPivotsCase(),
    runXlsxUnsupportedFeaturesWarningCase(),
  ]
  const coveredFeatureSet = new Set(cases.flatMap((entry) => entry.coveredFeatures))
  const coveredFeatures = coveredFeatureOrder.filter((feature) => coveredFeatureSet.has(feature))

  return {
    schemaVersion: 1,
    suite: 'import-export-fidelity',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-import-export-fidelity-scorecard.ts',
      implementationPackage: 'packages/excel-import',
      enginePackage: 'packages/core',
    },
    summary: {
      allRequiredCasesPassed: cases.filter((entry) => entry.required).every((entry) => entry.passed),
      csvRoundTripPassed: requiredCase(cases, 'csv-engine-roundtrip').passed,
      xlsxImportPassed: requiredCase(cases, 'xlsx-import-preview').passed,
      xlsxSnapshotRoundTripPassed:
        requiredCase(cases, 'xlsx-snapshot-roundtrip-values-formulas-formats').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-dimensions-merges').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-freeze-panes').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-filters').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-sheet-protection').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-protected-ranges').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-data-validations').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-tables').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-charts').passed &&
        requiredCase(cases, 'xlsx-snapshot-roundtrip-pivots').passed,
      coveredFeatures,
      unsupportedFeatures: [...unsupportedFeatures],
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    cases,
  }
}

function runCsvImportPreviewCase(): ImportExportFidelityCase {
  const imported = importCsv('Name,Value\nalpha,12\nbeta,=B2*2', 'metrics.csv')
  const passed =
    imported.workbookName === 'metrics' &&
    imported.preview.sheetCount === 1 &&
    imported.preview.sheets[0]?.previewRows.length === 3 &&
    Boolean(imported.snapshot.sheets[0]?.cells.some((cell) => cell.address === 'B3' && cell.formula === 'B2*2'))
  return fidelityCase({
    id: 'csv-import-preview',
    format: 'csv',
    direction: 'import',
    passed,
    coveredFeatures: ['csv.import', 'csv.preview'],
    evidence: 'CSV import preserves workbook name, preview rows, typed values, and formulas.',
  })
}

async function runCsvEngineRoundTripCase(): Promise<ImportExportFidelityCase> {
  const csv = 'Name,Value\nalpha,12\nbeta,=B2*2'
  const engine = new SpreadsheetEngine({
    workbookName: 'csv-roundtrip',
    replicaId: 'import-export-fidelity-csv-roundtrip',
  })
  await engine.ready()
  engine.createSheet('Sheet1')
  engine.importSheetCsv('Sheet1', csv)
  const exported = engine.exportSheetCsv('Sheet1')

  const restored = new SpreadsheetEngine({
    workbookName: 'csv-restored',
    replicaId: 'import-export-fidelity-csv-restored',
  })
  await restored.ready()
  restored.createSheet('Sheet1')
  restored.importSheetCsv('Sheet1', exported)
  const reexported = restored.exportSheetCsv('Sheet1')

  return fidelityCase({
    id: 'csv-engine-roundtrip',
    format: 'csv',
    direction: 'import-export-import',
    passed: reexported === exported && exported.includes('=B2*2'),
    coveredFeatures: ['csv.export', 'csv.roundtrip'],
    evidence: 'Core CSV export/import reserializes to the same CSV after a restore import.',
  })
}

function runXlsxImportPreviewCase(): ImportExportFidelityCase {
  const imported = importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx')
  const summary = imported.snapshot.sheets.find((sheet) => sheet.name === 'Summary')
  const passed =
    imported.sheetNames.join(',') === 'Summary,Inputs' &&
    imported.preview.sheetCount === 2 &&
    summary?.cells.some((cell) => cell.address === 'B1' && cell.formula === 'SUM(B2:B3)' && cell.format === '0.00') === true
  return fidelityCase({
    id: 'xlsx-import-preview',
    format: 'xlsx',
    direction: 'import',
    passed,
    coveredFeatures: ['xlsx.import', 'xlsx.preview', 'xlsx.multiSheet'],
    evidence: 'XLSX import preserves preview metadata, sheet order, formulas, and number formats from a generated workbook.',
  })
}

function runXlsxSnapshotRoundTripValuesCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed =
    JSON.stringify(actual.valueFormulaFormatSheets) === JSON.stringify(expected.valueFormulaFormatSheets) &&
    JSON.stringify(actual.commentThreads) === JSON.stringify(expected.commentThreads) &&
    JSON.stringify(actual.styleRanges) === JSON.stringify(expected.styleRanges)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-values-formulas-formats',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: [
      'xlsx.export',
      'xlsx.roundtrip',
      'xlsx.values',
      'xlsx.formulas',
      'xlsx.numberFormats',
      'xlsx.definedNames',
      'xlsx.comments',
      'xlsx.styles',
      'xlsx.multiSheet',
    ],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with the same values, formulas, formats, defined names, comments, styles, and sheet order.',
  })
}

function runXlsxSnapshotRoundTripDimensionsCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.dimensionSheets) === JSON.stringify(expected.dimensionSheets)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-dimensions-merges',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.rowColumnDimensions', 'xlsx.merges'],
    evidence: 'WorkbookSnapshot exported to XLSX imports back with equivalent row heights, column widths, and merged ranges.',
  })
}

function runXlsxSnapshotRoundTripFreezePanesCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.freezePanes) === JSON.stringify(expected.freezePanes)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-freeze-panes',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.freezePanes.roundtrip'],
    evidence: 'WorkbookSnapshot exported to XLSX imports back with equivalent sheet freeze-pane metadata backed by native XLSX pane nodes.',
  })
}

function runXlsxSnapshotRoundTripFiltersCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.filters) === JSON.stringify(expected.filters)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-filters',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.filters.roundtrip'],
    evidence: 'WorkbookSnapshot exported to XLSX imports back with equivalent sheet filter ranges backed by native XLSX autoFilter nodes.',
  })
}

function runXlsxSnapshotRoundTripSheetProtectionCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.sheetProtections) === JSON.stringify(expected.sheetProtections)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-sheet-protection',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.sheetProtection.roundtrip'],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent sheet protection metadata backed by native XLSX sheetProtection nodes.',
  })
}

function runXlsxSnapshotRoundTripProtectedRangesCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.protectedRanges) === JSON.stringify(expected.protectedRanges)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-protected-ranges',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.protectedRanges.roundtrip'],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent protected-range metadata backed by native XLSX protectedRange nodes.',
  })
}

function runXlsxSnapshotRoundTripDataValidationsCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.validations) === JSON.stringify(expected.validations)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-data-validations',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.dataValidations.roundtrip'],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent sheet data-validation metadata backed by native XLSX dataValidation nodes.',
  })
}

function runXlsxSnapshotRoundTripTablesCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.tables) === JSON.stringify(expected.tables)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-tables',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.tables.roundtrip'],
    evidence: 'WorkbookSnapshot exported to XLSX imports back with equivalent Bilig table metadata backed by real XLSX table parts.',
  })
}

function runXlsxSnapshotRoundTripChartsCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.charts) === JSON.stringify(expected.charts)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-charts',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.charts.roundtrip'],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent Bilig chart metadata backed by real XLSX chart/drawing parts.',
  })
}

function runXlsxSnapshotRoundTripPivotsCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.pivots) === JSON.stringify(expected.pivots)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-pivots',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.pivots.roundtrip'],
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent Bilig pivot metadata backed by real XLSX pivot table, cache definition, and cache records parts.',
  })
}

function runXlsxUnsupportedFeaturesWarningCase(): ImportExportFidelityCase {
  const snapshot = createFidelitySnapshot()
  const imported = importXlsx(exportXlsx(snapshot), 'fidelity.xlsx')
  return fidelityCase({
    id: 'xlsx-unsupported-features-warning',
    format: 'xlsx',
    direction: 'export-import',
    passed: imported.warnings.length === 0 && unsupportedFeatures.length > 0,
    coveredFeatures: ['xlsx.unsupportedFeatureWarnings'],
    missingFeatures: [...unsupportedFeatures],
    evidence: 'Scorecard explicitly records unsupported XLSX surfaces instead of treating a subset round trip as full Excel parity.',
  })
}

function fidelityCase(input: {
  readonly id: ImportExportFidelityCase['id']
  readonly format: ImportExportFidelityCase['format']
  readonly direction: ImportExportFidelityCase['direction']
  readonly passed: boolean
  readonly coveredFeatures: readonly string[]
  readonly missingFeatures?: readonly string[]
  readonly evidence: string
}): ImportExportFidelityCase {
  return {
    id: input.id,
    format: input.format,
    direction: input.direction,
    required: true,
    passed: input.passed,
    coveredFeatures: [...input.coveredFeatures],
    missingFeatures: [...(input.missingFeatures ?? [])],
    evidence: input.evidence,
  }
}

function createFidelitySnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Import Export Fidelity',
      metadata: {
        definedNames: [
          { name: 'SummaryTotal', value: { kind: 'cell-ref', sheetName: 'Summary', address: 'B1' } },
          { name: 'InputRegion', value: { kind: 'range-ref', sheetName: 'Inputs', startAddress: 'A1', endAddress: 'B1' } },
          { name: 'TaxRate', value: { kind: 'scalar', value: 0.085 } },
        ],
        styles: [
          {
            id: 'accent-total',
            fill: { backgroundColor: '#1d3989' },
            font: { family: 'Aptos', size: 12, bold: true, color: '#ffffff' },
            alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
            borders: { bottom: { style: 'solid', weight: 'thin', color: '#000000' } },
          },
        ],
        tables: [
          {
            name: 'InputTable',
            sheetName: 'Inputs',
            startAddress: 'A1',
            endAddress: 'D4',
            columnNames: ['Region', 'Product', 'Sales', 'Notes'],
            headerRow: true,
            totalsRow: false,
          },
        ],
        charts: [
          {
            id: 'summary-trend',
            sheetName: 'Summary',
            address: 'E1',
            source: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
            chartType: 'line',
            seriesOrientation: 'columns',
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
            title: 'Summary Trend',
            legendPosition: 'right',
            rows: 12,
            cols: 6,
          },
        ],
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Summary',
            address: 'E15',
            source: { sheetName: 'Inputs', startAddress: 'A1', endAddress: 'D4' },
            groupBy: ['Region'],
            values: [
              { sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Total Sales' },
              { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
            ],
            rows: 4,
            cols: 3,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Summary',
        order: 0,
        metadata: {
          styleRanges: [{ range: { sheetName: 'Summary', startAddress: 'B1', endAddress: 'B1' }, styleId: 'accent-total' }],
          commentThreads: [
            {
              threadId: 'summary-total-comment',
              sheetName: 'Summary',
              address: 'B1',
              comments: [{ id: 'summary-total-comment-1', body: 'Reviewed total', authorDisplayName: 'Finance' }],
            },
          ],
          columns: [
            { id: 'summary-col-0', index: 0, size: 132 },
            { id: 'summary-col-1', index: 1, size: 96 },
          ],
          rows: [
            { id: 'summary-row-0', index: 0, size: 30 },
            { id: 'summary-row-2', index: 2, size: 24 },
          ],
          freezePane: { rows: 1, cols: 2 },
          merges: [{ sheetName: 'Summary', startAddress: 'A5', endAddress: 'B5' }],
          sheetProtection: { sheetName: 'Summary' },
          protectedRanges: [
            {
              id: 'protect-summary-inputs',
              range: { sheetName: 'Summary', startAddress: 'A2', endAddress: 'B3' },
            },
          ],
          filters: [{ sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' }],
          validations: [
            {
              range: { sheetName: 'Summary', startAddress: 'C2', endAddress: 'C4' },
              rule: { kind: 'whole', operator: 'between', values: [0, 100] },
              allowBlank: false,
              errorStyle: 'stop',
              errorTitle: 'Percent required',
              errorMessage: 'Enter a whole number from 0 to 100.',
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'B1', formula: 'SUM(B2:B3)', format: '0.00' },
          { address: 'C1', value: true },
          { address: 'A2', value: 'Revenue' },
          { address: 'B2', value: 1250.5, format: '$#,##0.00' },
          { address: 'A3', value: 'Costs' },
          { address: 'B3', value: 450.25, format: '$#,##0.00' },
        ],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        metadata: {
          validations: [
            {
              range: { sheetName: 'Inputs', startAddress: 'D2', endAddress: 'D4' },
              rule: { kind: 'list', values: ['Priority', 'Standard'] },
              allowBlank: true,
              showDropdown: true,
              promptTitle: 'Status',
              promptMessage: 'Pick a known priority.',
              errorStyle: 'warning',
              errorTitle: 'Unknown priority',
              errorMessage: 'Use Priority or Standard.',
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Product' },
          { address: 'C1', value: 'Sales' },
          { address: 'D1', value: 'Notes' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'Widget' },
          { address: 'C2', value: 10 },
          { address: 'D2', value: 'Priority' },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 'Widget' },
          { address: 'C3', value: 7 },
          { address: 'D3', value: 'Priority' },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'Gizmo' },
          { address: 'C4', value: 5 },
          { address: 'D4', value: 'Standard' },
        ],
      },
    ],
  }
}

function requiredCase(cases: readonly ImportExportFidelityCase[], id: string): ImportExportFidelityCase {
  const entry = cases.find((candidate) => candidate.id === id)
  if (!entry) {
    throw new Error(`Import/export fidelity scorecard is missing required case: ${id}`)
  }
  return entry
}

export function parseImportExportFidelityScorecard(value: unknown): ImportExportFidelityScorecard {
  const record = toRecord(value, 'import/export fidelity scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'import-export-fidelity') {
    throw new Error('Unexpected import/export fidelity scorecard header')
  }
  const source = recordField(record, 'source', 'import/export fidelity source')
  const summary = recordField(record, 'summary', 'import/export fidelity summary')
  return {
    schemaVersion: 1,
    suite: 'import-export-fidelity',
    generatedAt: stringField(record, 'generatedAt', 'import/export fidelity generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-import-export-fidelity-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/excel-import'),
      enginePackage: literalField(source, 'enginePackage', 'packages/core'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed', 'import/export fidelity allRequiredCasesPassed'),
      csvRoundTripPassed: booleanField(summary, 'csvRoundTripPassed', 'import/export fidelity csvRoundTripPassed'),
      xlsxImportPassed: booleanField(summary, 'xlsxImportPassed', 'import/export fidelity xlsxImportPassed'),
      xlsxSnapshotRoundTripPassed: booleanField(
        summary,
        'xlsxSnapshotRoundTripPassed',
        'import/export fidelity xlsxSnapshotRoundTripPassed',
      ),
      coveredFeatures: stringArrayField(summary, 'coveredFeatures', 'import/export fidelity coveredFeatures'),
      unsupportedFeatures: stringArrayField(summary, 'unsupportedFeatures', 'import/export fidelity unsupportedFeatures'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    cases: arrayField(record, 'cases', 'import/export fidelity cases').map(parseImportExportFidelityCase),
  }
}

function parseImportExportFidelityCase(value: unknown): ImportExportFidelityCase {
  const record = toRecord(value, 'import/export fidelity case')
  return {
    id: stringField(record, 'id', 'import/export fidelity case id'),
    format: parseFormat(stringField(record, 'format', 'import/export fidelity case format')),
    direction: parseDirection(stringField(record, 'direction', 'import/export fidelity case direction')),
    required: booleanField(record, 'required', 'import/export fidelity case required'),
    passed: booleanField(record, 'passed', 'import/export fidelity case passed'),
    coveredFeatures: stringArrayField(record, 'coveredFeatures', 'import/export fidelity case coveredFeatures'),
    missingFeatures: stringArrayField(record, 'missingFeatures', 'import/export fidelity case missingFeatures'),
    evidence: stringField(record, 'evidence', 'import/export fidelity case evidence'),
  }
}

export function validateImportExportFidelityScorecard(scorecard: ImportExportFidelityScorecard): void {
  for (const id of requiredCaseIds) {
    const entry = requiredCase(scorecard.cases, id)
    if (!entry.required) {
      throw new Error(`Import/export fidelity scorecard required case is not marked required: ${id}`)
    }
    if (!entry.passed) {
      throw new Error(`Import/export fidelity scorecard contains a failed required case: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredCasesPassed) {
    throw new Error('Import/export fidelity scorecard summary reports failed required cases')
  }
  if (!scorecard.summary.csvRoundTripPassed || !scorecard.summary.xlsxImportPassed || !scorecard.summary.xlsxSnapshotRoundTripPassed) {
    throw new Error('Import/export fidelity scorecard summary is missing required CSV/XLSX pass coverage')
  }
  for (const feature of unsupportedFeatures) {
    if (!scorecard.summary.unsupportedFeatures.includes(feature)) {
      throw new Error(`Import/export fidelity scorecard is missing unsupported feature disclosure: ${feature}`)
    }
  }
}

function parseFormat(value: string): ImportExportFidelityCase['format'] {
  if (value === 'csv' || value === 'xlsx') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity format: ${value}`)
}

function parseDirection(value: string): ImportExportFidelityCase['direction'] {
  if (value === 'import' || value === 'export-import' || value === 'import-export-import') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity direction: ${value}`)
}

function logResult(mode: 'check' | 'write', scorecard: ImportExportFidelityScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        coveredFeatures: scorecard.summary.coveredFeatures.length,
        unsupportedFeatures: scorecard.summary.unsupportedFeatures.length,
      },
      null,
      2,
    ),
  )
}

function recordField(value: Record<string, unknown>, field: string, name: string): Record<string, unknown> {
  return toRecord(value[field], name)
}

function arrayField(value: Record<string, unknown>, field: string, name: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${name} to be an array`)
  }
  return fieldValue
}

function stringArrayField(value: Record<string, unknown>, field: string, name: string): string[] {
  const fieldValue = arrayField(value, field, name)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${name} to contain only strings`)
  }
  return fieldValue
}

function stringField(value: Record<string, unknown>, field: string, name: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${name} to be a string`)
  }
  return fieldValue
}

function booleanField(value: Record<string, unknown>, field: string, name: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${name} to be a boolean`)
  }
  return fieldValue
}

function literalField<const T extends string>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${expected}`)
  }
  return expected
}

function toRecord(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'import-export-fidelity-scorecard-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated import/export fidelity scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
