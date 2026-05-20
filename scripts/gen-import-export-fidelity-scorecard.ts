#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { SpreadsheetEngine } from '../packages/core/src/engine.js'
import { exportXlsx, importCsv, importXlsx, manualCalculationModeWarning } from '../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import {
  externalImportExportComparisonArtifactRepoPath,
  externalImportExportComparisonCoveredFeatures,
  parseExternalImportExportComparisonArtifact,
  validateExternalImportExportComparisonArtifact,
} from './import-export-external-sheets-excel-comparison.ts'
import { projectSupportedSnapshotSemantics } from './import-export-fidelity-projection.ts'
import { parseImportExportFidelityScorecard, validateImportExportFidelityScorecard } from './import-export-fidelity-scorecard-validation.ts'
import {
  buildImportExportSemanticLedger,
  importExportDeclinedRuntimeFeatures,
  importExportUnsupportedFeatures,
  type ImportExportSemanticLedgerEntry,
} from './import-export-semantic-loss-ledger.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export { parseImportExportFidelityScorecard, validateImportExportFidelityScorecard }

export interface ImportExportFidelityCase {
  readonly id: string
  readonly format: 'csv' | 'xlsx' | 'external-docs'
  readonly direction: 'import' | 'export-import' | 'import-export-import' | 'comparison'
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
    readonly externalImportExportComparisonArtifact: 'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json'
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly csvRoundTripPassed: boolean
    readonly xlsxImportPassed: boolean
    readonly xlsxSnapshotRoundTripPassed: boolean
    readonly coveredFeatures: string[]
    readonly unsupportedFeatures: string[]
    readonly declinedRuntimeFeatures: string[]
    readonly externalGoogleSheetsEvidence: 'official-docs-comparison-artifact'
    readonly externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact'
  }
  readonly semanticLedger: ImportExportSemanticLedgerEntry[]
  readonly cases: ImportExportFidelityCase[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'import-export-fidelity-scorecard.json')
const externalImportExportComparisonArtifactPath = join(rootDir, externalImportExportComparisonArtifactRepoPath)
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
  'xlsx.formulaAudit.context',
  'xlsx.formulaAudit.cacheStatus',
  'xlsx.numberFormats',
  'xlsx.workbookProperties',
  'xlsx.calculationSettings',
  'xlsx.calculationSettings.calcChainDiagnostics',
  'xlsx.definedNames',
  'xlsx.comments',
  'xlsx.styles',
  'xlsx.conditionalFormats.roundtrip',
  'xlsx.rowColumnDimensions',
  'xlsx.merges',
  'xlsx.freezePanes.roundtrip',
  'xlsx.filters.roundtrip',
  'xlsx.sorts.roundtrip',
  'xlsx.sheetProtection.roundtrip',
  'xlsx.protectedRanges.roundtrip',
  'xlsx.dataValidations.roundtrip',
  'xlsx.tables.roundtrip',
  'xlsx.charts.roundtrip',
  'xlsx.pivots.roundtrip',
  'xlsx.pivots.cacheSemantics',
  'xlsx.pivots.externalCacheOnlySemantics',
  'xlsx.multiSheet',
  'xlsx.macros.detectedNoExecution',
  'xlsx.macros.payloadRoundtrip',
  'xlsx.macros.codeNameRoundtrip',
  'xlsx.externalData.provenance',
  'xlsx.runtimeFeaturePolicyWarnings',
  ...externalImportExportComparisonCoveredFeatures,
] as const
const unsupportedFeatureDisclosures = importExportUnsupportedFeatures()
const declinedRuntimeFeatureDisclosures = importExportDeclinedRuntimeFeatures()

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
    runXlsxSnapshotRoundTripSortsCase(),
    runXlsxSnapshotRoundTripSheetProtectionCase(),
    runXlsxSnapshotRoundTripProtectedRangesCase(),
    runXlsxSnapshotRoundTripDataValidationsCase(),
    runXlsxSnapshotRoundTripTablesCase(),
    runXlsxSnapshotRoundTripChartsCase(),
    runXlsxSnapshotRoundTripPivotsCase(),
    runXlsxFormulaContextAuditCase(),
    runXlsxPivotCacheSemanticsCase(),
    runXlsxExternalDataProvenanceCase(),
    runXlsxMacroPayloadPreservedWithoutExecutionCase(),
    runXlsxRuntimeFeaturePolicyWarningCase(),
    runExternalSheetsExcelImportExportComparisonCase(),
  ]
  const coveredFeatureSet = new Set(cases.flatMap((entry) => entry.coveredFeatures))
  const coveredFeatures = coveredFeatureOrder.filter((feature) => coveredFeatureSet.has(feature))
  const semanticLedger = buildImportExportSemanticLedger(coveredFeatures)

  return {
    schemaVersion: 1,
    suite: 'import-export-fidelity',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-import-export-fidelity-scorecard.ts',
      implementationPackage: 'packages/excel-import',
      enginePackage: 'packages/core',
      externalImportExportComparisonArtifact: externalImportExportComparisonArtifactRepoPath,
    },
    summary: {
      allRequiredCasesPassed: cases.filter((entry) => entry.required).every((entry) => entry.passed),
      csvRoundTripPassed: scorecardCase(cases, 'csv-engine-roundtrip').passed,
      xlsxImportPassed: scorecardCase(cases, 'xlsx-import-preview').passed,
      xlsxSnapshotRoundTripPassed:
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-values-formulas-formats').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-dimensions-merges').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-freeze-panes').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-filters').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-sorts').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-sheet-protection').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-protected-ranges').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-data-validations').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-tables').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-charts').passed &&
        scorecardCase(cases, 'xlsx-snapshot-roundtrip-pivots').passed,
      coveredFeatures,
      unsupportedFeatures: [...importExportUnsupportedFeatures(semanticLedger)],
      declinedRuntimeFeatures: [...importExportDeclinedRuntimeFeatures(semanticLedger)],
      externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
      externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
    },
    semanticLedger: [...semanticLedger],
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
    JSON.stringify(actual.properties) === JSON.stringify(expected.properties) &&
    JSON.stringify(actual.calculationSettings) === JSON.stringify(expected.calculationSettings) &&
    JSON.stringify(actual.valueFormulaFormatSheets) === JSON.stringify(expected.valueFormulaFormatSheets) &&
    JSON.stringify(actual.commentThreads) === JSON.stringify(expected.commentThreads) &&
    JSON.stringify(actual.styleRanges) === JSON.stringify(expected.styleRanges) &&
    JSON.stringify(actual.conditionalFormats) === JSON.stringify(expected.conditionalFormats)
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
      'xlsx.workbookProperties',
      'xlsx.calculationSettings',
      'xlsx.definedNames',
      'xlsx.comments',
      'xlsx.styles',
      'xlsx.conditionalFormats.roundtrip',
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
    evidence:
      'WorkbookSnapshot exported to XLSX imports back with equivalent sheet filter ranges and criteria backed by native XLSX autoFilter nodes.',
  })
}

function runXlsxSnapshotRoundTripSortsCase(): ImportExportFidelityCase {
  const expected = projectSupportedSnapshotSemantics(createFidelitySnapshot())
  const actual = projectSupportedSnapshotSemantics(importXlsx(exportXlsx(createFidelitySnapshot()), 'fidelity.xlsx').snapshot)
  const passed = JSON.stringify(actual.sorts) === JSON.stringify(expected.sorts)
  return fidelityCase({
    id: 'xlsx-snapshot-roundtrip-sorts',
    format: 'xlsx',
    direction: 'export-import',
    passed,
    coveredFeatures: ['xlsx.sorts.roundtrip'],
    evidence: 'WorkbookSnapshot exported to XLSX imports back with equivalent sort metadata backed by native XLSX sortState nodes.',
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

function runXlsxFormulaContextAuditCase(): ImportExportFidelityCase {
  const imported = importXlsx(createFormulaContextAuditWorkbookBytes(), 'formula-context-audit.xlsx')
  const audit = imported.snapshot.workbook.metadata?.formulaAudit
  const worksheetFormula = audit?.formulas.find((entry) => entry.context === 'worksheet-cell' && entry.address === 'B1')
  const definedNameFormula = audit?.formulas.find((entry) => entry.context === 'defined-name' && entry.name === 'R1C1Name')
  const calcChain = audit?.calcChain
  const passed =
    worksheetFormula?.formula === 'A1*2' &&
    worksheetFormula.cacheStatus === 'staleRisk' &&
    definedNameFormula?.formula === 'R1C1' &&
    audit?.diagnostics.some((diagnostic) => diagnostic.code === 'r1c1-reference') === true &&
    calcChain?.cells.some((cell) => cell.sheetName === 'Sheet1' && cell.address === 'B1') === true &&
    imported.snapshot.workbook.metadata?.calculationSettings?.calcId === 191029 &&
    imported.snapshot.workbook.metadata.calculationSettings.forceFullCalc === true

  return fidelityCase({
    id: 'xlsx-formula-context-audit',
    format: 'xlsx',
    direction: 'import',
    passed,
    coveredFeatures: ['xlsx.formulaAudit.context', 'xlsx.formulaAudit.cacheStatus', 'xlsx.calculationSettings.calcChainDiagnostics'],
    evidence:
      'XLSX import records worksheet formula context, cached-result trust state, defined-name formula diagnostics, calcId/forceFullCalc metadata, and calc-chain cells.',
  })
}

function runXlsxPivotCacheSemanticsCase(): ImportExportFidelityCase {
  const imported = importXlsx(exportXlsx(createFidelitySnapshot()), 'pivot-cache-semantics.xlsx')
  const pivot = imported.snapshot.workbook.metadata?.pivots?.find((entry) => entry.name === 'SalesByRegion')
  const externalCachePivot = importXlsx(createExternalCacheOnlyPivotWorkbookBytes(), 'external-cache-only-pivot.xlsx').snapshot.workbook
    .metadata?.pivots?.[0]
  const passed =
    pivot?.sourceKind === 'worksheet' &&
    JSON.stringify(pivot.cacheFields) === JSON.stringify(['Region', 'Product', 'Sales', 'Notes']) &&
    (pivot.cachedRecords?.length ?? 0) === 3 &&
    pivot.values.some((value) => value.sourceColumn === 'Sales' && value.summarizeBy === 'sum') &&
    externalCachePivot?.sourceKind === 'external-cache-only' &&
    externalCachePivot.cacheOnly === true &&
    JSON.stringify(externalCachePivot.cachedRecords) ===
      JSON.stringify([
        ['East', 10],
        ['West', 7],
      ])

  return fidelityCase({
    id: 'xlsx-pivot-cache-semantics',
    format: 'xlsx',
    direction: 'import',
    passed,
    coveredFeatures: ['xlsx.pivots.cacheSemantics', 'xlsx.pivots.externalCacheOnlySemantics'],
    evidence:
      'XLSX import projects pivot cache field names, source kind, cached records, row fields, aggregate metadata, and cache-only external pivots separately from raw pivot part preservation.',
  })
}

function runXlsxExternalDataProvenanceCase(): ImportExportFidelityCase {
  const imported = importXlsx(createExternalDataWorkbookBytes(), 'external-data-provenance.xlsx')
  const externalConnections = imported.snapshot.workbook.metadata?.externalConnections
  const passed =
    externalConnections?.refreshExecution === 'disabled' &&
    externalConnections.connections.some((connection) => connection.name === 'Sales Query' && connection.sourceKind === 'database') &&
    externalConnections.externalLinks.some((link) => link.kind === 'external-workbook' && link.target === 'file:///tmp/source.xlsx') &&
    externalConnections.externalLinks.some((link) => link.kind === 'dde' && link.refreshExecution === 'disabled') &&
    externalConnections.externalLinks.some((link) => link.kind === 'ole' && link.refreshExecution === 'disabled')

  return fidelityCase({
    id: 'xlsx-external-data-provenance',
    format: 'xlsx',
    direction: 'import',
    passed,
    coveredFeatures: ['xlsx.externalData.provenance'],
    evidence:
      'XLSX import parses connection, external workbook, DDE, and OLE provenance while leaving all external refresh execution disabled.',
  })
}

function createFormulaContextAuditWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[2, null]])
  sheet.B1 = { t: 'n', f: 'A1*2', v: 4 }
  sheet['!ref'] = 'A1:B1'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const workbookXmlWithCalcPr = /<calcPr\b/u.test(workbookXml)
    ? workbookXml.replace(/<calcPr\b[^>]*\/>/u, '<calcPr calcId="191029" forceFullCalc="1"/>')
    : workbookXml.replace('</workbook>', '<calcPr calcId="191029" forceFullCalc="1"/></workbook>')
  zip['xl/workbook.xml'] = strToU8(
    workbookXmlWithCalcPr.replace('</sheets>', '</sheets><definedNames><definedName name="R1C1Name">R1C1</definedName></definedNames>'),
  )
  zip['xl/calcChain.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>',
    ].join(''),
  )
  return zipSync(zip)
}

function createExternalCacheOnlyPivotWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[]]), 'Pivot')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array()).replace(
      '</sheets>',
      '</sheets><pivotCaches><pivotCache cacheId="1" r:id="rIdExternalPivotCache"/></pivotCaches>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
      '</Relationships>',
      '<Relationship Id="rIdExternalPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>',
    ),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      .replace(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      )
      .replace('</worksheet>', '<pivotTableDefinition r:id="rIdExternalPivot"/></worksheet>'),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternalPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>',
  )
  zip['xl/pivotTables/pivotTable1.xml'] = strToU8(
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="ExternalSales" cacheId="1"><location ref="A1:B3" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField name="Sales Total" fld="1" subtotal="sum"/></dataFields></pivotTableDefinition>',
  )
  zip['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8(
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdRecords" recordCount="2"><cacheSource type="external"/><cacheFields count="2"><cacheField name="Region"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Sales"><sharedItems count="2"><n v="10"/><n v="7"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>',
  )
  zip['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>',
  )
  zip['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8(
    '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2"><r><x v="0"/><x v="0"/></r><r><x v="1"/><x v="1"/></r></pivotCacheRecords>',
  )
  return zipSync(zip)
}

function createExternalDataWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Local'], [1]]), 'Model')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array()).replace(
      '</workbook>',
      '<externalReferences><externalReference r:id="rIdExternal1"/></externalReferences></workbook>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
      '</Relationships>',
      '<Relationship Id="rIdExternal1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
    ),
  )
  zip['xl/connections.xml'] = strToU8(
    '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><connection id="1" name="Sales Query" type="5" refreshedVersion="8" refreshOnLoad="0"><dbPr connection="Provider=SQLOLEDB;Data Source=example" command="SELECT * FROM Sales" commandType="2"/></connection></connections>',
  )
  zip['xl/externalLinks/externalLink1.xml'] = strToU8(
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Source"/></sheetNames></externalBook><ddeLink ddeService="cmd" ddeTopic="topic"><ddeItems count="1"><ddeItem name="A1"/></ddeItems></ddeLink><oleLink progId="Word.Document" r:id="rId2"><oleItems count="1"><oleItem name="Document"/></oleItems></oleLink></externalLink>',
  )
  zip['xl/externalLinks/_rels/externalLink1.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/source.xlsx" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///tmp/document.docx" TargetMode="External"/></Relationships>',
  )
  return zipSync(zip)
}

function runXlsxMacroPayloadPreservedWithoutExecutionCase(): ImportExportFidelityCase {
  const imported = importXlsx(createMacroEnabledWorkbookBytes(), 'macro-enabled.xlsm')
  const exported = importXlsx(exportXlsx(imported.snapshot), 'macro-enabled-roundtrip.xlsm')
  const safeCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')
  const roundTripSafeCell = exported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')
  const macroPayload = imported.snapshot.workbook.metadata?.macroPayloads?.[0]
  const roundTripMacroPayload = exported.snapshot.workbook.metadata?.macroPayloads?.[0]
  const passed =
    imported.workbookName === 'macro-enabled' &&
    imported.warnings.includes('Macros were preserved but not executed during XLSX import.') &&
    exported.warnings.includes('Macros were preserved but not executed during XLSX import.') &&
    safeCell?.value === 'safe macro workbook value' &&
    roundTripSafeCell?.value === 'safe macro workbook value' &&
    macroPayload?.dataBase64 === 'AQIDBA==' &&
    macroPayload.workbookCodeName === 'ThisWorkbook' &&
    macroPayload.sheetCodeNames?.[0]?.codeName === 'Sheet1' &&
    roundTripMacroPayload?.dataBase64 === macroPayload.dataBase64 &&
    roundTripMacroPayload.workbookCodeName === macroPayload.workbookCodeName &&
    JSON.stringify(roundTripMacroPayload.sheetCodeNames) === JSON.stringify(macroPayload.sheetCodeNames)

  return fidelityCase({
    id: 'xlsx-macro-payload-preserved-without-execution',
    format: 'xlsx',
    direction: 'import-export-import',
    passed,
    coveredFeatures: [
      'xlsx.macros.detectedNoExecution',
      'xlsx.macros.payloadRoundtrip',
      'xlsx.macros.codeNameRoundtrip',
      'xlsx.runtimeFeaturePolicyWarnings',
    ],
    evidence:
      'Macro-enabled XLSM import preserves safe workbook cells, records a non-execution warning, exports an XLSM with the original VBA payload and code names, re-imports with identical macro payload metadata, and treats native macro execution as an unsafe runtime non-goal.',
  })
}

function runXlsxRuntimeFeaturePolicyWarningCase(): ImportExportFidelityCase {
  const snapshot = createFidelitySnapshot()
  const imported = importXlsx(exportXlsx(snapshot), 'fidelity.xlsx')
  const expectedWarnings = [manualCalculationModeWarning]
  return fidelityCase({
    id: 'xlsx-runtime-feature-policy-warning',
    format: 'xlsx',
    direction: 'export-import',
    passed:
      JSON.stringify(imported.warnings) === JSON.stringify(expectedWarnings) &&
      declinedRuntimeFeatureDisclosures.length > 0 &&
      unsupportedFeatureDisclosures.length === 0,
    coveredFeatures: ['xlsx.runtimeFeaturePolicyWarnings'],
    evidence:
      'Scorecard separates import/export compatibility from unsafe runtime execution surfaces: preserved manual calculation metadata emits the expected stale-cache warning, while native macro execution is disclosed as a declined runtime feature.',
  })
}

function createMacroEnabledWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['safe macro workbook value']]), 'Sheet1')
  workbook.Workbook = {
    WBProps: { CodeName: 'ThisWorkbook' },
    Sheets: [{ name: 'Sheet1', CodeName: 'Sheet1' }],
  }
  workbook.vbaraw = new Uint8Array([1, 2, 3, 4])
  return XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true })
}

function runExternalSheetsExcelImportExportComparisonCase(): ImportExportFidelityCase {
  const artifact = parseExternalImportExportComparisonArtifact(
    JSON.parse(readFileSync(externalImportExportComparisonArtifactPath, 'utf8')) as unknown,
  )
  const findings = validateExternalImportExportComparisonArtifact(artifact)
  const googleSourceCount = artifact.officialSources.filter((source) => source.vendor === 'google-sheets').length
  const microsoftSourceCount = artifact.officialSources.filter((source) => source.vendor === 'microsoft-excel').length

  return fidelityCase({
    id: 'external-sheets-excel-import-export-comparison',
    format: 'external-docs',
    direction: 'comparison',
    passed: findings.length === 0,
    coveredFeatures: externalImportExportComparisonCoveredFeatures,
    missingFeatures: findings,
    evidence:
      `Validated ${externalImportExportComparisonArtifactRepoPath} from ${artifact.sourceBasis}: ` +
      `${String(artifact.dimensions.length)} required comparison dimensions cite ${String(googleSourceCount)} official Google Sheets/Drive sources ` +
      `and ${String(microsoftSourceCount)} official Microsoft Excel sources.`,
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
        calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
        properties: [
          { key: 'locale', value: 'en-US' },
          { key: 'reviewed', value: true },
          { key: 'threshold', value: 0.085 },
        ],
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
          filters: [
            {
              sheetName: 'Summary',
              startAddress: 'A1',
              endAddress: 'B3',
              criteria: [
                { colId: 0, filters: { blank: false, values: ['Revenue'] } },
                { colId: 1, customFilters: { filters: [{ operator: 'greaterThan', value: '1000' }] } },
              ],
            },
          ],
          sorts: [
            {
              range: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
              keys: [{ keyAddress: 'B1', direction: 'desc' }],
            },
          ],
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
          conditionalFormats: [
            {
              id: 'summary-high-total',
              range: { sheetName: 'Summary', startAddress: 'B2', endAddress: 'B3' },
              rule: { kind: 'cellIs', operator: 'greaterThan', values: [1000] },
              style: { fill: { backgroundColor: '#f4cccc' }, font: { bold: true, color: '#990000' } },
              stopIfTrue: true,
              priority: 1,
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

function scorecardCase(cases: readonly ImportExportFidelityCase[], id: string): ImportExportFidelityCase {
  const entry = cases.find((candidate) => candidate.id === id)
  if (!entry) {
    throw new Error(`Import/export fidelity scorecard is missing required case: ${id}`)
  }
  return entry
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
        declinedRuntimeFeatures: scorecard.summary.declinedRuntimeFeatures.length,
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
