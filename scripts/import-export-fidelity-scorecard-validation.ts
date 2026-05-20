import type { ImportExportFidelityCase, ImportExportFidelityScorecard } from './gen-import-export-fidelity-scorecard.ts'
import {
  buildImportExportSemanticLedger,
  importExportDeclinedRuntimeFeatures,
  importExportUnsupportedFeatures,
  type ImportExportSemanticDisposition,
  type ImportExportSemanticLedgerEntry,
} from './import-export-semantic-loss-ledger.ts'

const requiredCaseIds = [
  'csv-import-preview',
  'csv-engine-roundtrip',
  'xlsx-import-preview',
  'xlsx-snapshot-roundtrip-values-formulas-formats',
  'xlsx-snapshot-roundtrip-dimensions-merges',
  'xlsx-snapshot-roundtrip-freeze-panes',
  'xlsx-snapshot-roundtrip-filters',
  'xlsx-snapshot-roundtrip-sorts',
  'xlsx-snapshot-roundtrip-sheet-protection',
  'xlsx-snapshot-roundtrip-protected-ranges',
  'xlsx-snapshot-roundtrip-data-validations',
  'xlsx-snapshot-roundtrip-tables',
  'xlsx-snapshot-roundtrip-charts',
  'xlsx-snapshot-roundtrip-pivots',
  'xlsx-formula-context-audit',
  'xlsx-pivot-cache-semantics',
  'xlsx-external-data-provenance',
  'xlsx-macro-payload-preserved-without-execution',
  'xlsx-runtime-feature-policy-warning',
  'external-sheets-excel-import-export-comparison',
] as const

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
      externalImportExportComparisonArtifact: literalField(
        source,
        'externalImportExportComparisonArtifact',
        'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json',
      ),
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
      declinedRuntimeFeatures: stringArrayField(summary, 'declinedRuntimeFeatures', 'import/export fidelity declinedRuntimeFeatures'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'official-docs-comparison-artifact'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'official-docs-comparison-artifact'),
    },
    semanticLedger: arrayField(record, 'semanticLedger', 'import/export fidelity semanticLedger').map(parseSemanticLedgerEntry),
    cases: arrayField(record, 'cases', 'import/export fidelity cases').map(parseImportExportFidelityCase),
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
    if (entry.missingFeatures.length > 0) {
      throw new Error(`Import/export fidelity scorecard required case reports missing features: ${id}`)
    }
  }
  if (!scorecard.summary.allRequiredCasesPassed) {
    throw new Error('Import/export fidelity scorecard summary reports failed required cases')
  }
  if (!scorecard.summary.csvRoundTripPassed || !scorecard.summary.xlsxImportPassed || !scorecard.summary.xlsxSnapshotRoundTripPassed) {
    throw new Error('Import/export fidelity scorecard summary is missing required CSV/XLSX pass coverage')
  }
  validateSummaryCoveredFeatures(scorecard)
  validateSemanticLedger(scorecard)
  const unsupportedFeatures = importExportUnsupportedFeatures(scorecard.semanticLedger)
  for (const feature of unsupportedFeatures) {
    if (!scorecard.summary.unsupportedFeatures.includes(feature)) {
      throw new Error(`Import/export fidelity scorecard is missing unsupported feature disclosure: ${feature}`)
    }
  }
  if (scorecard.summary.unsupportedFeatures.length !== unsupportedFeatures.length) {
    throw new Error('Import/export fidelity scorecard reports unexpected unsupported import/export features')
  }
  const declinedRuntimeFeatures = importExportDeclinedRuntimeFeatures(scorecard.semanticLedger)
  for (const feature of declinedRuntimeFeatures) {
    if (!scorecard.summary.declinedRuntimeFeatures.includes(feature)) {
      throw new Error(`Import/export fidelity scorecard is missing declined runtime feature disclosure: ${feature}`)
    }
  }
}

function validateSemanticLedger(scorecard: ImportExportFidelityScorecard): void {
  const expectedLedger = buildImportExportSemanticLedger(scorecard.summary.coveredFeatures)
  if (JSON.stringify(scorecard.semanticLedger) !== JSON.stringify(expectedLedger)) {
    throw new Error('Import/export fidelity semantic ledger is stale against the current feature evidence')
  }
  const dispositions = new Set(scorecard.semanticLedger.map((entry) => entry.disposition))
  for (const requiredDisposition of ['preserved', 'external', 'declined-runtime'] satisfies ImportExportSemanticDisposition[]) {
    if (!dispositions.has(requiredDisposition)) {
      throw new Error(`Import/export fidelity semantic ledger is missing ${requiredDisposition} entries`)
    }
  }
}

function validateSummaryCoveredFeatures(scorecard: ImportExportFidelityScorecard): void {
  const caseFeatures = new Set(scorecard.cases.flatMap((entry) => entry.coveredFeatures))
  const summaryFeatures = new Set(scorecard.summary.coveredFeatures)
  for (const feature of caseFeatures) {
    if (!summaryFeatures.has(feature)) {
      throw new Error(`Import/export fidelity scorecard summary is missing covered feature: ${feature}`)
    }
  }
  for (const feature of summaryFeatures) {
    if (!caseFeatures.has(feature)) {
      throw new Error(`Import/export fidelity scorecard summary reports uncovered feature: ${feature}`)
    }
  }
}

function requiredCase(cases: readonly ImportExportFidelityCase[], id: string): ImportExportFidelityCase {
  const entry = cases.find((candidate) => candidate.id === id)
  if (!entry) {
    throw new Error(`Import/export fidelity scorecard is missing required case: ${id}`)
  }
  return entry
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

function parseSemanticLedgerEntry(value: unknown): ImportExportSemanticLedgerEntry {
  const record = toRecord(value, 'import/export fidelity semantic ledger entry')
  return {
    feature: stringField(record, 'feature', 'import/export fidelity semantic ledger feature'),
    disposition: parseSemanticDisposition(stringField(record, 'disposition', 'import/export fidelity semantic ledger disposition')),
    reason: stringField(record, 'reason', 'import/export fidelity semantic ledger reason'),
  }
}

function parseSemanticDisposition(value: string): ImportExportSemanticDisposition {
  if (value === 'preserved' || value === 'unsupported' || value === 'external' || value === 'declined-runtime') {
    return value
  }
  throw new Error(`Unexpected import/export semantic ledger disposition: ${value}`)
}

function parseFormat(value: string): ImportExportFidelityCase['format'] {
  if (value === 'csv' || value === 'xlsx' || value === 'external-docs') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity format: ${value}`)
}

function parseDirection(value: string): ImportExportFidelityCase['direction'] {
  if (value === 'import' || value === 'export-import' || value === 'import-export-import' || value === 'comparison') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity direction: ${value}`)
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
