import type {
  PublicWorkbookArtifact,
  PublicWorkbookCaseStatus,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookFeatureCounts,
  PublicWorkbookExternalReferenceSummary,
  PublicWorkbookLicenseEvidence,
  PublicWorkbookManifest,
  PublicWorkbookSource,
  PublicWorkbookSourceKind,
  PublicWorkbookValidationSummary,
  PublicWorkbookVerificationPhase,
  PublicWorkbookVerificationPhaseTiming,
} from './public-workbook-corpus-types.ts'

const allowedLicenseTokens = [
  'cc0',
  'cc-by',
  'cc-by-sa',
  'creative commons attribution',
  'public domain',
  'odc-by',
  'odbl',
  'ogl',
  'open government licence',
  'open data commons',
  'mit',
  'apache-2.0',
  'bsd-2-clause',
  'bsd-3-clause',
  'mpl-2.0',
]

export const defaultPublicWorkbookTargetCount = 10_000

export function createEmptyPublicWorkbookManifest(
  generatedAt = new Date().toISOString(),
  targetWorkbookCount = defaultPublicWorkbookTargetCount,
): PublicWorkbookManifest {
  validateTargetWorkbookCount(targetWorkbookCount)
  return {
    schemaVersion: 1,
    corpus: 'public-workbook-corpus',
    targetWorkbookCount,
    generatedAt,
    sources: [],
    artifacts: [],
  }
}

export function validatePublicWorkbookManifest(manifest: PublicWorkbookManifest): void {
  if (manifest.schemaVersion !== 1 || manifest.corpus !== 'public-workbook-corpus') {
    throw new Error('Unexpected public workbook corpus manifest header')
  }
  validateTargetWorkbookCount(manifest.targetWorkbookCount)
  const sourceIds = new Set<string>()
  for (const source of manifest.sources) {
    if (sourceIds.has(source.id)) {
      throw new Error(`Duplicate public workbook source id: ${source.id}`)
    }
    sourceIds.add(source.id)
    if (!hasUsableLicenseEvidence(source.license)) {
      throw new Error(`Public workbook source ${source.id} is missing usable license evidence`)
    }
    if (!isSpreadsheetUrl(source.downloadUrl) && !isSpreadsheetFileName(source.fileName)) {
      throw new Error(`Public workbook source ${source.id} is not a spreadsheet candidate`)
    }
  }
  const exhaustedSourceIds = new Set<string>()
  for (const sourceId of manifest.fetchState?.exhaustedSourceIds ?? []) {
    if (exhaustedSourceIds.has(sourceId)) {
      throw new Error(`Duplicate exhausted public workbook source id: ${sourceId}`)
    }
    exhaustedSourceIds.add(sourceId)
    if (!sourceIds.has(sourceId)) {
      throw new Error(`Exhausted public workbook source ${sourceId} is not in the manifest`)
    }
  }
  const artifactIds = new Set<string>()
  const hashes = new Set<string>()
  const fingerprints = new Set<string>()
  for (const artifact of manifest.artifacts) {
    if (artifactIds.has(artifact.id)) {
      throw new Error(`Duplicate public workbook artifact id: ${artifact.id}`)
    }
    artifactIds.add(artifact.id)
    if (!sourceIds.has(artifact.sourceId)) {
      throw new Error(`Public workbook artifact ${artifact.id} references unknown source ${artifact.sourceId}`)
    }
    if (!/^[0-9a-f]{64}$/u.test(artifact.sha256)) {
      throw new Error(`Public workbook artifact ${artifact.id} has an invalid sha256`)
    }
    if (hashes.has(artifact.sha256)) {
      throw new Error(`Duplicate public workbook artifact sha256: ${artifact.sha256}`)
    }
    hashes.add(artifact.sha256)
    if (fingerprints.has(artifact.workbookFingerprint)) {
      throw new Error(`Duplicate public workbook structure fingerprint: ${artifact.workbookFingerprint}`)
    }
    fingerprints.add(artifact.workbookFingerprint)
    if (!hasUsableLicenseEvidence(artifact.license)) {
      throw new Error(`Public workbook artifact ${artifact.id} is missing usable license evidence`)
    }
  }
}

export function parsePublicWorkbookManifestJson(value: unknown): PublicWorkbookManifest {
  const record = asRecord(value)
  const fetchState = parsePublicWorkbookFetchState(record['fetchState'])
  const manifest: PublicWorkbookManifest = {
    schemaVersion: readExpectedNumber(record, 'schemaVersion', 1),
    corpus: readExpectedString(record, 'corpus', 'public-workbook-corpus'),
    targetWorkbookCount: readTargetWorkbookCount(record, 'targetWorkbookCount'),
    generatedAt: readRequiredString(record, 'generatedAt'),
    sources: readRequiredArray(record, 'sources').map(parsePublicWorkbookSource),
    artifacts: readRequiredArray(record, 'artifacts').map(parsePublicWorkbookArtifact),
    ...(fetchState ? { fetchState } : {}),
  }
  validatePublicWorkbookManifest(manifest)
  return manifest
}

export function parsePublicWorkbookCorpusScorecardJson(value: unknown): PublicWorkbookCorpusScorecard {
  const record = asRecord(value)
  const summary = asRecord(record['summary'])
  const scorecard: PublicWorkbookCorpusScorecard = {
    schemaVersion: readExpectedNumber(record, 'schemaVersion', 1),
    suite: readExpectedString(record, 'suite', 'public-workbook-corpus'),
    generatedAt: readRequiredString(record, 'generatedAt'),
    summary: {
      targetWorkbookCount: readTargetWorkbookCount(summary, 'targetWorkbookCount'),
      sourceCount: readRequiredInteger(summary, 'sourceCount'),
      cachedWorkbookCount: readRequiredInteger(summary, 'cachedWorkbookCount'),
      importedWorkbookCount: readRequiredInteger(summary, 'importedWorkbookCount'),
      passedWorkbookCount: readRequiredInteger(summary, 'passedWorkbookCount'),
      failedWorkbookCount: readRequiredInteger(summary, 'failedWorkbookCount'),
      errorWorkbookCount: readRequiredInteger(summary, 'errorWorkbookCount'),
      unsupportedWorkbookCount: readRequiredInteger(summary, 'unsupportedWorkbookCount'),
      formulaOracleComparisonCount: readRequiredInteger(summary, 'formulaOracleComparisonCount'),
      formulaOracleMatchCount: readRequiredInteger(summary, 'formulaOracleMatchCount'),
      structuralSmokeRunCount: readRequiredInteger(summary, 'structuralSmokeRunCount'),
      allCachedWorkbooksPassed: readRequiredBoolean(summary, 'allCachedWorkbooksPassed'),
      remainingToTarget: readRequiredInteger(summary, 'remainingToTarget'),
    },
    cases: readRequiredArray(record, 'cases').map(parsePublicWorkbookCorpusCase),
  }
  validatePublicWorkbookCorpusScorecard(scorecard)
  return scorecard
}

export function parsePublicWorkbookCorpusCase(value: unknown): PublicWorkbookCorpusCase {
  const record = asRecord(value)
  const elapsedMs = readOptionalNonNegativeInteger(record, 'elapsedMs')
  const peakRssBytes = readOptionalNonNegativeIntegerOrNull(record, 'peakRssBytes')
  const phaseTimings = parseOptionalPublicWorkbookVerificationPhaseTimings(record['phaseTimings'])
  const externalWorkbookReferences = parseOptionalPublicWorkbookExternalReferenceSummary(record['externalWorkbookReferences'])
  return {
    id: readRequiredString(record, 'id'),
    sourceId: readRequiredString(record, 'sourceId'),
    sourceUrl: readRequiredString(record, 'sourceUrl'),
    fileName: readRequiredString(record, 'fileName'),
    sha256: readRequiredString(record, 'sha256'),
    byteSize: readRequiredInteger(record, 'byteSize'),
    license: parsePublicWorkbookLicenseEvidence(record['license']),
    status: parsePublicWorkbookCaseStatus(readRequiredString(record, 'status')),
    passed: readRequiredBoolean(record, 'passed'),
    ...(elapsedMs !== undefined ? { elapsedMs } : {}),
    ...(peakRssBytes !== undefined ? { peakRssBytes } : {}),
    ...(phaseTimings !== undefined ? { phaseTimings } : {}),
    ...(externalWorkbookReferences !== undefined ? { externalWorkbookReferences } : {}),
    featureCounts: parsePublicWorkbookFeatureCounts(record['featureCounts']),
    workbookMetadata: parsePublicWorkbookMetadata(record['workbookMetadata']),
    validation: parsePublicWorkbookValidationSummary(record['validation']),
    unsupportedFeatureClassifications: readStringArray(record, 'unsupportedFeatureClassifications'),
    evidence: readStringArray(record, 'evidence'),
  }
}

export function hasUsableLicenseEvidence(license: PublicWorkbookLicenseEvidence): boolean {
  const title = license.title.trim().toLowerCase()
  const spdx = license.spdxId?.trim().toLowerCase() ?? ''
  const evidenceUrl = license.evidenceUrl?.trim() ?? ''
  if (title.length === 0 || evidenceUrl.length === 0) {
    return false
  }
  const haystack = `${spdx} ${title}`
  return allowedLicenseTokens.some((token) => haystack.includes(token))
}

export function isSpreadsheetUrl(value: string): boolean {
  return /\.(xlsx|xlsm|xls)(?:[?#].*)?$/iu.test(value.trim())
}

export function isSpreadsheetFileName(value: string): boolean {
  return /\.(xlsx|xlsm|xls)$/iu.test(value.trim())
}

export function spreadsheetExtension(fileName: string): 'xlsx' | 'xlsm' | 'xls' {
  const lower = fileName.toLowerCase()
  if (lower.endsWith('.xlsm')) {
    return 'xlsm'
  }
  if (lower.endsWith('.xls')) {
    return 'xls'
  }
  return 'xlsx'
}

export function asRecord(value: unknown): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error('Expected object')
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

export function asRecordOrNull(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? asRecord(value) : null
}

export function readString(value: Record<string, unknown>, key: string): string | null {
  const fieldValue = value[key]
  return typeof fieldValue === 'string' && fieldValue.trim().length > 0 ? fieldValue.trim() : null
}

export function readArray(value: Record<string, unknown>, key: string): unknown[] {
  const fieldValue = value[key]
  return Array.isArray(fieldValue) ? fieldValue : []
}

function validatePublicWorkbookCorpusScorecard(scorecard: PublicWorkbookCorpusScorecard): void {
  if (scorecard.schemaVersion !== 1 || scorecard.suite !== 'public-workbook-corpus') {
    throw new Error('Unexpected public workbook corpus scorecard header')
  }
  validateTargetWorkbookCount(scorecard.summary.targetWorkbookCount)
  if (scorecard.cases.length !== scorecard.summary.cachedWorkbookCount) {
    throw new Error('Public workbook corpus scorecard case count does not match cached workbook count')
  }
  if (scorecard.summary.remainingToTarget !== Math.max(0, scorecard.summary.targetWorkbookCount - scorecard.summary.cachedWorkbookCount)) {
    throw new Error('Public workbook corpus scorecard remaining target count is stale')
  }
  const passedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'passed').length
  const failedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'failed').length
  const errorWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'error').length
  const unsupportedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'unsupported').length
  if (scorecard.summary.passedWorkbookCount !== passedWorkbookCount) {
    throw new Error('Public workbook corpus scorecard passed workbook count is stale')
  }
  if (scorecard.summary.failedWorkbookCount !== failedWorkbookCount) {
    throw new Error('Public workbook corpus scorecard failed workbook count is stale')
  }
  if (scorecard.summary.errorWorkbookCount !== errorWorkbookCount) {
    throw new Error('Public workbook corpus scorecard error workbook count is stale')
  }
  if (scorecard.summary.unsupportedWorkbookCount !== unsupportedWorkbookCount) {
    throw new Error('Public workbook corpus scorecard unsupported workbook count is stale')
  }
  const importedWorkbookCount = scorecard.cases.filter((entry) => entry.validation.importPassed).length
  if (scorecard.summary.importedWorkbookCount !== importedWorkbookCount) {
    throw new Error('Public workbook corpus scorecard imported workbook count is stale')
  }
  const formulaOracleComparisonCount = scorecard.cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0)
  if (scorecard.summary.formulaOracleComparisonCount !== formulaOracleComparisonCount) {
    throw new Error('Public workbook corpus scorecard formula oracle comparison count is stale')
  }
  if (scorecard.summary.formulaOracleMatchCount !== countFormulaOracleMatches(scorecard.cases)) {
    throw new Error('Public workbook corpus scorecard formula oracle match count is stale')
  }
  if (scorecard.summary.allCachedWorkbooksPassed !== scorecard.cases.every((entry) => entry.passed)) {
    throw new Error('Public workbook corpus scorecard pass summary is stale')
  }
  if (!scorecard.summary.allCachedWorkbooksPassed) {
    throw new Error('Public workbook corpus scorecard has cached workbooks that did not pass')
  }
}

function parsePublicWorkbookFetchState(value: unknown): PublicWorkbookManifest['fetchState'] | undefined {
  if (value === undefined) {
    return undefined
  }
  const record = asRecord(value)
  return {
    exhaustedSourceIds: readStringArray(record, 'exhaustedSourceIds'),
  }
}

function parsePublicWorkbookSource(value: unknown): PublicWorkbookSource {
  const record = asRecord(value)
  const portal = readOptionalString(record, 'portal')
  const datasetId = readOptionalString(record, 'datasetId')
  const resourceId = readOptionalString(record, 'resourceId')
  return {
    id: readRequiredString(record, 'id'),
    kind: parsePublicWorkbookSourceKind(readRequiredString(record, 'kind')),
    sourceUrl: readRequiredString(record, 'sourceUrl'),
    downloadUrl: readRequiredString(record, 'downloadUrl'),
    fileName: readRequiredString(record, 'fileName'),
    discoveredAt: readRequiredString(record, 'discoveredAt'),
    license: parsePublicWorkbookLicenseEvidence(record['license']),
    ...(readOptionalStringArray(record, 'topicEvidence') ? { topicEvidence: readOptionalStringArray(record, 'topicEvidence') } : {}),
    ...(portal ? { portal } : {}),
    ...(datasetId ? { datasetId } : {}),
    ...(resourceId ? { resourceId } : {}),
  }
}

export function parsePublicWorkbookArtifact(value: unknown): PublicWorkbookArtifact {
  const record = asRecord(value)
  return {
    id: readRequiredString(record, 'id'),
    sourceId: readRequiredString(record, 'sourceId'),
    sourceUrl: readRequiredString(record, 'sourceUrl'),
    downloadUrl: readRequiredString(record, 'downloadUrl'),
    fileName: readRequiredString(record, 'fileName'),
    cachePath: readRequiredString(record, 'cachePath'),
    sha256: readRequiredString(record, 'sha256'),
    byteSize: readRequiredInteger(record, 'byteSize'),
    workbookFingerprint: readRequiredString(record, 'workbookFingerprint'),
    fetchedAt: readRequiredString(record, 'fetchedAt'),
    license: parsePublicWorkbookLicenseEvidence(record['license']),
    ...(readOptionalStringArray(record, 'topicEvidence') ? { topicEvidence: readOptionalStringArray(record, 'topicEvidence') } : {}),
  }
}

function parsePublicWorkbookLicenseEvidence(value: unknown): PublicWorkbookLicenseEvidence {
  const record = asRecord(value)
  return {
    spdxId: readNullableString(record, 'spdxId'),
    title: readRequiredString(record, 'title'),
    evidenceUrl: readNullableString(record, 'evidenceUrl'),
  }
}

function parsePublicWorkbookFeatureCounts(value: unknown): PublicWorkbookFeatureCounts {
  const record = asRecord(value)
  return {
    sheetCount: readRequiredInteger(record, 'sheetCount'),
    cellCount: readRequiredInteger(record, 'cellCount'),
    formulaCellCount: readRequiredInteger(record, 'formulaCellCount'),
    valueCellCount: readRequiredInteger(record, 'valueCellCount'),
    definedNameCount: readRequiredInteger(record, 'definedNameCount'),
    tableCount: readRequiredInteger(record, 'tableCount'),
    chartCount: readRequiredInteger(record, 'chartCount'),
    pivotCount: readRequiredInteger(record, 'pivotCount'),
    mergeCount: readRequiredInteger(record, 'mergeCount'),
    styleRangeCount: readRequiredInteger(record, 'styleRangeCount'),
    conditionalFormatCount: readRequiredInteger(record, 'conditionalFormatCount'),
    dataValidationCount: readRequiredInteger(record, 'dataValidationCount'),
    macroPayloadCount: readRequiredInteger(record, 'macroPayloadCount'),
    warningCount: readRequiredInteger(record, 'warningCount'),
  }
}

function parseOptionalPublicWorkbookVerificationPhaseTimings(value: unknown): readonly PublicWorkbookVerificationPhaseTiming[] | undefined {
  if (value === undefined) {
    return undefined
  }
  return readRequiredArray({ phaseTimings: value }, 'phaseTimings').map((entry) => {
    const record = asRecord(entry)
    return {
      phase: parsePublicWorkbookVerificationPhase(readRequiredString(record, 'phase')),
      elapsedMs: readRequiredInteger(record, 'elapsedMs'),
    }
  })
}

function parseOptionalPublicWorkbookExternalReferenceSummary(value: unknown): PublicWorkbookExternalReferenceSummary | undefined {
  if (value === undefined) {
    return undefined
  }
  const record = asRecord(value)
  return {
    linkedWorkbookCount: readRequiredInteger(record, 'linkedWorkbookCount'),
    formulaDependencyCount: readRequiredInteger(record, 'formulaDependencyCount'),
    cachedValueDependencyCount: readRequiredInteger(record, 'cachedValueDependencyCount'),
  }
}

function parsePublicWorkbookMetadata(value: unknown): PublicWorkbookCorpusCase['workbookMetadata'] {
  const record = asRecord(value)
  return {
    workbookName: readRequiredString(record, 'workbookName'),
    sheetNames: readSheetNameArray(record, 'sheetNames'),
    dimensions: readRequiredArray(record, 'dimensions').map((dimension) => {
      const dimensionRecord = asRecord(dimension)
      const usedRange = parseOptionalUsedRange(dimensionRecord['usedRange'])
      const parsedDimension: PublicWorkbookCorpusCase['workbookMetadata']['dimensions'][number] = {
        sheetName: readRequiredSheetName(dimensionRecord, 'sheetName'),
        rowCount: readRequiredInteger(dimensionRecord, 'rowCount'),
        columnCount: readRequiredInteger(dimensionRecord, 'columnCount'),
        nonEmptyCellCount: readRequiredInteger(dimensionRecord, 'nonEmptyCellCount'),
      }
      if (usedRange !== undefined) {
        Object.assign(parsedDimension, { usedRange })
      }
      return parsedDimension
    }),
  }
}

function parseOptionalUsedRange(
  value: unknown,
): PublicWorkbookCorpusCase['workbookMetadata']['dimensions'][number]['usedRange'] | undefined {
  if (value === undefined) {
    return undefined
  }
  if (value === null) {
    return null
  }
  const record = asRecord(value)
  return {
    startRow: readRequiredInteger(record, 'startRow'),
    startColumn: readRequiredInteger(record, 'startColumn'),
    endRow: readRequiredInteger(record, 'endRow'),
    endColumn: readRequiredInteger(record, 'endColumn'),
  }
}

function parsePublicWorkbookValidationSummary(value: unknown): PublicWorkbookValidationSummary {
  const record = asRecord(value)
  return {
    importPassed: readRequiredBoolean(record, 'importPassed'),
    formulaOraclePassed: readRequiredBoolean(record, 'formulaOraclePassed'),
    formulaOracleComparisons: readRequiredInteger(record, 'formulaOracleComparisons'),
    formulaOracleMismatches: readStringArray(record, 'formulaOracleMismatches'),
    roundTripPassed: readRequiredBoolean(record, 'roundTripPassed'),
    structuralSmokePassed: readBooleanOrNull(record, 'structuralSmokePassed'),
  }
}

function parsePublicWorkbookSourceKind(value: string): PublicWorkbookSourceKind {
  switch (value) {
    case 'direct-url':
    case 'ckan-resource':
    case 'github-contents':
      return value
    default:
      throw new Error(`Unexpected public workbook source kind: ${value}`)
  }
}

function parsePublicWorkbookVerificationPhase(value: string): PublicWorkbookVerificationPhase {
  switch (value) {
    case 'read-cache':
    case 'inspect-footprint':
    case 'import-xlsx':
    case 'formula-oracle':
    case 'round-trip':
    case 'structural-smoke':
      return value
    default:
      throw new Error(`Unexpected public workbook verification phase: ${value}`)
  }
}

function parsePublicWorkbookCaseStatus(value: string): PublicWorkbookCaseStatus {
  switch (value) {
    case 'passed':
    case 'failed':
    case 'error':
    case 'unsupported':
      return value
    default:
      throw new Error(`Unexpected public workbook case status: ${value}`)
  }
}

function countFormulaOracleMatches(cases: readonly PublicWorkbookCorpusCase[]): number {
  return cases.reduce(
    (sum, entry) => sum + Math.max(0, entry.validation.formulaOracleComparisons - entry.validation.formulaOracleMismatches.length),
    0,
  )
}

function readRequiredString(value: Record<string, unknown>, key: string): string {
  const result = readString(value, key)
  if (!result) {
    throw new Error(`Expected ${key} to be a non-empty string`)
  }
  return result
}

function readRequiredSheetName(value: Record<string, unknown>, key: string): string {
  const fieldValue = value[key]
  if (typeof fieldValue !== 'string' || fieldValue.length === 0) {
    throw new Error(`Expected ${key} to be a non-empty sheet name`)
  }
  return fieldValue
}

function readOptionalString(value: Record<string, unknown>, key: string): string | undefined {
  return readString(value, key) ?? undefined
}

function readNullableString(value: Record<string, unknown>, key: string): string | null {
  const fieldValue = value[key]
  if (fieldValue === null || fieldValue === undefined) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${key} to be a string or null`)
  }
  return fieldValue.trim().length > 0 ? fieldValue.trim() : null
}

function readRequiredInteger(value: Record<string, unknown>, key: string): number {
  const fieldValue = value[key]
  if (!Number.isInteger(fieldValue) || fieldValue < 0) {
    throw new Error(`Expected ${key} to be a non-negative integer`)
  }
  return fieldValue
}

function readOptionalNonNegativeInteger(value: Record<string, unknown>, key: string): number | undefined {
  if (value[key] === undefined) {
    return undefined
  }
  return readRequiredInteger(value, key)
}

function readOptionalNonNegativeIntegerOrNull(value: Record<string, unknown>, key: string): number | null | undefined {
  if (value[key] === undefined) {
    return undefined
  }
  if (value[key] === null) {
    return null
  }
  return readRequiredInteger(value, key)
}

function readTargetWorkbookCount(value: Record<string, unknown>, key: string): number {
  const targetWorkbookCount = readRequiredInteger(value, key)
  validateTargetWorkbookCount(targetWorkbookCount)
  return targetWorkbookCount
}

function validateTargetWorkbookCount(targetWorkbookCount: number): void {
  if (!Number.isInteger(targetWorkbookCount) || targetWorkbookCount <= 0) {
    throw new Error('Public workbook corpus target workbook count must be a positive integer')
  }
}

function readRequiredBoolean(value: Record<string, unknown>, key: string): boolean {
  const fieldValue = value[key]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${key} to be a boolean`)
  }
  return fieldValue
}

function readBooleanOrNull(value: Record<string, unknown>, key: string): boolean | null {
  const fieldValue = value[key]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${key} to be a boolean or null`)
  }
  return fieldValue
}

function readRequiredArray(value: Record<string, unknown>, key: string): unknown[] {
  const fieldValue = value[key]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${key} to be an array`)
  }
  return fieldValue
}

function readStringArray(value: Record<string, unknown>, key: string): string[] {
  return readRequiredArray(value, key).map((entry, index) => {
    if (typeof entry !== 'string') {
      throw new Error(`Expected ${key}[${String(index)}] to be a string`)
    }
    return entry
  })
}

function readSheetNameArray(value: Record<string, unknown>, key: string): string[] {
  return readRequiredArray(value, key).map((entry, index) => {
    if (typeof entry !== 'string' || entry.length === 0) {
      throw new Error(`Expected ${key}[${String(index)}] to be a non-empty sheet name`)
    }
    return entry
  })
}

function readOptionalStringArray(value: Record<string, unknown>, key: string): string[] | undefined {
  if (value[key] === undefined) {
    return undefined
  }
  return readStringArray(value, key)
}

function readExpectedNumber<const Expected extends number>(value: Record<string, unknown>, key: string, expected: Expected): Expected {
  const actual = readRequiredInteger(value, key)
  if (actual !== expected) {
    throw new Error(`Expected ${key} to equal ${String(expected)}`)
  }
  return expected
}

function readExpectedString<const Expected extends string>(value: Record<string, unknown>, key: string, expected: Expected): Expected {
  const actual = readRequiredString(value, key)
  if (actual !== expected) {
    throw new Error(`Expected ${key} to equal ${expected}`)
  }
  return expected
}
