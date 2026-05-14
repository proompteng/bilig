import type { CellValue } from '../packages/protocol/src/types.js'

export type PublicWorkbookSourceKind = 'direct-url' | 'ckan-resource' | 'github-contents'
export type PublicWorkbookCaseStatus = 'passed' | 'failed' | 'error' | 'unsupported'

export interface PublicWorkbookLicenseEvidence {
  readonly spdxId: string | null
  readonly title: string
  readonly evidenceUrl: string | null
}

export interface PublicWorkbookSource {
  readonly id: string
  readonly kind: PublicWorkbookSourceKind
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly discoveredAt: string
  readonly license: PublicWorkbookLicenseEvidence
  readonly topicEvidence?: readonly string[]
  readonly portal?: string
  readonly datasetId?: string
  readonly resourceId?: string
}

export interface PublicWorkbookArtifact {
  readonly id: string
  readonly sourceId: string
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly cachePath: string
  readonly sha256: string
  readonly byteSize: number
  readonly workbookFingerprint: string
  readonly fetchedAt: string
  readonly license: PublicWorkbookLicenseEvidence
  readonly topicEvidence?: readonly string[]
}

export interface PublicWorkbookManifest {
  readonly schemaVersion: 1
  readonly corpus: 'public-workbook-corpus'
  readonly targetWorkbookCount: number
  readonly generatedAt: string
  readonly sources: readonly PublicWorkbookSource[]
  readonly artifacts: readonly PublicWorkbookArtifact[]
  readonly fetchState?: PublicWorkbookFetchState
}

export interface PublicWorkbookFetchState {
  readonly exhaustedSourceIds: readonly string[]
}

export interface PublicWorkbookFeatureCounts {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly chartCount: number
  readonly pivotCount: number
  readonly mergeCount: number
  readonly styleRangeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount: number
  readonly macroPayloadCount: number
  readonly warningCount: number
}

export interface PublicWorkbookValidationSummary {
  readonly importPassed: boolean
  readonly formulaOraclePassed: boolean
  readonly formulaOracleComparisons: number
  readonly formulaOracleMismatches: readonly string[]
  readonly roundTripPassed: boolean
  readonly structuralSmokePassed: boolean | null
}

export interface PublicWorkbookCorpusCase {
  readonly id: string
  readonly sourceId: string
  readonly sourceUrl: string
  readonly fileName: string
  readonly sha256: string
  readonly byteSize: number
  readonly license: PublicWorkbookLicenseEvidence
  readonly status: PublicWorkbookCaseStatus
  readonly passed: boolean
  readonly featureCounts: PublicWorkbookFeatureCounts
  readonly workbookMetadata: {
    readonly workbookName: string
    readonly sheetNames: readonly string[]
    readonly dimensions: readonly {
      readonly sheetName: string
      readonly rowCount: number
      readonly columnCount: number
      readonly nonEmptyCellCount: number
      readonly usedRange?: {
        readonly startRow: number
        readonly startColumn: number
        readonly endRow: number
        readonly endColumn: number
      } | null
    }[]
  }
  readonly validation: PublicWorkbookValidationSummary
  readonly unsupportedFeatureClassifications: readonly string[]
  readonly evidence: readonly string[]
}

export interface PublicWorkbookCorpusScorecard {
  readonly schemaVersion: 1
  readonly suite: 'public-workbook-corpus'
  readonly generatedAt: string
  readonly summary: {
    readonly targetWorkbookCount: number
    readonly sourceCount: number
    readonly cachedWorkbookCount: number
    readonly importedWorkbookCount: number
    readonly passedWorkbookCount: number
    readonly failedWorkbookCount: number
    readonly errorWorkbookCount: number
    readonly unsupportedWorkbookCount: number
    readonly formulaOracleComparisonCount: number
    readonly formulaOracleMatchCount: number
    readonly structuralSmokeRunCount: number
    readonly allCachedWorkbooksPassed: boolean
    readonly remainingToTarget: number
  }
  readonly cases: readonly PublicWorkbookCorpusCase[]
}

export interface FormulaOracle {
  readonly sheetName: string
  readonly address: string
  readonly expected: CellValue
}

export interface FormulaOracleValidationResult {
  readonly comparisons: number
  readonly mismatches: readonly string[]
}

export interface BuildScorecardArgs {
  readonly manifest: PublicWorkbookManifest
  readonly cacheDir: string
  readonly generatedAt?: string
  readonly manifestPath?: string
  readonly isolatedVerification?: boolean
  readonly structuralSmokeSampleLimit?: number
  readonly verifyConcurrency?: number
  readonly verifyTimeoutMs?: number
  readonly verifyMaxRssBytes?: number
  readonly verifyRssCheckIntervalMs?: number
  readonly verifyMaxCellCount?: number
  readonly reusableCases?: readonly PublicWorkbookCorpusCase[]
  readonly onCaseVerified?: (progress: PublicWorkbookCorpusVerificationProgress) => void
}

export interface PublicWorkbookCorpusVerificationProgress {
  readonly completedCount: number
  readonly totalCount: number
  readonly latestCase: PublicWorkbookCorpusCase
}

export interface DiscoverCkanArgs {
  readonly manifest: PublicWorkbookManifest
  readonly portalBases: readonly string[]
  readonly query: string
  readonly limit: number
  readonly rowsPerRequest: number
  readonly discoveredAt?: string
  readonly requiredTopic?: 'financial-workpapers'
}

export interface FetchCorpusArgs {
  readonly manifest: PublicWorkbookManifest
  readonly cacheDir: string
  readonly limit: number
  readonly fetchedAt?: string
  readonly maxBytes?: number
  readonly downloadTimeoutMs?: number
  readonly fetchBatchSize?: number
  readonly fetchConcurrency?: number
  readonly fingerprintTimeoutMs?: number
  readonly fingerprintMaxRssBytes?: number
  readonly fingerprintRssCheckIntervalMs?: number
  readonly isolatedFingerprinting?: boolean
  readonly onArtifactsCommitted?: (
    manifest: PublicWorkbookManifest,
    progress: PublicWorkbookCorpusFetchCheckpointProgress,
  ) => void | Promise<void>
  readonly sourceIds?: readonly string[]
}

export interface PublicWorkbookCorpusFetchCheckpointProgress {
  readonly artifactCount: number
  readonly exhaustedSourceCount: number
  readonly committedArtifactCount: number
  readonly exhaustedSourceDelta: number
  readonly failedSourceCount: number
  readonly duplicateHashSourceCount: number
  readonly duplicateFingerprintSourceCount: number
  readonly failedSourceSamples: readonly PublicWorkbookCorpusFetchFailureSample[]
}

export interface PublicWorkbookCorpusFetchFailureSample {
  readonly sourceId: string
  readonly fileName: string
  readonly error: string
}

export interface CkanPageRequest {
  readonly portalBase: string
  readonly query: string
  readonly rowsPerRequest: number
  readonly start: number
}

export interface CkanPageResult {
  readonly portalBase: string
  readonly packages: readonly Record<string, unknown>[]
}

export interface WorkbookDownloadResult {
  readonly source: PublicWorkbookSource
  readonly bytes: Uint8Array | null
  readonly sha256: string | null
  readonly workbookFingerprint: string | null
  readonly error: string | null
}
