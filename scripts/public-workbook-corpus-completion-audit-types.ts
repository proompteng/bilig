import type { PublicWorkbookCorpusUnsupportedClassificationCount } from './public-workbook-corpus-completion-audit-helpers.ts'

export type PublicWorkbookCorpusCompletionStatus = 'achieved' | 'active-not-achieved'

export interface PublicWorkbookCorpusCompletionAudit {
  readonly schemaVersion: 1
  readonly generatedAt: string
  readonly objective: string
  readonly completionVerdict: {
    readonly goalStatus: PublicWorkbookCorpusCompletionStatus
    readonly allChecklistItemsPassed: boolean
    readonly targetComplete: boolean
    readonly stopMarkerActive: boolean
    readonly nextCorpusRunRequiresExplicitResume: boolean
    readonly unmetRequirements: readonly string[]
  }
  readonly currentState: PublicWorkbookCorpusAuditState
  readonly secondaryFormulaCorpus: PublicWorkbookCorpusSecondaryFormulaCorpusStatus
  readonly nextActions: readonly PublicWorkbookCorpusAuditNextAction[]
  readonly checklist: readonly PublicWorkbookCorpusAuditChecklistItem[]
}

export interface PublicWorkbookCorpusAuditNextAction {
  readonly id: PublicWorkbookCorpusNextActionId
  readonly priority: number
  readonly reason: string
  readonly commands: readonly string[]
  readonly blockedCommands: readonly string[]
}

export interface PublicWorkbookCorpusAuditState {
  readonly targetWorkbookCount: number
  readonly financialWorkbookTargetCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly financialSourceCount: number
  readonly financialCachedArtifactCount: number
  readonly xlsxArtifactCount: number
  readonly nonXlsxArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly recordedFinancialManifestArtifactCount: number
  readonly recordedFinancialNonPassingCaseCount: number
  readonly missingCachedArtifactCount: number
  readonly missingVerificationCount: number
  readonly staleRecordedVerificationCount: number
  readonly missingFeatureWitnessCount: number
  readonly missingFeatureWitnesses: readonly string[]
  readonly recordedPassedCaseCount: number
  readonly recordedUnsupportedCaseCount: number
  readonly staleRecordedUnsupportedCaseCount: number
  readonly currentRecordedUnsupportedCaseCount: number
  readonly currentUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  readonly staleUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  readonly recordedFailedCaseCount: number
  readonly recordedErrorCaseCount: number
  readonly recordedFormulaOracleComparisonCount: number
  readonly recordedFormulaOracleMismatchCount: number
  readonly recordedStructuralSmokeRunCount: number
  readonly recordedRoundTripPassedCount: number
  readonly recordedRoundTripSkippedCount: number
  readonly recordedRoundTripFailureCount: number
}

export interface PublicWorkbookCorpusAuditChecklistItem {
  readonly id: PublicWorkbookCorpusRequirementId
  readonly priority: number
  readonly promptRequirement: string
  readonly passed: boolean
  readonly evidence: readonly string[]
  readonly evidenceArtifacts: readonly string[]
  readonly checkCommands: readonly string[]
  readonly gaps: readonly string[]
}

export interface PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
  readonly artifact: string
  readonly artifactPresent: boolean
  readonly suite: string | null
  readonly resultCount: number
  readonly comparableCount: number
  readonly workpaperWins: number
  readonly hyperformulaWins: number
  readonly comparableVerificationEquivalentCount: number
  readonly allComparableVerificationEquivalent: boolean
  readonly parseError: string | null
}

export type PublicWorkbookCorpusRequirementId =
  | 'download-10000-public-spreadsheets'
  | 'financial-accounting-workpapers-5000'
  | 'source-license-hash-metadata-manifest'
  | 'hash-and-structure-dedupe'
  | 'import-every-workbook'
  | 'validate-workbook-features'
  | 'formula-recalc-oracle'
  | 'structural-smoke'
  | 'roundtrip-supported-workbooks'
  | 'scorecard-all-10000'
  | 'ci-offline-cached-mode'
  | 'unsupported-features-evidence'
  | 'hyperformula-secondary-corpus'

export type PublicWorkbookCorpusNextActionId =
  | 'resume-public-corpus-ingest'
  | 'verify-missing-cached-artifacts'
  | 'refresh-stale-verification-evidence'
  | 'fill-feature-witnesses'
  | 'resume-financial-workpapers'
  | 'inspect-non-passing-scorecard-cases'
