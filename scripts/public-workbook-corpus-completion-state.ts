import { planPublicWorkbookCorpusFetch, type PublicWorkbookCorpusFetchPlan } from './public-workbook-corpus-fetch.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import type { PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import {
  buildFeatureWitnessCoverage,
  buildUnsupportedCaseSummary,
  financialWorkbookTargetCount,
  hasFinancialTopicEvidence,
} from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'
import type { PublicWorkbookCorpusAuditState } from './public-workbook-corpus-completion-audit-types.ts'

const roundTripSkippedEvidencePrefix = 'Round-trip projection skipped because'

export function buildPublicWorkbookCorpusAuditState(args: {
  readonly financialManifest: PublicWorkbookManifest | null
  readonly financialRecordedCases: readonly PublicWorkbookCorpusCase[]
  readonly manifest: PublicWorkbookManifest | null
  readonly recordedCases: readonly PublicWorkbookCorpusCase[]
  readonly status: PublicWorkbookCorpusStatus
}): PublicWorkbookCorpusAuditState {
  const financialArtifactCandidates = args.financialManifest ? args.financialManifest.artifacts : (args.manifest?.artifacts ?? [])
  const financialSourceCandidates = args.financialManifest ? args.financialManifest.sources : (args.manifest?.sources ?? [])
  const financialArtifacts = financialArtifactCandidates.filter(hasFinancialTopicEvidence)
  const financialSources = financialSourceCandidates.filter(hasFinancialTopicEvidence)
  const financialCaseCandidates = args.financialManifest ? args.financialRecordedCases : args.recordedCases
  const recordedCasesById = new Map(financialCaseCandidates.map((entry) => [entry.id, entry]))
  const fetchPlan = planPublicWorkbookCorpusFetchForAudit(args.manifest, args.status.targetWorkbookCount)
  const missingFeatureWitnesses = buildFeatureWitnessCoverage(args.recordedCases)
    .filter((entry) => entry.witnessCaseCount === 0)
    .map((entry) => entry.label)
  const unsupportedCaseSummary = buildUnsupportedCaseSummary(args.recordedCases)
  const recordedFinancialCases = financialArtifacts.flatMap((artifact) => {
    const candidate = recordedCasesById.get(artifact.id)
    return candidate && publicWorkbookCorpusCaseMatchesArtifact(candidate, artifact) ? [candidate] : []
  })
  return {
    targetWorkbookCount: args.status.targetWorkbookCount,
    financialWorkbookTargetCount: financialWorkbookTargetCount(args.status.targetWorkbookCount),
    sourceCount: args.status.sourceCount,
    fetchCandidateSourceCount: fetchPlan?.candidateSourceCount ?? 0,
    fetchCandidateSourceDeficitCount:
      fetchPlan?.candidateSourceDeficitCount ?? Math.max(0, args.status.targetWorkbookCount - args.status.sourceCount),
    fetchTargetReachableFromKnownCandidates:
      fetchPlan?.targetReachableFromKnownCandidates ?? args.status.cachedArtifactCount >= args.status.targetWorkbookCount,
    recommendedDiscoveryLimit: fetchPlan?.recommendedDiscoveryLimit ?? args.status.targetWorkbookCount,
    cachedArtifactCount: args.status.cachedArtifactCount,
    financialSourceCount: financialSources.length,
    financialCachedArtifactCount: financialArtifacts.length,
    financialSourceWithoutTopicEvidenceCount: args.financialManifest ? financialSourceCandidates.length - financialSources.length : 0,
    financialArtifactWithoutTopicEvidenceCount: args.financialManifest ? financialArtifactCandidates.length - financialArtifacts.length : 0,
    xlsxArtifactCount: (args.manifest?.artifacts ?? []).filter((entry) => isXlsxArtifact(entry.fileName, entry.cachePath)).length,
    nonXlsxArtifactCount: (args.manifest?.artifacts ?? []).filter((entry) => !isXlsxArtifact(entry.fileName, entry.cachePath)).length,
    scorecardCaseCount: args.status.scorecardCaseCount,
    checkpointCaseCount: args.status.checkpointCaseCount,
    recordedManifestArtifactCount: args.status.recordedManifestArtifactCount,
    recordedFinancialManifestArtifactCount: recordedFinancialCases.length,
    recordedFinancialNonPassingCaseCount: recordedFinancialCases.filter((entry) => !entry.passed).length,
    missingCachedArtifactCount: Math.max(0, args.status.targetWorkbookCount - args.status.cachedArtifactCount),
    missingVerificationCount: args.status.missingManifestArtifactCount,
    staleRecordedVerificationCount: args.status.staleRecordedVerificationCount,
    missingFeatureWitnessCount: missingFeatureWitnesses.length,
    missingFeatureWitnesses,
    recordedPassedCaseCount: args.status.recordedPassedCaseCount,
    recordedUnsupportedCaseCount: args.status.recordedUnsupportedCaseCount,
    ...unsupportedCaseSummary,
    recordedFailedCaseCount: args.status.recordedFailedCaseCount,
    recordedErrorCaseCount: args.status.recordedErrorCaseCount,
    recordedFormulaOracleComparisonCount: args.recordedCases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0),
    recordedFormulaOracleMismatchCount: args.recordedCases.reduce((sum, entry) => sum + entry.validation.formulaOracleMismatches.length, 0),
    recordedStructuralSmokeRunCount: args.recordedCases.filter((entry) => entry.validation.structuralSmokePassed !== null).length,
    recordedRoundTripPassedCount: args.recordedCases.filter((entry) => isSupportedRoundTripSuccess(entry)).length,
    recordedRoundTripSkippedCount: args.recordedCases.filter((entry) => hasRoundTripSkippedEvidence(entry)).length,
    recordedRoundTripFailureCount: args.recordedCases.filter(
      (entry) => !entry.validation.roundTripPassed && entry.status !== 'unsupported' && !hasRoundTripSkippedEvidence(entry),
    ).length,
  }
}

function isXlsxArtifact(fileName: string, cachePath: string): boolean {
  return fileName.toLowerCase().endsWith('.xlsx') || cachePath.toLowerCase().endsWith('.xlsx')
}

function planPublicWorkbookCorpusFetchForAudit(
  manifest: PublicWorkbookManifest | null,
  targetWorkbookCount: number,
): PublicWorkbookCorpusFetchPlan | null {
  if (!manifest) {
    return null
  }
  try {
    return planPublicWorkbookCorpusFetch({ manifest, limit: targetWorkbookCount, sampleLimit: 0 })
  } catch {
    return null
  }
}

function isSupportedRoundTripSuccess(entry: PublicWorkbookCorpusCase): boolean {
  return entry.validation.roundTripPassed && entry.status !== 'unsupported' && !hasRoundTripSkippedEvidence(entry)
}

function hasRoundTripSkippedEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.some((line) => line.startsWith(roundTripSkippedEvidencePrefix))
}
