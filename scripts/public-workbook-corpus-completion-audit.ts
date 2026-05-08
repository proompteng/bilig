#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { hasUsableLicenseEvidence, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import { readPublicWorkbookCorpusStatus, type PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import { readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { auditPublicWorkbookCorpusCiOfflineCachedMode } from './public-workbook-corpus-ci-offline-audit.ts'
import type { PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

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
  readonly checklist: readonly PublicWorkbookCorpusAuditChecklistItem[]
}

export interface PublicWorkbookCorpusAuditState {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly missingCachedArtifactCount: number
  readonly missingVerificationCount: number
  readonly recordedPassedCaseCount: number
  readonly recordedUnsupportedCaseCount: number
  readonly recordedFailedCaseCount: number
  readonly recordedErrorCaseCount: number
  readonly recordedFormulaOracleComparisonCount: number
  readonly recordedFormulaOracleMismatchCount: number
  readonly recordedStructuralSmokeRunCount: number
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

type PublicWorkbookCorpusRequirementId =
  | 'download-10000-public-spreadsheets'
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

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

const objective =
  'Build a 10,000-spreadsheet legally usable public workbook corpus, verify every workbook through repeatable bilig correctness checks, and keep unsupported workbook behavior classified with evidence.'

const baselineScorecardArtifact = 'packages/benchmarks/baselines/public-workbook-corpus-scorecard.json'
const hyperFormulaSecondaryCorpusArtifact = 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json'
const manifestArtifact = '.cache/public-workbook-corpus/manifest.json'
const checkpointArtifact = '.cache/public-workbook-corpus/verification-checkpoint.json'
const requiredFeatureWitnesses = [
  { id: 'formulas', label: 'formulas', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.formulaCellCount },
  { id: 'values', label: 'values', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.valueCellCount },
  { id: 'names', label: 'defined names', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.definedNameCount },
  { id: 'tables', label: 'tables', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.tableCount },
  { id: 'charts', label: 'charts', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.chartCount },
  { id: 'pivots', label: 'pivots', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.pivotCount },
  { id: 'styles', label: 'styles', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.styleRangeCount },
  { id: 'merged ranges', label: 'merged ranges', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.mergeCount },
  {
    id: 'conditional formats',
    label: 'conditional formats',
    count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.conditionalFormatCount,
  },
] as const

function main(): void {
  const audit = buildPublicWorkbookCorpusCompletionAuditFromArgs()
  const requireComplete = readFlagArg('--require-complete')
  if (readFlagArg('--check')) {
    const findings = validatePublicWorkbookCorpusCompletionAudit(audit, { requireComplete })
    if (findings.length > 0) {
      process.stderr.write(`Public workbook corpus completion audit is invalid: ${findings.join('; ')}\n`)
      process.exitCode = 1
      return
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          schemaVersion: audit.schemaVersion,
          goalStatus: audit.completionVerdict.goalStatus,
          allChecklistItemsPassed: audit.completionVerdict.allChecklistItemsPassed,
          targetComplete: audit.completionVerdict.targetComplete,
          checklistItemCount: audit.checklist.length,
          unmetRequirementCount: audit.completionVerdict.unmetRequirements.length,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  process.stdout.write(`${JSON.stringify(audit, null, 2)}\n`)
}

export function buildPublicWorkbookCorpusCompletionAuditFromArgs(): PublicWorkbookCorpusCompletionAudit {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const manifest = existsSync(manifestPath) ? parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8'))) : null
  const recordedCases = readRecordedCases({ manifest, scorecardPath, verifyCheckpointPath })
  return buildPublicWorkbookCorpusCompletionAudit({
    generatedAt,
    hyperformulaSecondaryCorpus: readHyperFormulaSecondaryCorpus(resolve(rootDir, hyperFormulaSecondaryCorpusArtifact)),
    manifest,
    recordedCases,
    status: readPublicWorkbookCorpusStatus({
      cacheDir,
      manifestPath,
      scorecardPath,
      verifyCheckpointPath,
      corpusRunStopMarkerPath: stopMarkerPath,
    }),
    stopMarkerActive: existsSync(stopMarkerPath),
  })
}

export function buildPublicWorkbookCorpusCompletionAudit(args: {
  readonly generatedAt: string
  readonly hyperformulaSecondaryCorpus?: PublicWorkbookCorpusSecondaryFormulaCorpusStatus
  readonly manifest: PublicWorkbookManifest | null
  readonly recordedCases: readonly PublicWorkbookCorpusCase[]
  readonly status: PublicWorkbookCorpusStatus
  readonly stopMarkerActive: boolean
}): PublicWorkbookCorpusCompletionAudit {
  const currentState = buildAuditState(args.status, args.recordedCases)
  const secondaryFormulaCorpus = args.hyperformulaSecondaryCorpus ?? missingHyperFormulaSecondaryCorpus()
  const context = {
    currentState,
    hyperformulaSecondaryCorpus: secondaryFormulaCorpus,
    manifest: args.manifest,
    recordedCases: args.recordedCases,
    status: args.status,
  }
  const checklist = requirementBuilders.map((builder) => builder(context))
  const allChecklistItemsPassed = checklist.every((entry) => entry.passed)
  const unmetRequirements = checklist.flatMap((entry) => (entry.passed ? [] : entry.gaps.map((gap) => `${entry.id}: ${gap}`)))
  return {
    schemaVersion: 1,
    generatedAt: args.generatedAt,
    objective,
    completionVerdict: {
      goalStatus: allChecklistItemsPassed && args.status.targetComplete ? 'achieved' : 'active-not-achieved',
      allChecklistItemsPassed,
      targetComplete: args.status.targetComplete,
      stopMarkerActive: args.stopMarkerActive,
      nextCorpusRunRequiresExplicitResume: args.stopMarkerActive && !args.status.targetComplete,
      unmetRequirements,
    },
    currentState,
    secondaryFormulaCorpus,
    checklist,
  }
}

export function validatePublicWorkbookCorpusCompletionAudit(
  audit: PublicWorkbookCorpusCompletionAudit,
  options: { readonly requireComplete?: boolean } = {},
): string[] {
  const findings: string[] = []
  const checklistIds = audit.checklist.map((entry) => entry.id)
  if (JSON.stringify(checklistIds) !== JSON.stringify(requiredRequirementIds)) {
    findings.push(`checklist ids do not match the public corpus objective: ${checklistIds.join(', ')}`)
  }
  const duplicateIds = checklistIds.filter((id, index) => checklistIds.indexOf(id) !== index)
  if (duplicateIds.length > 0) {
    findings.push(`duplicate checklist ids: ${[...new Set(duplicateIds)].join(', ')}`)
  }
  for (const item of audit.checklist) {
    if (!item.promptRequirement.trim()) {
      findings.push(`${item.id} is missing the mapped prompt requirement`)
    }
    if (item.evidence.length === 0) {
      findings.push(`${item.id} has no evidence`)
    }
    if (item.evidenceArtifacts.length === 0) {
      findings.push(`${item.id} has no evidence artifacts`)
    }
    for (const artifact of item.evidenceArtifacts) {
      if (isRepoEvidenceArtifact(artifact) && !existsSync(resolve(rootDir, artifact))) {
        findings.push(`${item.id} evidence artifact does not exist: ${artifact}`)
      }
    }
    if (item.checkCommands.length === 0) {
      findings.push(`${item.id} has no check commands`)
    }
    for (const command of item.checkCommands) {
      const scriptName = pnpmScriptName(command)
      if (scriptName && !packageScripts().has(scriptName)) {
        findings.push(`${item.id} check command references missing package script: ${scriptName}`)
      }
    }
    if (item.passed && item.gaps.length > 0) {
      findings.push(`${item.id} is passed while still reporting gaps`)
    }
    if (!item.passed && item.gaps.length === 0) {
      findings.push(`${item.id} is failed without an explicit gap`)
    }
  }
  const allChecklistItemsPassed = audit.checklist.every((entry) => entry.passed)
  if (audit.completionVerdict.allChecklistItemsPassed !== allChecklistItemsPassed) {
    findings.push('completion verdict allChecklistItemsPassed does not match checklist state')
  }
  if (audit.completionVerdict.goalStatus === 'achieved') {
    if (!allChecklistItemsPassed || !audit.completionVerdict.targetComplete || audit.completionVerdict.unmetRequirements.length > 0) {
      findings.push('goal is achieved without complete checklist, target, and unmet-requirement evidence')
    }
  }
  if (audit.completionVerdict.goalStatus === 'active-not-achieved' && audit.completionVerdict.unmetRequirements.length === 0) {
    findings.push('active goal has no unmet requirement evidence')
  }
  if (audit.completionVerdict.nextCorpusRunRequiresExplicitResume && !audit.completionVerdict.stopMarkerActive) {
    findings.push('explicit corpus resume is required without an active stop marker')
  }
  if (options.requireComplete && audit.completionVerdict.goalStatus !== 'achieved') {
    findings.push(`public workbook corpus goal is not achieved: ${audit.completionVerdict.unmetRequirements.join('; ')}`)
  }
  return findings
}

interface RequirementContext {
  readonly currentState: PublicWorkbookCorpusAuditState
  readonly hyperformulaSecondaryCorpus: PublicWorkbookCorpusSecondaryFormulaCorpusStatus
  readonly manifest: PublicWorkbookManifest | null
  readonly recordedCases: readonly PublicWorkbookCorpusCase[]
  readonly status: PublicWorkbookCorpusStatus
}

const requiredRequirementIds: readonly PublicWorkbookCorpusRequirementId[] = [
  'download-10000-public-spreadsheets',
  'source-license-hash-metadata-manifest',
  'hash-and-structure-dedupe',
  'import-every-workbook',
  'validate-workbook-features',
  'formula-recalc-oracle',
  'structural-smoke',
  'roundtrip-supported-workbooks',
  'scorecard-all-10000',
  'ci-offline-cached-mode',
  'unsupported-features-evidence',
  'hyperformula-secondary-corpus',
]

const requirementBuilders: readonly ((context: RequirementContext) => PublicWorkbookCorpusAuditChecklistItem)[] = [
  (context) =>
    checklistItem({
      id: 'download-10000-public-spreadsheets',
      priority: 1,
      promptRequirement: 'Download 10,000 legally usable public spreadsheet files, prioritizing .xlsx.',
      passed:
        context.currentState.sourceCount >= context.currentState.targetWorkbookCount &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `target workbooks: ${String(context.currentState.targetWorkbookCount)}`,
        `discovered sources: ${String(context.currentState.sourceCount)}`,
        `cached artifacts: ${String(context.currentState.cachedArtifactCount)}`,
      ],
      gaps: [
        ...countGap(context.currentState.sourceCount, context.currentState.targetWorkbookCount, 'discovered sources below target'),
        ...countGap(context.currentState.cachedArtifactCount, context.currentState.targetWorkbookCount, 'cached artifacts below target'),
      ],
    }),
  (context) => {
    const artifacts = context.manifest?.artifacts ?? []
    const sources = context.manifest?.sources ?? []
    const sourceLicenseGapCount = sources.filter((entry) => !hasUsableLicenseEvidence(entry.license)).length
    const artifactLicenseGapCount = artifacts.filter((entry) => !hasUsableLicenseEvidence(entry.license)).length
    const artifactMetadataGapCount = artifacts.filter(
      (entry) => !entry.sourceUrl || !entry.downloadUrl || !entry.fetchedAt || !entry.sha256 || entry.byteSize <= 0,
    ).length
    const artifactCachePathGapCount = artifacts.filter((entry) => !isHashAddressedCachePath(entry.cachePath, entry.sha256)).length
    const metadataCaseGapCount = context.recordedCases.filter(
      (entry) => !hasWorkbookMetadata(entry) && !isResourceLimitedUnsupportedCase(entry),
    ).length
    const provenanceCaseGapCount = context.recordedCases.filter((entry) => !hasRecordedProvenanceEvidence(entry)).length
    const cacheIntegrityFailureCount = context.recordedCases.filter(hasCacheIntegrityFailureEvidence).length
    const resourceLimitedMetadataUnavailableCount = context.recordedCases.filter(
      (entry) => !hasWorkbookMetadata(entry) && isResourceLimitedUnsupportedCase(entry),
    ).length
    return checklistItem({
      id: 'source-license-hash-metadata-manifest',
      priority: 1,
      promptRequirement: 'Record source URL, license/usage evidence, fetch timestamp, file hash, size, and workbook metadata.',
      passed:
        Boolean(context.manifest) &&
        sourceLicenseGapCount === 0 &&
        artifactLicenseGapCount === 0 &&
        artifactMetadataGapCount === 0 &&
        artifactCachePathGapCount === 0 &&
        metadataCaseGapCount === 0 &&
        provenanceCaseGapCount === 0 &&
        cacheIntegrityFailureCount === 0 &&
        context.status.recordedCoversManifest &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `manifest present: ${String(Boolean(context.manifest))}`,
        `manifest sources with license evidence: ${String(sources.length - sourceLicenseGapCount)}/${String(sources.length)}`,
        `artifacts with license evidence: ${String(artifacts.length - artifactLicenseGapCount)}/${String(artifacts.length)}`,
        `hash-addressed artifact cache paths: ${String(artifacts.length - artifactCachePathGapCount)}/${String(artifacts.length)}`,
        `recorded source/license/hash evidence cases: ${String(context.recordedCases.length - provenanceCaseGapCount)}/${String(
          context.recordedCases.length,
        )}`,
        `recorded cache integrity failures: ${String(cacheIntegrityFailureCount)}`,
        `recorded workbook metadata cases: ${String(context.recordedCases.length - metadataCaseGapCount)}/${String(
          context.currentState.cachedArtifactCount,
        )}`,
        `resource-limited unsupported cases with metadata unavailable: ${String(resourceLimitedMetadataUnavailableCount)}`,
      ],
      gaps: [
        ...(context.manifest ? [] : ['manifest artifact is missing']),
        ...(sourceLicenseGapCount === 0 ? [] : [`sources missing usable license evidence: ${String(sourceLicenseGapCount)}`]),
        ...(artifactLicenseGapCount === 0 ? [] : [`artifacts missing usable license evidence: ${String(artifactLicenseGapCount)}`]),
        ...(artifactMetadataGapCount === 0 ? [] : [`artifacts missing source/hash/fetch metadata: ${String(artifactMetadataGapCount)}`]),
        ...(artifactCachePathGapCount === 0 ? [] : [`artifact cache paths not hash-addressed: ${String(artifactCachePathGapCount)}`]),
        ...(metadataCaseGapCount === 0 ? [] : [`recorded cases missing workbook metadata: ${String(metadataCaseGapCount)}`]),
        ...(provenanceCaseGapCount === 0 ? [] : [`recorded cases missing source/license/hash evidence: ${String(provenanceCaseGapCount)}`]),
        ...(cacheIntegrityFailureCount === 0 ? [] : [`recorded cache integrity failures: ${String(cacheIntegrityFailureCount)}`]),
        ...(context.status.recordedCoversManifest
          ? []
          : [
              `recorded verification cases below cached artifacts: ${String(
                context.currentState.recordedManifestArtifactCount,
              )}/${String(context.currentState.cachedArtifactCount)}`,
            ]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'metadata does not yet cover target artifacts',
        ),
      ],
    })
  },
  (context) => {
    const artifacts = context.manifest?.artifacts ?? []
    return checklistItem({
      id: 'hash-and-structure-dedupe',
      priority: 1,
      promptRequirement: 'Deduplicate files by hash and workbook structure.',
      passed:
        Boolean(context.manifest) &&
        artifacts.length === new Set(artifacts.map((entry) => entry.sha256)).size &&
        artifacts.length === new Set(artifacts.map((entry) => entry.workbookFingerprint)).size,
      evidence: [
        `manifest present: ${String(Boolean(context.manifest))}`,
        `unique hashes: ${String(new Set(artifacts.map((entry) => entry.sha256)).size)}/${String(artifacts.length)}`,
        `unique workbook fingerprints: ${String(new Set(artifacts.map((entry) => entry.workbookFingerprint)).size)}/${String(
          artifacts.length,
        )}`,
      ],
      gaps: [
        ...(context.manifest ? [] : ['manifest artifact is missing']),
        ...duplicateGap(
          artifacts.map((entry) => entry.sha256),
          'duplicate artifact hashes',
        ),
        ...duplicateGap(
          artifacts.map((entry) => entry.workbookFingerprint),
          'duplicate workbook structure fingerprints',
        ),
      ],
    })
  },
  (context) => {
    const failedImportCount = context.recordedCases.filter(
      (entry) => !entry.validation.importPassed && entry.status !== 'unsupported',
    ).length
    return checklistItem({
      id: 'import-every-workbook',
      priority: 3,
      promptRequirement: 'Import every cached workbook into bilig or classify unsupported/resource-limited cases with evidence.',
      passed:
        context.status.recordedCoversManifest &&
        failedImportCount === 0 &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `recorded verification cases: ${String(context.currentState.recordedManifestArtifactCount)}/${String(
          context.currentState.cachedArtifactCount,
        )}`,
        `non-unsupported import failures: ${String(failedImportCount)}`,
      ],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing import evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...(failedImportCount === 0 ? [] : [`non-unsupported import failures: ${String(failedImportCount)}`]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'import evidence does not yet cover target artifacts',
        ),
      ],
    })
  },
  (context) => {
    const featureValidationGapCount = context.recordedCases.filter((entry) => !hasFeatureValidationEvidence(entry)).length
    const usedRangeGapCount = context.recordedCases.filter((entry) => !hasUsedRangeValidationEvidence(entry)).length
    const featureWitnessCoverage = buildFeatureWitnessCoverage(context.recordedCases)
    const missingFeatureWitnesses = featureWitnessCoverage.filter((entry) => entry.witnessCaseCount === 0)
    return checklistItem({
      id: 'validate-workbook-features',
      priority: 3,
      promptRequirement:
        'Validate sheets, dimensions, used ranges, formulas, values, names, tables, charts, pivots, styles, merged ranges, conditional formats, and unsupported feature classifications.',
      passed:
        context.status.recordedCoversManifest &&
        featureValidationGapCount === 0 &&
        usedRangeGapCount === 0 &&
        missingFeatureWitnesses.length === 0 &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `feature-count fields validated by recorded cases: ${String(context.recordedCases.filter(hasFeatureValidationEvidence).length)}/${String(
          context.currentState.cachedArtifactCount,
        )}`,
        ...featureWitnessCoverage.map(
          (entry) => `${entry.label} witnessed cases: ${String(entry.witnessCaseCount)}; total recorded count: ${String(entry.totalCount)}`,
        ),
        `recorded unsupported cases: ${String(context.currentState.recordedUnsupportedCaseCount)}`,
      ],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing feature validation evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...(featureValidationGapCount === 0
          ? []
          : [`recorded cases with incomplete feature validation evidence: ${String(featureValidationGapCount)}`]),
        ...(usedRangeGapCount === 0 ? [] : [`recorded cases missing explicit used-range evidence: ${String(usedRangeGapCount)}`]),
        ...missingFeatureWitnesses.map((entry) => `no recorded ${entry.label} witness in corpus evidence`),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'feature validation does not yet cover target artifacts',
        ),
      ],
    })
  },
  (context) =>
    checklistItem({
      id: 'formula-recalc-oracle',
      priority: 4,
      promptRequirement: 'Recalculate formulas and compare stable expected values where source values provide an oracle.',
      passed:
        context.status.recordedCoversManifest &&
        context.currentState.recordedFormulaOracleComparisonCount > 0 &&
        context.currentState.recordedFormulaOracleMismatchCount === 0 &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `formula oracle comparisons: ${String(context.currentState.recordedFormulaOracleComparisonCount)}`,
        `formula oracle mismatches: ${String(context.currentState.recordedFormulaOracleMismatchCount)}`,
      ],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing formula oracle evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...(context.currentState.recordedFormulaOracleComparisonCount > 0 ? [] : ['no formula oracle comparisons recorded']),
        ...(context.currentState.recordedFormulaOracleMismatchCount === 0
          ? []
          : [`formula oracle mismatches: ${String(context.currentState.recordedFormulaOracleMismatchCount)}`]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'formula oracle evidence does not yet cover target artifacts',
        ),
      ],
    }),
  (context) => {
    const failedSmokeCount = context.recordedCases.filter((entry) => entry.validation.structuralSmokePassed === false).length
    return checklistItem({
      id: 'structural-smoke',
      priority: 5,
      promptRequirement: 'Run structural smoke operations on representative workbooks.',
      passed:
        context.status.recordedCoversManifest &&
        context.currentState.recordedStructuralSmokeRunCount > 0 &&
        failedSmokeCount === 0 &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `structural smoke run count: ${String(context.currentState.recordedStructuralSmokeRunCount)}`,
        `failed structural smoke cases: ${String(failedSmokeCount)}`,
      ],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing structural smoke eligibility evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...(context.currentState.recordedStructuralSmokeRunCount > 0 ? [] : ['no structural smoke operations recorded']),
        ...(failedSmokeCount === 0 ? [] : [`failed structural smoke cases: ${String(failedSmokeCount)}`]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'structural smoke evidence does not yet cover target artifacts',
        ),
      ],
    })
  },
  (context) =>
    checklistItem({
      id: 'roundtrip-supported-workbooks',
      priority: 5,
      promptRequirement: 'Round-trip export/import every supported workbook.',
      passed:
        context.status.recordedCoversManifest &&
        context.currentState.recordedRoundTripFailureCount === 0 &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [`round-trip failures among recorded cases: ${String(context.currentState.recordedRoundTripFailureCount)}`],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing round-trip evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...(context.currentState.recordedRoundTripFailureCount === 0
          ? []
          : [`round-trip failures: ${String(context.currentState.recordedRoundTripFailureCount)}`]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'round-trip evidence does not yet cover target artifacts',
        ),
      ],
    }),
  (context) =>
    checklistItem({
      id: 'scorecard-all-10000',
      priority: 6,
      promptRequirement: 'Produce a pass/fail/error/unsupported scorecard covering all 10,000 workbooks.',
      passed:
        context.status.scorecardCoversManifest &&
        context.currentState.recordedFailedCaseCount === 0 &&
        context.currentState.recordedErrorCaseCount === 0 &&
        context.currentState.scorecardCaseCount >= context.currentState.targetWorkbookCount &&
        context.status.targetComplete,
      evidence: [
        `scorecard cases: ${String(context.currentState.scorecardCaseCount)}/${String(context.currentState.targetWorkbookCount)}`,
        `scorecard passed cases: ${String(context.currentState.recordedPassedCaseCount)}`,
        `scorecard failed cases: ${String(context.currentState.recordedFailedCaseCount)}`,
        `scorecard error cases: ${String(context.currentState.recordedErrorCaseCount)}`,
        `scorecard unsupported cases: ${String(context.currentState.recordedUnsupportedCaseCount)}`,
        `scorecard covers manifest: ${String(context.status.scorecardCoversManifest)}`,
        `target complete: ${String(context.status.targetComplete)}`,
      ],
      gaps: [
        ...(context.status.scorecardCoversManifest
          ? []
          : [
              `scorecard cases do not cover manifest artifacts: ${String(context.currentState.scorecardCaseCount)}/${String(
                context.currentState.cachedArtifactCount,
              )}`,
            ]),
        ...(context.currentState.recordedFailedCaseCount === 0
          ? []
          : [`failed scorecard cases: ${String(context.currentState.recordedFailedCaseCount)}`]),
        ...(context.currentState.recordedErrorCaseCount === 0
          ? []
          : [`error scorecard cases: ${String(context.currentState.recordedErrorCaseCount)}`]),
        ...countGap(context.currentState.scorecardCaseCount, context.currentState.targetWorkbookCount, 'scorecard cases below target'),
        ...(context.status.targetComplete ? [] : context.status.gaps),
      ],
    }),
  (_context) => {
    const ciOfflineCachedMode = auditPublicWorkbookCorpusCiOfflineCachedMode({
      scripts: packageScripts(),
      ciSource: readFileSync(resolve(rootDir, 'scripts', 'run-ci.ts'), 'utf8'),
    })
    return checklistItem({
      id: 'ci-offline-cached-mode',
      priority: 6,
      promptRequirement: 'Add CI-friendly focused gates plus an offline cached corpus mode.',
      passed: ciOfflineCachedMode.passed,
      evidence: ciOfflineCachedMode.evidence,
      evidenceArtifacts: [manifestArtifact, baselineScorecardArtifact, checkpointArtifact, 'package.json', 'scripts/run-ci.ts'],
      checkCommands: [
        'pnpm public-workbook-corpus:check:offline',
        'pnpm public-workbook-corpus:resume-plan:check',
        'pnpm public-workbook-corpus:completion-audit:check',
        'pnpm test:correctness:corpus',
      ],
      gaps: ciOfflineCachedMode.gaps,
    })
  },
  (context) => {
    const unsupportedCases = context.recordedCases.filter((entry) => entry.status === 'unsupported')
    const unsupportedWithoutEvidenceCount = unsupportedCases.filter(
      (entry) => entry.unsupportedFeatureClassifications.length === 0 || entry.evidence.length === 0,
    ).length
    return checklistItem({
      id: 'unsupported-features-evidence',
      priority: 6,
      promptRequirement: 'Keep unsupported features classified with evidence instead of silently passing.',
      passed: unsupportedWithoutEvidenceCount === 0 && context.currentState.recordedUnsupportedCaseCount === unsupportedCases.length,
      evidence: [
        `unsupported cases: ${String(unsupportedCases.length)}`,
        `unsupported cases missing classifications/evidence: ${String(unsupportedWithoutEvidenceCount)}`,
      ],
      gaps:
        unsupportedWithoutEvidenceCount === 0
          ? []
          : [`unsupported cases missing classifications/evidence: ${String(unsupportedWithoutEvidenceCount)}`],
    })
  },
  (context) =>
    checklistItem({
      id: 'hyperformula-secondary-corpus',
      priority: 7,
      promptRequirement: 'Fold HyperFormula parity cases into the same reporting system as a secondary formula-behavior corpus.',
      passed:
        context.hyperformulaSecondaryCorpus.artifactPresent &&
        context.hyperformulaSecondaryCorpus.suite === 'workpaper-vs-hyperformula' &&
        context.hyperformulaSecondaryCorpus.comparableCount > 0 &&
        context.hyperformulaSecondaryCorpus.allComparableVerificationEquivalent &&
        hasPackageScripts(['workpaper:parity:check', 'workpaper:bench:competitive:check', 'public-workbook-corpus:completion-audit:check']),
      evidence: [
        `secondary artifact present: ${String(context.hyperformulaSecondaryCorpus.artifactPresent)}`,
        `secondary suite: ${context.hyperformulaSecondaryCorpus.suite ?? 'missing'}`,
        `secondary result count: ${String(context.hyperformulaSecondaryCorpus.resultCount)}`,
        `secondary comparable parity cases: ${String(context.hyperformulaSecondaryCorpus.comparableCount)}`,
        `secondary comparable verification-equivalent cases: ${String(
          context.hyperformulaSecondaryCorpus.comparableVerificationEquivalentCount,
        )}/${String(context.hyperformulaSecondaryCorpus.comparableCount)}`,
        `WorkPaper wins in secondary parity artifact: ${String(context.hyperformulaSecondaryCorpus.workpaperWins)}`,
        `HyperFormula wins in secondary parity artifact: ${String(context.hyperformulaSecondaryCorpus.hyperformulaWins)}`,
      ],
      gaps: [
        ...(context.hyperformulaSecondaryCorpus.artifactPresent ? [] : ['HyperFormula secondary corpus artifact is missing']),
        ...(context.hyperformulaSecondaryCorpus.parseError
          ? [`HyperFormula secondary corpus artifact could not be parsed: ${context.hyperformulaSecondaryCorpus.parseError}`]
          : []),
        ...(context.hyperformulaSecondaryCorpus.suite === 'workpaper-vs-hyperformula'
          ? []
          : [`unexpected HyperFormula secondary corpus suite: ${context.hyperformulaSecondaryCorpus.suite ?? 'missing'}`]),
        ...(context.hyperformulaSecondaryCorpus.comparableCount > 0 ? [] : ['no comparable HyperFormula parity cases recorded']),
        ...(context.hyperformulaSecondaryCorpus.allComparableVerificationEquivalent
          ? []
          : [
              `comparable HyperFormula parity cases missing equivalent verification: ${String(
                context.hyperformulaSecondaryCorpus.comparableVerificationEquivalentCount,
              )}/${String(context.hyperformulaSecondaryCorpus.comparableCount)}`,
            ]),
        ...missingPackageScripts([
          'workpaper:parity:check',
          'workpaper:bench:competitive:check',
          'public-workbook-corpus:completion-audit:check',
        ]).map((script) => `missing package script: ${script}`),
      ],
      evidenceArtifacts: [
        baselineScorecardArtifact,
        hyperFormulaSecondaryCorpusArtifact,
        'packages/headless/src/__tests__/fixtures/hyperformula-surface.json',
      ],
      checkCommands: [
        'pnpm workpaper:parity:check',
        'pnpm workpaper:bench:competitive:check',
        'pnpm public-workbook-corpus:completion-audit:check',
      ],
    }),
]

function buildAuditState(
  status: PublicWorkbookCorpusStatus,
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookCorpusAuditState {
  return {
    targetWorkbookCount: status.targetWorkbookCount,
    sourceCount: status.sourceCount,
    cachedArtifactCount: status.cachedArtifactCount,
    scorecardCaseCount: status.scorecardCaseCount,
    checkpointCaseCount: status.checkpointCaseCount,
    recordedManifestArtifactCount: status.recordedManifestArtifactCount,
    missingCachedArtifactCount: Math.max(0, status.targetWorkbookCount - status.cachedArtifactCount),
    missingVerificationCount: status.missingManifestArtifactCount,
    recordedPassedCaseCount: status.recordedPassedCaseCount,
    recordedUnsupportedCaseCount: status.recordedUnsupportedCaseCount,
    recordedFailedCaseCount: status.recordedFailedCaseCount,
    recordedErrorCaseCount: status.recordedErrorCaseCount,
    recordedFormulaOracleComparisonCount: recordedCases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0),
    recordedFormulaOracleMismatchCount: recordedCases.reduce((sum, entry) => sum + entry.validation.formulaOracleMismatches.length, 0),
    recordedStructuralSmokeRunCount: recordedCases.filter((entry) => entry.validation.structuralSmokePassed !== null).length,
    recordedRoundTripFailureCount: recordedCases.filter((entry) => !entry.validation.roundTripPassed && entry.status !== 'unsupported')
      .length,
  }
}

function readRecordedCases(args: {
  readonly manifest: PublicWorkbookManifest | null
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): PublicWorkbookCorpusCase[] {
  const reusableCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
  if (!args.manifest) {
    return reusableCases
  }
  const casesById = new Map(reusableCases.map((entry) => [entry.id, entry]))
  return args.manifest.artifacts.flatMap((artifact) => {
    const candidate = casesById.get(artifact.id)
    return candidate && publicWorkbookCorpusCaseMatchesArtifact(candidate, artifact) ? [candidate] : []
  })
}

function readHyperFormulaSecondaryCorpus(path: string): PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
  if (!existsSync(path)) {
    return missingHyperFormulaSecondaryCorpus()
  }
  try {
    const record = JSON.parse(readFileSync(path, 'utf8')) as unknown
    if (!isRecord(record)) {
      throw new Error('artifact root is not an object')
    }
    const scorecard = isRecord(record['scorecard']) ? record['scorecard'] : {}
    const results = Array.isArray(record['results']) ? record['results'] : []
    const comparableResults = results.filter((entry) => isRecord(entry) && entry['comparable'] === true)
    const comparableVerificationEquivalentCount = comparableResults.filter((entry) => {
      const comparison = isRecord(entry) ? entry['comparison'] : null
      return isRecord(comparison) && comparison['verificationEquivalent'] === true
    }).length
    const comparableCount = comparableResults.length || readNonNegativeInteger(scorecard, 'comparableCount', 0)
    return {
      artifact: hyperFormulaSecondaryCorpusArtifact,
      artifactPresent: true,
      suite: typeof record['suite'] === 'string' ? record['suite'] : null,
      resultCount: results.length,
      comparableCount,
      workpaperWins: readNonNegativeInteger(scorecard, 'workpaperWins', 0),
      hyperformulaWins: readNonNegativeInteger(scorecard, 'hyperformulaWins', 0),
      comparableVerificationEquivalentCount,
      allComparableVerificationEquivalent: comparableCount > 0 && comparableVerificationEquivalentCount === comparableCount,
      parseError: null,
    }
  } catch (error) {
    return {
      ...missingHyperFormulaSecondaryCorpus(),
      artifactPresent: true,
      parseError: error instanceof Error ? error.message : String(error),
    }
  }
}

function missingHyperFormulaSecondaryCorpus(): PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
  return {
    artifact: hyperFormulaSecondaryCorpusArtifact,
    artifactPresent: false,
    suite: null,
    resultCount: 0,
    comparableCount: 0,
    workpaperWins: 0,
    hyperformulaWins: 0,
    comparableVerificationEquivalentCount: 0,
    allComparableVerificationEquivalent: false,
    parseError: null,
  }
}

function checklistItem(
  item: Omit<PublicWorkbookCorpusAuditChecklistItem, 'evidenceArtifacts' | 'checkCommands'> & {
    readonly evidenceArtifacts?: readonly string[]
    readonly checkCommands?: readonly string[]
  },
): PublicWorkbookCorpusAuditChecklistItem {
  return {
    evidenceArtifacts: [manifestArtifact, baselineScorecardArtifact, checkpointArtifact],
    checkCommands: ['pnpm public-workbook-corpus:status', 'pnpm public-workbook-corpus:completion-audit:check'],
    ...item,
  }
}

function hasFeatureValidationEvidence(entry: PublicWorkbookCorpusCase): boolean {
  const dimensionCellCount = entry.workbookMetadata.dimensions.reduce((sum, dimension) => sum + dimension.nonEmptyCellCount, 0)
  return (
    entry.workbookMetadata.workbookName.trim().length > 0 &&
    entry.workbookMetadata.sheetNames.length === entry.featureCounts.sheetCount &&
    entry.workbookMetadata.dimensions.length === entry.featureCounts.sheetCount &&
    dimensionCellCount === entry.featureCounts.cellCount &&
    entry.workbookMetadata.dimensions.every(
      (dimension) =>
        dimension.sheetName.trim().length > 0 &&
        dimension.rowCount >= 0 &&
        dimension.columnCount >= 0 &&
        dimension.nonEmptyCellCount >= 0 &&
        (dimension.nonEmptyCellCount === 0 || (dimension.rowCount > 0 && dimension.columnCount > 0)),
    ) &&
    entry.featureCounts.cellCount >= entry.featureCounts.formulaCellCount + entry.featureCounts.valueCellCount &&
    entry.featureCounts.definedNameCount >= 0 &&
    entry.featureCounts.tableCount >= 0 &&
    entry.featureCounts.chartCount >= 0 &&
    entry.featureCounts.pivotCount >= 0 &&
    entry.featureCounts.mergeCount >= 0 &&
    entry.featureCounts.styleRangeCount >= 0 &&
    entry.featureCounts.conditionalFormatCount >= 0 &&
    entry.featureCounts.dataValidationCount >= 0 &&
    entry.featureCounts.macroPayloadCount >= 0 &&
    entry.featureCounts.warningCount >= 0
  )
}

function hasUsedRangeValidationEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.workbookMetadata.dimensions.every((dimension) => {
    if (!Object.hasOwn(dimension, 'usedRange')) {
      return false
    }
    const range = dimension.usedRange
    if (dimension.nonEmptyCellCount === 0) {
      return range === null
    }
    return (
      range !== null &&
      range !== undefined &&
      range.startRow >= 0 &&
      range.startColumn >= 0 &&
      range.endRow >= range.startRow &&
      range.endColumn >= range.startColumn &&
      dimension.rowCount === range.endRow + 1 &&
      dimension.columnCount === range.endColumn + 1
    )
  })
}

function hasRecordedProvenanceEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return (
    entry.evidence.includes(`source=${entry.sourceUrl}`) &&
    entry.evidence.includes(`license=${entry.license.title}`) &&
    entry.evidence.includes(`sha256=${entry.sha256}`)
  )
}

function hasCacheIntegrityFailureEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.some(
    (line) => line.startsWith('Missing cached workbook file:') || line.startsWith('Cached workbook hash mismatch:'),
  )
}

function isHashAddressedCachePath(cachePath: string, sha256: string): boolean {
  const parts = cachePath.split(/[\\/]/u)
  return !cachePath.startsWith('/') && !parts.includes('..') && parts.some((part) => part.startsWith(sha256))
}

function buildFeatureWitnessCoverage(cases: readonly PublicWorkbookCorpusCase[]): {
  readonly id: string
  readonly label: string
  readonly totalCount: number
  readonly witnessCaseCount: number
}[] {
  return requiredFeatureWitnesses.map((family) => ({
    id: family.id,
    label: family.label,
    totalCount: cases.reduce((sum, entry) => sum + family.count(entry), 0),
    witnessCaseCount: cases.filter((entry) => family.count(entry) > 0).length,
  }))
}

function hasWorkbookMetadata(entry: PublicWorkbookCorpusCase): boolean {
  return entry.workbookMetadata.sheetNames.length > 0 && entry.workbookMetadata.dimensions.length > 0
}

function isResourceLimitedUnsupportedCase(entry: PublicWorkbookCorpusCase): boolean {
  return (
    entry.status === 'unsupported' &&
    entry.passed &&
    entry.unsupportedFeatureClassifications.some((classification) => classification.startsWith('xlsx.publicCorpus.resourceLimit:')) &&
    entry.evidence.some((line) => line.startsWith('Public corpus verification RSS limit exceeded:'))
  )
}

function countGap(actual: number, required: number, label: string): string[] {
  return actual >= required ? [] : [`${label}: ${String(actual)}/${String(required)}`]
}

function duplicateGap(values: readonly string[], label: string): string[] {
  return new Set(values).size === values.length ? [] : [`${label}: ${String(values.length - new Set(values).size)}`]
}

function readNonNegativeInteger(record: Record<string, unknown>, key: string, fallback: number): number {
  const value = record[key]
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : fallback
}

function packageScripts(): ReadonlyMap<string, string> {
  const parsed = JSON.parse(readFileSync(resolve(rootDir, 'package.json'), 'utf8')) as unknown
  if (!isRecord(parsed) || !isRecord(parsed['scripts'])) {
    throw new Error('package.json is missing scripts')
  }
  const scripts = new Map<string, string>()
  for (const [name, command] of Object.entries(parsed['scripts'])) {
    if (typeof command !== 'string') {
      throw new Error(`package.json script is not a string: ${name}`)
    }
    scripts.set(name, command)
  }
  return scripts
}

function hasPackageScripts(names: readonly string[]): boolean {
  const scripts = packageScripts()
  return names.every((name) => scripts.has(name))
}

function missingPackageScripts(names: readonly string[]): readonly string[] {
  const scripts = packageScripts()
  return names.filter((name) => !scripts.has(name))
}

function isRepoEvidenceArtifact(artifact: string): boolean {
  return (
    artifact.length > 0 &&
    !artifact.startsWith('.cache/') &&
    !artifact.startsWith('http://') &&
    !artifact.startsWith('https://') &&
    !artifact.includes('<') &&
    !artifact.includes('>') &&
    !artifact.startsWith('$') &&
    !artifact.startsWith('/')
  )
}

function pnpmScriptName(command: string): string | null {
  const parts = command
    .trim()
    .split(/\s+/u)
    .filter((part) => !/^[A-Za-z_][A-Za-z0-9_]*=.*/u.test(part))
  const pnpmIndex = parts.indexOf('pnpm')
  if (pnpmIndex < 0) {
    return null
  }
  const candidate = parts[pnpmIndex + 1]
  if (!candidate || candidate === '--' || candidate === 'exec' || candidate === 'dlx') {
    return null
  }
  if (candidate === 'run') {
    return parts[pnpmIndex + 2] ?? null
  }
  if (candidate.startsWith('-')) {
    return null
  }
  return candidate
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
