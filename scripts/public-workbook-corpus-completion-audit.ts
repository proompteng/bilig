#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { hasUsableLicenseEvidence, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import { readPublicWorkbookCorpusStatus, type PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import { readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
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

function main(): void {
  const audit = buildPublicWorkbookCorpusCompletionAuditFromArgs()
  const requireComplete = readFlagArg('--require-complete')
  if (readFlagArg('--check')) {
    const findings = validatePublicWorkbookCorpusCompletionAudit(audit, { requireComplete })
    if (findings.length > 0) {
      throw new Error(`Public workbook corpus completion audit is invalid: ${findings.join('; ')}`)
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
    const metadataCaseGapCount = context.recordedCases.filter(
      (entry) => entry.workbookMetadata.sheetNames.length === 0 || entry.workbookMetadata.dimensions.length === 0,
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
        metadataCaseGapCount === 0 &&
        context.status.recordedCoversManifest &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `manifest present: ${String(Boolean(context.manifest))}`,
        `manifest sources with license evidence: ${String(sources.length - sourceLicenseGapCount)}/${String(sources.length)}`,
        `artifacts with license evidence: ${String(artifacts.length - artifactLicenseGapCount)}/${String(artifacts.length)}`,
        `recorded workbook metadata cases: ${String(context.recordedCases.length - metadataCaseGapCount)}/${String(
          context.currentState.cachedArtifactCount,
        )}`,
      ],
      gaps: [
        ...(context.manifest ? [] : ['manifest artifact is missing']),
        ...(sourceLicenseGapCount === 0 ? [] : [`sources missing usable license evidence: ${String(sourceLicenseGapCount)}`]),
        ...(artifactLicenseGapCount === 0 ? [] : [`artifacts missing usable license evidence: ${String(artifactLicenseGapCount)}`]),
        ...(artifactMetadataGapCount === 0 ? [] : [`artifacts missing source/hash/fetch metadata: ${String(artifactMetadataGapCount)}`]),
        ...(metadataCaseGapCount === 0 ? [] : [`recorded cases missing workbook metadata: ${String(metadataCaseGapCount)}`]),
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
  (context) =>
    checklistItem({
      id: 'validate-workbook-features',
      priority: 3,
      promptRequirement:
        'Validate sheets, dimensions, used ranges, formulas, values, names, tables, charts, pivots, styles, merged ranges, conditional formats, and unsupported feature classifications.',
      passed:
        context.status.recordedCoversManifest &&
        context.recordedCases.every(hasFeatureValidationEvidence) &&
        context.currentState.cachedArtifactCount >= context.currentState.targetWorkbookCount,
      evidence: [
        `feature-count fields validated by recorded cases: ${String(context.recordedCases.filter(hasFeatureValidationEvidence).length)}/${String(
          context.currentState.cachedArtifactCount,
        )}`,
        `recorded unsupported cases: ${String(context.currentState.recordedUnsupportedCaseCount)}`,
      ],
      gaps: [
        ...(context.status.recordedCoversManifest
          ? []
          : [`cached artifacts missing feature validation evidence: ${String(context.currentState.missingVerificationCount)}`]),
        ...countGap(
          context.currentState.cachedArtifactCount,
          context.currentState.targetWorkbookCount,
          'feature validation does not yet cover target artifacts',
        ),
      ],
    }),
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
        context.currentState.scorecardCaseCount >= context.currentState.targetWorkbookCount &&
        context.status.targetComplete,
      evidence: [
        `scorecard cases: ${String(context.currentState.scorecardCaseCount)}/${String(context.currentState.targetWorkbookCount)}`,
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
        ...countGap(context.currentState.scorecardCaseCount, context.currentState.targetWorkbookCount, 'scorecard cases below target'),
        ...(context.status.targetComplete ? [] : context.status.gaps),
      ],
    }),
  (_context) =>
    checklistItem({
      id: 'ci-offline-cached-mode',
      priority: 6,
      promptRequirement: 'Add CI-friendly focused gates plus an offline cached corpus mode.',
      passed: hasPackageScripts([
        'public-workbook-corpus:check',
        'public-workbook-corpus:resume-plan:check',
        'public-workbook-corpus:completion-audit:check',
        'test:correctness:corpus',
      ]),
      evidence: [
        'public-workbook-corpus:check validates cached scorecard/manifest evidence',
        'public-workbook-corpus:resume-plan:check validates bounded resume commands without running workers',
        'public-workbook-corpus:completion-audit:check maps the objective to current artifacts',
        'test:correctness:corpus covers the offline corpus harness and XLSX import/export regressions',
      ],
      gaps: missingPackageScripts([
        'public-workbook-corpus:check',
        'public-workbook-corpus:resume-plan:check',
        'public-workbook-corpus:completion-audit:check',
        'test:correctness:corpus',
      ]).map((script) => `missing package script: ${script}`),
    }),
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
  return (
    entry.workbookMetadata.sheetNames.length === entry.featureCounts.sheetCount &&
    entry.workbookMetadata.dimensions.length === entry.featureCounts.sheetCount &&
    entry.featureCounts.cellCount >= entry.featureCounts.formulaCellCount + entry.featureCounts.valueCellCount &&
    entry.featureCounts.definedNameCount >= 0 &&
    entry.featureCounts.tableCount >= 0 &&
    entry.featureCounts.chartCount >= 0 &&
    entry.featureCounts.pivotCount >= 0 &&
    entry.featureCounts.mergeCount >= 0 &&
    entry.featureCounts.styleRangeCount >= 0 &&
    entry.featureCounts.conditionalFormatCount >= 0
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

function packageScripts(): ReadonlySet<string> {
  const parsed = JSON.parse(readFileSync(resolve(rootDir, 'package.json'), 'utf8')) as unknown
  if (!isRecord(parsed) || !isRecord(parsed['scripts'])) {
    throw new Error('package.json is missing scripts')
  }
  return new Set(Object.keys(parsed['scripts']))
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
