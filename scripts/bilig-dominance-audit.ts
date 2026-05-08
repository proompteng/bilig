#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import type { BiligDominanceScorecard, DominanceCategory, DominanceCompletionCriterion } from './bilig-dominance-scorecard-types.ts'
import { buildBiligDominanceStatusFromArgs, type BiligDominanceStatus } from './bilig-dominance-status.ts'
import { loadBiligDominanceScorecardInput } from './bilig-dominance-scorecard-input.ts'
import { buildBiligDominanceScorecard } from './gen-bilig-dominance-scorecard.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)

export interface BiligDominancePromptArtifactAudit {
  readonly schemaVersion: 1
  readonly objective: string
  readonly completionVerdict: {
    readonly goalStatus: BiligDominanceStatus['goalStatus']
    readonly blanketTenXClaimAllowed: boolean
    readonly allChecklistItemsPassed: boolean
    readonly unmetRequirements: readonly string[]
  }
  readonly checklist: readonly BiligDominancePromptArtifactChecklistItem[]
  readonly liveLocalCiResourceGuard: BiligDominanceStatus['localCiResourceGuard']
  readonly livePublicWorkbookCorpus: BiligDominanceStatus['publicWorkbookCorpus']
  readonly liveUiSameCorpus: BiligDominanceStatus['uiSameCorpus']
}

export interface BiligDominancePromptArtifactChecklistItem {
  readonly id: string
  readonly objectiveCategory: string
  readonly promptRequirement: string
  readonly passed: boolean
  readonly evidence: readonly string[]
  readonly evidenceArtifacts: readonly string[]
  readonly checkCommands: readonly string[]
  readonly gaps: readonly string[]
  readonly liveBlockers: readonly string[]
}

function main(): void {
  const status = buildBiligDominanceStatusFromArgs()
  const scorecard = buildBiligDominanceScorecard(loadBiligDominanceScorecardInput())
  const audit = buildBiligDominancePromptArtifactAudit({ scorecard, status })
  if (process.argv.includes('--check')) {
    const findings = validateBiligDominancePromptArtifactAudit(audit)
    if (findings.length > 0) {
      throw new Error(`Bilig dominance prompt-to-artifact audit is invalid: ${findings.join('; ')}`)
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          checklistItemCount: audit.checklist.length,
          goalStatus: audit.completionVerdict.goalStatus,
          blanketTenXClaimAllowed: audit.completionVerdict.blanketTenXClaimAllowed,
          allChecklistItemsPassed: audit.completionVerdict.allChecklistItemsPassed,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  process.stdout.write(`${JSON.stringify(audit, null, 2)}\n`)
}

export function buildBiligDominancePromptArtifactAudit(args: {
  readonly scorecard: BiligDominanceScorecard
  readonly status: BiligDominanceStatus
}): BiligDominancePromptArtifactAudit {
  const checklist = args.scorecard.completionAudit.criteria.map((criterion) =>
    buildChecklistItem({
      category: requiredCategory(args.scorecard.categories, criterion.id),
      criterion,
      liveBlockers: liveBlockersForCriterion(args.status, criterion.id),
      liveLocalCiResourceGuard: args.status.localCiResourceGuard,
      livePublicWorkbookCorpus: args.status.publicWorkbookCorpus,
      liveUiSameCorpus: args.status.uiSameCorpus,
    }),
  )
  return {
    schemaVersion: 1,
    objective: args.scorecard.objective,
    completionVerdict: {
      goalStatus: args.status.goalStatus,
      blanketTenXClaimAllowed: args.status.blanketTenXClaimAllowed,
      allChecklistItemsPassed: checklist.every((entry) => entry.passed),
      unmetRequirements: args.status.unmetRequirements,
    },
    checklist,
    liveLocalCiResourceGuard: args.status.localCiResourceGuard,
    livePublicWorkbookCorpus: args.status.publicWorkbookCorpus,
    liveUiSameCorpus: args.status.uiSameCorpus,
  }
}

export function validateBiligDominancePromptArtifactAudit(audit: BiligDominancePromptArtifactAudit): string[] {
  const findings: string[] = []
  const checklistIds = audit.checklist.map((entry) => entry.id)
  if (JSON.stringify(checklistIds) !== JSON.stringify(requiredChecklistIds)) {
    findings.push(`checklist ids do not match required objective categories: ${checklistIds.join(', ')}`)
  }
  const duplicateIds = checklistIds.filter((id, index) => checklistIds.indexOf(id) !== index)
  if (duplicateIds.length > 0) {
    findings.push(`duplicate checklist ids: ${[...new Set(duplicateIds)].join(', ')}`)
  }
  for (const item of audit.checklist) {
    if (!item.promptRequirement.trim()) {
      findings.push(`${item.id} is missing a prompt requirement`)
    }
    if (item.evidence.length === 0) {
      findings.push(`${item.id} has no evidence entries`)
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
    if (item.passed && (item.gaps.length > 0 || item.liveBlockers.length > 0)) {
      findings.push(`${item.id} is passed while still reporting gaps or live blockers`)
    }
    if (!item.passed && item.gaps.length === 0) {
      findings.push(`${item.id} is failed without an explicit gap`)
    }
  }
  const allChecklistItemsPassed = audit.checklist.every((entry) => entry.passed)
  if (audit.completionVerdict.allChecklistItemsPassed !== allChecklistItemsPassed) {
    findings.push('completion verdict allChecklistItemsPassed does not match checklist state')
  }
  if (audit.completionVerdict.blanketTenXClaimAllowed && !allChecklistItemsPassed) {
    findings.push('blanket 10x claim is allowed before every checklist item passed')
  }
  if (audit.completionVerdict.goalStatus === 'achieved' && (!audit.completionVerdict.blanketTenXClaimAllowed || !allChecklistItemsPassed)) {
    findings.push('goal is achieved without blanket claim permission and complete checklist evidence')
  }
  const importExportItem = audit.checklist.find((entry) => entry.id === 'import-export-compatibility')
  const uiResponsivenessItem = audit.checklist.find((entry) => entry.id === 'ui-responsiveness')
  const operatorWorkflowItem = audit.checklist.find((entry) => entry.id === 'operator-developer-workflow')
  if (!audit.livePublicWorkbookCorpus.targetComplete) {
    if (!importExportItem) {
      findings.push('live public workbook corpus is incomplete but import/export checklist item is missing')
    } else {
      if (importExportItem.passed) {
        findings.push('import/export checklist passed while live public workbook corpus target is incomplete')
      }
      if (importExportItem.liveBlockers.length === 0) {
        findings.push('live public workbook corpus target is incomplete but import/export live blockers are empty')
      }
    }
  }
  if (!audit.liveUiSameCorpus.tenXRequirementSatisfied) {
    if (!uiResponsivenessItem) {
      findings.push('live same-corpus UI proof is incomplete but UI responsiveness checklist item is missing')
    } else {
      if (uiResponsivenessItem.passed) {
        findings.push('UI responsiveness checklist passed while live same-corpus UI proof is incomplete')
      }
      if (uiResponsivenessItem.liveBlockers.length === 0) {
        findings.push('live same-corpus UI proof is incomplete but UI responsiveness live blockers are empty')
      }
    }
  }
  if (audit.liveLocalCiResourceGuard.active) {
    if (!operatorWorkflowItem) {
      findings.push('local CI resource guard is active but operator/developer workflow checklist item is missing')
    } else {
      if (operatorWorkflowItem.passed) {
        findings.push('operator/developer workflow checklist passed while local CI resource guard is active')
      }
      if (operatorWorkflowItem.liveBlockers.length === 0) {
        findings.push('local CI resource guard is active but operator/developer workflow live blockers are empty')
      }
    }
  }
  if (audit.livePublicWorkbookCorpus.nextCorpusRunRequiresExplicitResume && !audit.livePublicWorkbookCorpus.corpusRunStopMarkerActive) {
    findings.push('corpus resume is required without an active stop marker')
  }
  return findings
}

const requiredChecklistIds = [
  'calculation-correctness',
  'recalculation-speed',
  'structural-edit-performance',
  'large-workbook-scale',
  'ui-responsiveness',
  'collaboration',
  'automation-api-extensibility',
  'import-export-compatibility',
  'auditability',
  'reliability',
  'security',
  'operator-developer-workflow',
] as const

function buildChecklistItem(args: {
  readonly category: DominanceCategory
  readonly criterion: DominanceCompletionCriterion
  readonly liveBlockers: readonly string[]
  readonly liveLocalCiResourceGuard: BiligDominanceStatus['localCiResourceGuard']
  readonly livePublicWorkbookCorpus: BiligDominanceStatus['publicWorkbookCorpus']
  readonly liveUiSameCorpus: BiligDominanceStatus['uiSameCorpus']
}): BiligDominancePromptArtifactChecklistItem {
  const gaps = uniqueStrings([...args.criterion.gaps, ...args.liveBlockers])
  return {
    id: args.criterion.id,
    objectiveCategory: args.category.objectiveCategory,
    promptRequirement: args.criterion.requirement,
    passed: args.criterion.passed && args.liveBlockers.length === 0,
    evidence:
      args.criterion.id === 'import-export-compatibility'
        ? [
            ...args.criterion.evidence,
            `live public workbook corpus cached artifacts: ${String(args.livePublicWorkbookCorpus.cachedArtifactCount)}/${String(
              args.livePublicWorkbookCorpus.targetWorkbookCount,
            )}`,
            `live public workbook corpus recorded verification cases: ${String(
              args.livePublicWorkbookCorpus.recordedManifestArtifactCount,
            )}/${String(args.livePublicWorkbookCorpus.cachedArtifactCount)}`,
            `live public workbook corpus scorecard cases: ${String(args.livePublicWorkbookCorpus.scorecardCaseCount)}/${String(
              args.livePublicWorkbookCorpus.cachedArtifactCount,
            )}`,
            `live public workbook corpus recorded all cases passed: ${String(args.livePublicWorkbookCorpus.recordedAllCasesPassed)}`,
            ...args.category.currentEvidence.filter((entry) => entry.startsWith('public workbook corpus next')),
          ]
        : args.criterion.id === 'ui-responsiveness'
          ? [
              ...args.criterion.evidence,
              `live same-corpus UI proof captured: ${String(args.liveUiSameCorpus.captured)}`,
              `live same-corpus UI 10x cases: ${String(args.liveUiSameCorpus.tenXMeanAndP95CaseCount)}/${String(
                args.liveUiSameCorpus.requiredCaseCount,
              )}`,
              `live same-corpus UI required workloads: ${args.liveUiSameCorpus.requiredWorkloads.join(', ') || 'none'}`,
              `live same-corpus UI missing required workloads: ${args.liveUiSameCorpus.missingRequiredWorkloads.join(', ') || 'none'}`,
              `live same-corpus UI scroll-event evidence cases: ${String(args.liveUiSameCorpus.scrollEventEvidenceCaseCount)}/${String(
                args.liveUiSameCorpus.requiredCaseCount,
              )}`,
              `live same-corpus UI cases missing scroll-event evidence: ${
                args.liveUiSameCorpus.casesMissingScrollEventEvidence.join(', ') || 'none'
              }`,
              `live same-corpus UI missing inputs: ${args.liveUiSameCorpus.missingInputs.join(', ') || 'none'}`,
              `live same-corpus UI Google Sheets URL source: ${args.liveUiSameCorpus.googleSheetsUrlSource}`,
              `live same-corpus UI browser capture guard active: ${String(args.liveUiSameCorpus.browserCaptureGuard.active)}`,
            ]
          : args.criterion.id === 'operator-developer-workflow'
            ? [
                ...args.criterion.evidence,
                `local CI resource guard active: ${String(args.liveLocalCiResourceGuard.active)}`,
                `local CI resource guard markers: ${args.liveLocalCiResourceGuard.activeMarkerPaths.join(', ') || 'none'}`,
                `local CI resource guard override env: ${args.liveLocalCiResourceGuard.overrideEnvVar}`,
              ]
            : args.criterion.evidence,
    evidenceArtifacts:
      args.criterion.id === 'import-export-compatibility'
        ? [...args.category.evidenceArtifacts, 'packages/benchmarks/baselines/public-workbook-corpus-scorecard.json']
        : args.category.evidenceArtifacts,
    checkCommands:
      args.criterion.id === 'import-export-compatibility'
        ? [...args.category.checkCommands, 'pnpm public-workbook-corpus:check']
        : args.category.checkCommands,
    gaps,
    liveBlockers: args.liveBlockers,
  }
}

function liveBlockersForCriterion(status: BiligDominanceStatus, criterionId: string): readonly string[] {
  if (criterionId === 'import-export-compatibility') {
    return status.importExportBlockers
  }
  if (criterionId === 'ui-responsiveness') {
    return uiSameCorpusLiveBlockers(status.uiSameCorpus)
  }
  if (criterionId === 'operator-developer-workflow') {
    return localCiResourceGuardLiveBlockers(status.localCiResourceGuard)
  }
  return []
}

function uiSameCorpusLiveBlockers(status: BiligDominanceStatus['uiSameCorpus']): readonly string[] {
  const blockers: string[] = []
  if (!status.captured) {
    blockers.push('same-corpus UI browser capture has not been recorded')
  }
  if (status.missingRequiredWorkloads.length > 0) {
    blockers.push(`same-corpus UI proof missing required workloads: ${status.missingRequiredWorkloads.join(', ')}`)
  }
  if (status.casesMissingScrollEventEvidence.length > 0) {
    blockers.push(`same-corpus UI proof missing scroll-event evidence: ${status.casesMissingScrollEventEvidence.join(', ')}`)
  }
  if (status.captured && !status.tenXRequirementSatisfied) {
    blockers.push(
      `same-corpus UI proof has ${String(status.tenXMeanAndP95CaseCount)}/${String(status.requiredCaseCount)} required 10x cases`,
    )
  }
  if (status.missingInputs.length > 0) {
    blockers.push(`same-corpus UI proof missing inputs: ${status.missingInputs.join(', ')}`)
  }
  if (status.browserCaptureGuard.active) {
    blockers.push(
      `same-corpus UI browser capture paused by local resource guard: ${status.browserCaptureGuard.activeMarkerPaths.join(', ')}`,
    )
  }
  return blockers
}

function localCiResourceGuardLiveBlockers(status: BiligDominanceStatus['localCiResourceGuard']): readonly string[] {
  return status.active ? [`local CI resource guard active: ${status.activeMarkerPaths.join(', ')}`] : []
}

function uniqueStrings(values: readonly string[]): readonly string[] {
  return [...new Set(values)]
}

function requiredCategory(categories: readonly DominanceCategory[], id: string): DominanceCategory {
  const category = categories.find((candidate) => candidate.id === id)
  if (!category) {
    throw new Error(`Dominance scorecard is missing category for completion criterion: ${id}`)
  }
  return category
}

let packageScriptCache: ReadonlySet<string> | null = null

function packageScripts(): ReadonlySet<string> {
  if (packageScriptCache) {
    return packageScriptCache
  }
  const parsed = JSON.parse(readFileSync(resolve(rootDir, 'package.json'), 'utf8')) as unknown
  if (!isRecord(parsed) || !isRecord(parsed['scripts'])) {
    throw new Error('package.json is missing scripts')
  }
  packageScriptCache = new Set(Object.keys(parsed['scripts']))
  return packageScriptCache
}

function isRepoEvidenceArtifact(artifact: string): boolean {
  return (
    artifact.length > 0 &&
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
