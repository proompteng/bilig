import type { BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'
import { localCiResourceGuardOverrideEnv, type LocalCiResourceGuardStatus } from './ci-local-resource-guard.ts'
import type { UiResponsivenessSameCorpusWorkload } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
} from './ui-responsiveness-same-corpus-workloads.ts'
import type { SameCorpusPublicAccessCheck } from './ui-responsiveness-same-corpus-public-access-check.ts'
import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'

export type UiSameCorpusGoogleSheetsUrlSource = 'argument-or-environment' | 'public-access-check' | 'checked-in-capture' | 'missing'

export interface UiSameCorpusStatus {
  readonly captured: boolean
  readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
  readonly requiredProductCount: number
  readonly requiredCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
  readonly tenXRequirementSatisfied: boolean
  readonly requiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly scrollEventEvidenceCaseCount: number
  readonly casesMissingScrollEventEvidence: readonly string[]
  readonly coveredCorpusCaseIds: readonly string[]
  readonly limitations: readonly string[]
  readonly fixture: UiSameCorpusFixtureStatus
  readonly googleSheetsUrl: string | null
  readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
  readonly googleSheetsUrlEnvVar: string
  readonly microsoftExcelWebEditableUrl: string | null
  readonly microsoftExcelWebEditableUrlEnvVar: string
  readonly publicAccessCheckPath: string
  readonly missingInputs: readonly string[]
  readonly nextFixtureCheckCommand: string
  readonly nextPublicAccessCheckCommand: string
  readonly nextGoogleSheetsStorageStateCommand: string | null
  readonly nextMicrosoftExcelWebStorageStateCommand: string | null
  readonly nextGoogleSheetsUploadInstruction: string | null
  readonly nextMicrosoftExcelWebUploadInstruction: string | null
  readonly nextPreflightCommand: string | null
  readonly nextAuthenticatedPreflightCommand: string | null
  readonly nextCaptureCommand: string | null
  readonly nextAuthenticatedCaptureCommand: string | null
  readonly blockedCommands: readonly string[]
  readonly browserCaptureGuard: UiSameCorpusBrowserCaptureGuardStatus
  readonly nextScorecardGenerateCommand: string | null
  readonly nextDominanceCheckCommand: string
}

export interface UiSameCorpusFixtureStatus {
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly materializedCells: number
  readonly localXlsxPath: string
  readonly publicGithubRawUrl: string
  readonly publicForgejoRawUrl: string
  readonly microsoftExcelWebUrl: string
}

export interface UiSameCorpusBrowserCaptureGuardStatus {
  readonly active: boolean
  readonly activeMarkerPaths: readonly string[]
  readonly overrideEnvVar: string
  readonly overridePrefix: string | null
  readonly nextPreflightRequiresOverride: boolean
  readonly nextCaptureRequiresOverride: boolean
}

export const uiSameCorpusGoogleSheetsUrlEnvVar = 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL'
export const uiSameCorpusMicrosoftExcelWebUrlEnvVar = 'BILIG_UI_SAME_CORPUS_MICROSOFT_EXCEL_WEB_URL'

const defaultUiSameCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
const requiredUiSameCorpusWorkloads = requiredUiResponsivenessSameCorpusWorkloads

export function buildUiSameCorpusStatus(
  input: BuildScorecardInput,
  args: {
    readonly googleSheetsUrl: string | null
    readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
    readonly localCiResourceGuardStatus: LocalCiResourceGuardStatus
    readonly microsoftExcelWebEditableUrl: string | null
    readonly publicAccessCheckPath: string
  },
): UiSameCorpusStatus {
  const proof = input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof
  const fixture = uiSameCorpusFixtureStatus(defaultUiSameCorpusId)
  const coveredWorkloads = new Set(proof.cases.map((entry) => entry.workload))
  const missingRequiredWorkloads = requiredUiSameCorpusWorkloads.filter((workload) => !coveredWorkloads.has(workload))
  const requiredProofCases = proof.cases.filter((entry) => requiredUiSameCorpusWorkloads.includes(entry.workload))
  const scrollEvidenceRequiredProofCases = requiredProofCases.filter((entry) =>
    uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload),
  )
  const casesMissingScrollEventEvidence = scrollEvidenceRequiredProofCases
    .filter((entry) => !uiSameCorpusCaseHasScrollEventEvidence(entry))
    .map((entry) => entry.id)
  const scrollEventEvidenceCaseCount = Math.max(0, scrollEvidenceRequiredProofCases.length - casesMissingScrollEventEvidence.length)
  const tenXRequirementSatisfied = uiSameCorpusTenXRequirementSatisfied(proof, missingRequiredWorkloads, casesMissingScrollEventEvidence)
  const googleSheetsUrlArgument = args.googleSheetsUrl ?? '<google-sheets-url>'
  const microsoftExcelWebUrlArgument = args.microsoftExcelWebEditableUrl ?? '<microsoft-excel-web-editable-url>'
  const browserCaptureGuard = buildBrowserCaptureGuardStatus(args.localCiResourceGuardStatus)
  const missingInputs = args.googleSheetsUrl || tenXRequirementSatisfied ? [] : ['googleSheetsUrlForUploadedSameCorpusWorkbook']
  const nextGoogleSheetsUploadInstruction = missingInputs.includes('googleSheetsUrlForUploadedSameCorpusWorkbook')
    ? `Upload ${fixture.localXlsxPath} to Google Sheets as a native Google Sheet, share it to anyone with the link, then pass its edit URL as --google-sheets-url.`
    : null
  const nextMicrosoftExcelWebUploadInstruction = missingInputs.includes('microsoftExcelWebEditableUrlForUploadedSameCorpusWorkbook')
    ? `Upload ${fixture.localXlsxPath} to OneDrive or Microsoft 365, open it as an editable Excel Web workbook, then pass its browser URL as --microsoft-excel-web-url. The Office viewer URL is only valid for public XLSX identity checks.`
    : null
  const nextGoogleSheetsStorageStateCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--save-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
    '--auth-product',
    'google-sheets',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--corpus',
    fixture.corpusCaseId,
  ]
    .map(shellQuote)
    .join(' ')
  const nextMicrosoftExcelWebStorageStateCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--save-storage-state',
    '.cache/ui-responsiveness/microsoft-excel-web-storage-state.json',
    '--auth-product',
    'microsoft-excel-web',
    '--microsoft-excel-web-url',
    microsoftExcelWebUrlArgument,
    '--corpus',
    fixture.corpusCaseId,
  ]
    .map(shellQuote)
    .join(' ')
  const nextPreflightCommand = ['pnpm', 'ui:same-corpus:capture', '--', '--preflight', '--google-sheets-url', googleSheetsUrlArgument]
    .map(shellQuote)
    .join(' ')
  const nextAuthenticatedPreflightCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--preflight',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--google-sheets-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
  ]
    .map(shellQuote)
    .join(' ')
  const nextCaptureCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--output',
    '.cache/ui-responsiveness/same-corpus-capture.json',
    '--google-sheets-url',
    googleSheetsUrlArgument,
  ]
    .map(shellQuote)
    .join(' ')
  const nextAuthenticatedCaptureCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--output',
    '.cache/ui-responsiveness/same-corpus-capture.json',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--google-sheets-storage-state',
    '.cache/ui-responsiveness/google-sheets-storage-state.json',
  ]
    .map(shellQuote)
    .join(' ')
  const nextScorecardGenerateCommand = 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json'
  return {
    captured: proof.captured,
    evidenceKind: proof.evidenceKind,
    requiredProductCount: proof.requiredProductCount,
    requiredCaseCount: proof.requiredCaseCount,
    tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount,
    tenXRequirementSatisfied,
    requiredWorkloads: requiredUiSameCorpusWorkloads,
    missingRequiredWorkloads,
    scrollEventEvidenceCaseCount,
    casesMissingScrollEventEvidence,
    coveredCorpusCaseIds: proof.coveredCorpusCaseIds,
    limitations: proof.limitations,
    fixture,
    googleSheetsUrl: args.googleSheetsUrl,
    googleSheetsUrlSource: args.googleSheetsUrlSource,
    googleSheetsUrlEnvVar: uiSameCorpusGoogleSheetsUrlEnvVar,
    microsoftExcelWebEditableUrl: args.microsoftExcelWebEditableUrl,
    microsoftExcelWebEditableUrlEnvVar: uiSameCorpusMicrosoftExcelWebUrlEnvVar,
    publicAccessCheckPath: args.publicAccessCheckPath,
    missingInputs,
    nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
    nextPublicAccessCheckCommand: [
      'pnpm',
      'ui:same-corpus:public-check',
      '--',
      '--output',
      args.publicAccessCheckPath,
      '--google-sheets-url',
      googleSheetsUrlArgument,
      '--microsoft-excel-web-url',
      fixture.microsoftExcelWebUrl,
    ]
      .map(shellQuote)
      .join(' '),
    nextGoogleSheetsStorageStateCommand: browserCaptureGuard.active ? null : nextGoogleSheetsStorageStateCommand,
    nextMicrosoftExcelWebStorageStateCommand: browserCaptureGuard.active ? null : nextMicrosoftExcelWebStorageStateCommand,
    nextGoogleSheetsUploadInstruction,
    nextMicrosoftExcelWebUploadInstruction,
    nextPreflightCommand: browserCaptureGuard.active ? null : nextPreflightCommand,
    nextAuthenticatedPreflightCommand: browserCaptureGuard.active ? null : nextAuthenticatedPreflightCommand,
    nextCaptureCommand: browserCaptureGuard.active ? null : nextCaptureCommand,
    nextAuthenticatedCaptureCommand: browserCaptureGuard.active ? null : nextAuthenticatedCaptureCommand,
    blockedCommands: browserCaptureGuard.active
      ? [
          nextGoogleSheetsStorageStateCommand,
          nextMicrosoftExcelWebStorageStateCommand,
          nextPreflightCommand,
          nextAuthenticatedPreflightCommand,
          nextCaptureCommand,
          nextAuthenticatedCaptureCommand,
          nextScorecardGenerateCommand,
        ].map(localCiResourceGuardOverrideCommand)
      : [],
    browserCaptureGuard,
    nextScorecardGenerateCommand: browserCaptureGuard.active ? null : nextScorecardGenerateCommand,
    nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
  }
}

export function resolveUiSameCorpusGoogleSheetsUrl(args: {
  readonly corpusCaseId?: WorkbookBenchmarkCorpusId
  readonly explicitGoogleSheetsUrl: string | null
  readonly publicAccessCheck: SameCorpusPublicAccessCheck | null
  readonly sameCorpusProof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof']
}): {
  readonly googleSheetsUrl: string | null
  readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
} {
  const corpusCaseId = args.corpusCaseId ?? defaultUiSameCorpusId
  if (args.explicitGoogleSheetsUrl) {
    return {
      googleSheetsUrl: args.explicitGoogleSheetsUrl,
      googleSheetsUrlSource: 'argument-or-environment',
    }
  }
  const verifiedPublicAccessUrl = verifiedGoogleSheetsUrlFromPublicAccessCheck(args.publicAccessCheck, corpusCaseId)
  if (verifiedPublicAccessUrl) {
    return {
      googleSheetsUrl: verifiedPublicAccessUrl,
      googleSheetsUrlSource: 'public-access-check',
    }
  }
  const verifiedCheckedInCaptureUrl = verifiedGoogleSheetsUrlFromSameCorpusProof(args.sameCorpusProof, corpusCaseId)
  if (verifiedCheckedInCaptureUrl) {
    return {
      googleSheetsUrl: verifiedCheckedInCaptureUrl,
      googleSheetsUrlSource: 'checked-in-capture',
    }
  }
  return {
    googleSheetsUrl: null,
    googleSheetsUrlSource: 'missing',
  }
}

function buildBrowserCaptureGuardStatus(status: LocalCiResourceGuardStatus): UiSameCorpusBrowserCaptureGuardStatus {
  const active = status.activeMarkerPaths.length > 0
  return {
    active,
    activeMarkerPaths: status.activeMarkerPaths,
    overrideEnvVar: localCiResourceGuardOverrideEnv,
    overridePrefix: active ? `${localCiResourceGuardOverrideEnv}=1` : null,
    nextPreflightRequiresOverride: active,
    nextCaptureRequiresOverride: active,
  }
}

function verifiedGoogleSheetsUrlFromPublicAccessCheck(
  check: SameCorpusPublicAccessCheck | null,
  corpusCaseId: WorkbookBenchmarkCorpusId,
): string | null {
  if (!check || check.corpusCaseId !== corpusCaseId) {
    return null
  }
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  if (check.materializedCells !== corpus.materializedCellCount) {
    return null
  }
  const product = check.products.find((entry) => entry.product === 'google-sheets')
  return product?.corpusVerification.verified ? product.source : null
}

function verifiedGoogleSheetsUrlFromSameCorpusProof(
  proof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof'],
  corpusCaseId: WorkbookBenchmarkCorpusId,
): string | null {
  if (!proof.captured || proof.evidenceKind !== 'same-corpus-browser-capture' || !proof.coveredCorpusCaseIds.includes(corpusCaseId)) {
    return null
  }
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  const corpusCases = proof.cases.filter((entry) => entry.corpusCaseId === corpusCaseId)
  if (corpusCases.length === 0) {
    return null
  }
  let url: string | null = null
  for (const entry of corpusCases) {
    const googleSheets = entry.googleSheets
    const source = googleSheets.source.trim()
    if (
      source.length === 0 ||
      !googleSheets.corpusVerification.verified ||
      googleSheets.corpusVerification.method !== 'google-sheets-xlsx-export' ||
      googleSheets.corpusVerification.materializedCells !== corpus.materializedCellCount
    ) {
      return null
    }
    if (url !== null && url !== source) {
      return null
    }
    url = source
  }
  return url
}

function uiSameCorpusTenXRequirementSatisfied(
  proof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof'],
  missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[],
  casesMissingScrollEventEvidence: readonly string[],
): boolean {
  return (
    proof.captured &&
    proof.evidenceKind === 'same-corpus-browser-capture' &&
    proof.requiredProductCount === 2 &&
    proof.requiredCaseCount > 0 &&
    proof.cases.length === proof.requiredCaseCount &&
    proof.tenXMeanAndP95CaseCount === proof.requiredCaseCount &&
    missingRequiredWorkloads.length === 0 &&
    casesMissingScrollEventEvidence.length === 0 &&
    proof.cases.every((entry) => entry.tenXMeanAndP95AgainstGoogleSheets && entry.passed)
  )
}

function uiSameCorpusCaseHasScrollEventEvidence(
  entry: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof']['cases'][number],
): boolean {
  return (
    entry.tenXMeanAndP95Metric === 'scrollEventResponseMs' &&
    Boolean(entry.bilig.scrollEventResponseMs) &&
    Boolean(entry.googleSheets.scrollEventResponseMs) &&
    Boolean(entry.bilig.scrollMovementPx) &&
    Boolean(entry.googleSheets.scrollMovementPx)
  )
}

function uiSameCorpusFixtureStatus(corpusCaseId: WorkbookBenchmarkCorpusId): UiSameCorpusFixtureStatus {
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  const localXlsxPath = `packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`
  const publicGithubRawUrl = `https://raw.githubusercontent.com/proompteng/bilig/main/${localXlsxPath}`
  return {
    corpusCaseId,
    materializedCells: corpus.materializedCellCount,
    localXlsxPath,
    publicGithubRawUrl,
    publicForgejoRawUrl: `https://code.proompteng.ai/kalmyk/bilig/raw/branch/main/${localXlsxPath}`,
    microsoftExcelWebUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicGithubRawUrl)}`,
  }
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function localCiResourceGuardOverrideCommand(command: string): string {
  if (command.includes(`${localCiResourceGuardOverrideEnv}=1`)) {
    return command
  }
  return `${localCiResourceGuardOverrideEnv}=1 ${command}`
}
