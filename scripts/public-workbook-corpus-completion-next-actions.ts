import type { PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag } from './public-workbook-corpus-cli.ts'
import type {
  PublicWorkbookCorpusAuditNextAction,
  PublicWorkbookCorpusAuditState,
  PublicWorkbookCorpusNextActionId,
} from './public-workbook-corpus-completion-audit-types.ts'

const resumeFetchBatchSize = 6
const financialFetchTrancheSize = 20
const financialFetchBatchSize = 6
const financialManifestPath = '.cache/public-workbook-corpus-financial/manifest.json'
const financialCacheDir = '.cache/public-workbook-corpus-financial'
const financialScorecardPath = '.cache/public-workbook-corpus-financial/scorecard.json'
const financialVerifyCheckpointPath = '.cache/public-workbook-corpus-financial/verification-checkpoint.json'

export function buildPublicWorkbookCorpusAuditNextActions(args: {
  readonly currentState: PublicWorkbookCorpusAuditState
  readonly status: PublicWorkbookCorpusStatus
  readonly stopMarkerActive: boolean
}): PublicWorkbookCorpusAuditNextAction[] {
  const actions: PublicWorkbookCorpusAuditNextAction[] = []
  if (args.currentState.missingCachedArtifactCount > 0) {
    actions.push(
      nextAction({
        id: 'resume-public-corpus-ingest',
        priority: 1,
        reason: `cached artifacts below target by ${String(args.currentState.missingCachedArtifactCount)}`,
        commands: [
          'pnpm public-workbook-corpus:resume-plan:check',
          ...(args.currentState.fetchTargetReachableFromKnownCandidates
            ? []
            : [`pnpm public-workbook-corpus:discover:plan -- --limit ${String(args.currentState.recommendedDiscoveryLimit)}`]),
          'pnpm public-workbook-corpus:fetch:plan',
        ],
        blockedCommands: args.stopMarkerActive ? resumePublicCorpusIngestBlockedCommands(args.currentState) : [],
      }),
    )
  }
  if (args.currentState.missingVerificationCount > 0) {
    actions.push(
      nextAction({
        id: 'verify-missing-cached-artifacts',
        priority: 3,
        reason: `cached artifacts missing verification evidence: ${String(args.currentState.missingVerificationCount)}`,
        commands: [
          args.status.nextMissingVerificationPlanCommand,
          ...(args.stopMarkerActive ? [] : [args.status.nextMissingVerificationCommand]),
          'pnpm public-workbook-corpus:completion-audit:check',
        ],
        blockedCommands: args.stopMarkerActive ? [args.status.blockedMissingVerificationCommand] : [],
      }),
    )
  }
  if (args.currentState.staleRecordedVerificationCount > 0) {
    actions.push(
      nextAction({
        id: 'refresh-stale-verification-evidence',
        priority: 3,
        reason: `recorded verification cases need evidence refresh: ${String(args.currentState.staleRecordedVerificationCount)}`,
        commands: [
          args.status.nextStaleVerificationPlanCommand,
          ...(args.stopMarkerActive ? [] : [args.status.nextStaleVerificationCommand]),
          'pnpm public-workbook-corpus:completion-audit:check',
        ],
        blockedCommands: args.stopMarkerActive ? [args.status.blockedStaleVerificationCommand] : [],
      }),
    )
  }
  if (args.currentState.missingFeatureWitnessCount > 0) {
    actions.push(
      nextAction({
        id: 'fill-feature-witnesses',
        priority: 3,
        reason: `missing feature witness coverage: ${args.currentState.missingFeatureWitnesses.join(', ')}`,
        commands: ['pnpm public-workbook-corpus:feature-witness:check', 'pnpm public-workbook-corpus:feature-witness:plan'],
      }),
    )
  }
  if (
    args.currentState.financialCachedArtifactCount < args.currentState.financialWorkbookTargetCount ||
    args.currentState.recordedFinancialManifestArtifactCount < args.currentState.financialWorkbookTargetCount ||
    args.currentState.recordedFinancialNonPassingCaseCount > 0
  ) {
    actions.push(
      nextAction({
        id: 'resume-financial-workpapers',
        priority: 2,
        reason: [
          `financial/accounting cached artifacts: ${String(args.currentState.financialCachedArtifactCount)}/${String(
            args.currentState.financialWorkbookTargetCount,
          )}`,
          `recorded cases: ${String(args.currentState.recordedFinancialManifestArtifactCount)}/${String(
            args.currentState.financialWorkbookTargetCount,
          )}`,
          `non-passing cases: ${String(args.currentState.recordedFinancialNonPassingCaseCount)}`,
        ].join('; '),
        commands: [
          'pnpm public-workbook-corpus:discover-financial:check',
          'pnpm public-workbook-corpus:resume-financial:check',
          'pnpm public-workbook-corpus:fetch-financial:plan',
        ],
        blockedCommands: args.stopMarkerActive ? resumeFinancialWorkbookBlockedCommands(args.currentState) : [],
      }),
    )
  }
  if (args.currentState.recordedFailedCaseCount > 0 || args.currentState.recordedErrorCaseCount > 0) {
    actions.push(
      nextAction({
        id: 'inspect-non-passing-scorecard-cases',
        priority: 6,
        reason: `failed/error scorecard cases: ${String(args.currentState.recordedFailedCaseCount)}/${String(
          args.currentState.recordedErrorCaseCount,
        )}`,
        commands: ['pnpm public-workbook-corpus:status', 'pnpm public-workbook-corpus:completion-audit:check'],
      }),
    )
  }
  return actions.toSorted((left, right) => left.priority - right.priority || left.id.localeCompare(right.id))
}

function resumePublicCorpusIngestBlockedCommands(state: PublicWorkbookCorpusAuditState): string[] {
  if (state.missingCachedArtifactCount <= 0) {
    return []
  }
  const batchSize = Math.min(resumeFetchBatchSize, state.missingCachedArtifactCount)
  const commands: string[] = []
  if (!state.fetchTargetReachableFromKnownCandidates) {
    commands.push(blockedCommand(['pnpm', 'public-workbook-corpus:discover', '--', '--limit', String(state.recommendedDiscoveryLimit)]))
  }
  commands.push(
    blockedCommand([
      'pnpm',
      'public-workbook-corpus:fetch',
      '--',
      '--limit',
      String(state.cachedArtifactCount + batchSize),
      '--fetch-batch-size',
      String(batchSize),
    ]),
  )
  return commands
}

function resumeFinancialWorkbookBlockedCommands(state: PublicWorkbookCorpusAuditState): string[] {
  const commands: string[] = []
  const missingFinancialArtifacts = Math.max(0, state.financialWorkbookTargetCount - state.financialCachedArtifactCount)
  if (missingFinancialArtifacts > 0) {
    commands.push(
      blockedCommand([
        'pnpm',
        'public-workbook-corpus:fetch-financial',
        '--',
        '--limit',
        String(Math.min(financialFetchTrancheSize, missingFinancialArtifacts)),
        '--fetch-batch-size',
        String(financialFetchBatchSize),
      ]),
    )
  }
  if (state.financialCachedArtifactCount > state.recordedFinancialManifestArtifactCount) {
    commands.push(
      blockedCommand([
        'pnpm',
        'public-workbook-corpus:verify-missing',
        '--',
        '--manifest',
        financialManifestPath,
        '--scorecard',
        financialScorecardPath,
        '--verify-checkpoint',
        financialVerifyCheckpointPath,
        '--cache-dir',
        financialCacheDir,
        '--limit',
        '1',
      ]),
    )
  }
  return commands
}

function blockedCommand(parts: readonly string[]): string {
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${[...parts, publicCorpusStopMarkerOverrideFlag].map(shellQuote).join(' ')}`
}

function nextAction(action: {
  readonly id: PublicWorkbookCorpusNextActionId
  readonly priority: number
  readonly reason: string
  readonly commands: readonly (string | null)[]
  readonly blockedCommands?: readonly (string | null)[]
}): PublicWorkbookCorpusAuditNextAction {
  return {
    id: action.id,
    priority: action.priority,
    reason: action.reason,
    commands: action.commands.filter((command): command is string => typeof command === 'string' && command.trim().length > 0),
    blockedCommands: (action.blockedCommands ?? []).filter(
      (command): command is string => typeof command === 'string' && command.trim().length > 0,
    ),
  }
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}
