import type { WorkbookAgentAppliedBy, WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from '@bilig/agent-api'
import type { ZeroSyncService } from '../zero/service.js'
import type { WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import type { WorkbookAgentSessionRegistry } from './workbook-agent-session-registry.js'
import type { WorkbookAgentBundleApplicationContext } from './workbook-agent-service-application.js'
import type { WorkbookAgentReviewActionContext } from './workbook-agent-service-review-actions.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function createWorkbookAgentBundleApplicationContext(input: {
  readonly zeroSyncService: ZeroSyncService
  readonly now: () => number
  readonly featureFlags: Pick<WorkbookAgentFeatureFlags, 'autoApplyLowRiskEnabled'>
  readonly sessionRegistry: WorkbookAgentSessionRegistry
  readonly isRolloutAllowed: (documentId: string, userId: string) => boolean
}): WorkbookAgentBundleApplicationContext {
  return {
    zeroSyncService: input.zeroSyncService,
    now: input.now,
    autoApplyLowRiskEnabled: input.featureFlags.autoApplyLowRiskEnabled,
    isRolloutAllowed: input.isRolloutAllowed,
    touchSession: (sessionState) => input.sessionRegistry.touch(sessionState),
  }
}

export function createWorkbookAgentReviewActionContext(input: {
  readonly zeroSyncService: ZeroSyncService
  readonly now: () => number
  readonly sessionRegistry: WorkbookAgentSessionRegistry
  readonly applyCommandBundleForSessionState: (input: {
    readonly sessionState: WorkbookAgentThreadState
    readonly commandBundle: WorkbookAgentCommandBundle
    readonly actorUserId: string
    readonly appliedBy: WorkbookAgentAppliedBy
    readonly commandIndexes?: readonly number[] | null | undefined
  }) => Promise<WorkbookAgentExecutionRecord>
  readonly shouldApplyToolBundleImmediately: (sessionState: WorkbookAgentThreadState, bundle: WorkbookAgentCommandBundle) => boolean
  readonly persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
}): WorkbookAgentReviewActionContext {
  return {
    now: input.now,
    getWorkbookHeadRevision: async (documentId) => await input.zeroSyncService.getWorkbookHeadRevision(documentId),
    applyCommandBundleForSessionState: input.applyCommandBundleForSessionState,
    shouldApplyToolBundleImmediately: input.shouldApplyToolBundleImmediately,
    persistSessionState: input.persistSessionState,
    emitSnapshot: (threadId) => input.sessionRegistry.emitSnapshot(threadId),
    touchSession: (sessionState) => input.sessionRegistry.touch(sessionState),
  }
}
