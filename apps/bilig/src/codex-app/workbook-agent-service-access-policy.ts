import type { WorkbookAgentThreadSummary } from '@bilig/contracts'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { isWorkbookAgentRolloutAllowed, type WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function assertWorkbookAgentSharedThreadAccess(input: {
  readonly featureFlags: Pick<WorkbookAgentFeatureFlags, 'sharedThreadsEnabled' | 'allowlistedUserIds' | 'allowlistedDocumentIds'>
  readonly documentId: string
  readonly userId: string
  readonly disabledCode?: string
  readonly rolloutBlockedCode?: string
}): void {
  if (!input.featureFlags.sharedThreadsEnabled) {
    throw createWorkbookAgentServiceError({
      code: input.disabledCode ?? 'WORKBOOK_AGENT_SHARED_THREADS_DISABLED',
      message: 'Shared workbook assistant threads are currently disabled.',
      statusCode: 409,
      retryable: false,
    })
  }
  if (
    !isWorkbookAgentRolloutAllowed(input.featureFlags, {
      documentId: input.documentId,
      userId: input.userId,
    })
  ) {
    throw createWorkbookAgentServiceError({
      code: input.rolloutBlockedCode ?? 'WORKBOOK_AGENT_SHARED_THREADS_ROLLOUT_BLOCKED',
      message: 'Shared workbook assistant threads are still limited to the rollout allowlist.',
      statusCode: 409,
      retryable: false,
    })
  }
}

export function assertWorkbookAgentSessionAccessPolicy(input: {
  readonly featureFlags: Pick<WorkbookAgentFeatureFlags, 'sharedThreadsEnabled' | 'allowlistedUserIds' | 'allowlistedDocumentIds'>
  readonly sessionState: Pick<WorkbookAgentThreadState, 'scope'>
  readonly documentId: string
  readonly userId: string
}): void {
  if (input.sessionState.scope !== 'shared') {
    return
  }
  assertWorkbookAgentSharedThreadAccess({
    featureFlags: input.featureFlags,
    documentId: input.documentId,
    userId: input.userId,
  })
}

export function filterWorkbookAgentThreadSummariesByAccessPolicy(input: {
  readonly featureFlags: Pick<WorkbookAgentFeatureFlags, 'sharedThreadsEnabled' | 'allowlistedUserIds' | 'allowlistedDocumentIds'>
  readonly documentId: string
  readonly summaries: readonly WorkbookAgentThreadSummary[]
  readonly userId: string
}): WorkbookAgentThreadSummary[] {
  return input.summaries.filter((summary) => {
    if (summary.scope !== 'shared') {
      return true
    }
    return (
      input.featureFlags.sharedThreadsEnabled &&
      isWorkbookAgentRolloutAllowed(input.featureFlags, {
        documentId: input.documentId,
        userId: input.userId,
      })
    )
  })
}
