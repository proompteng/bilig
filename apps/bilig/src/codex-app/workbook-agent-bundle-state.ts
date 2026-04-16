import type { WorkbookAgentTimelineCitation } from '@bilig/contracts'
import type { WorkbookAgentCommandBundle, WorkbookAgentSharedReviewState } from '@bilig/agent-api'
import { requiresWorkbookAgentOwnerReview } from '@bilig/agent-api'

interface WorkbookAgentBundleSessionStateRef {
  readonly scope: 'private' | 'shared'
  readonly storageActorUserId: string
}

export function createBundleRangeCitations(bundle: Pick<WorkbookAgentCommandBundle, 'affectedRanges'>): WorkbookAgentTimelineCitation[] {
  return bundle.affectedRanges.map((range) => ({
    kind: 'range',
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
    role: range.role,
  }))
}

export function appendRevisionCitation(
  citations: readonly WorkbookAgentTimelineCitation[],
  revision: number,
): WorkbookAgentTimelineCitation[] {
  return [...citations, { kind: 'revision', revision }]
}

export function createWorkflowTurnId(runId: string): string {
  return `workflow:${runId}`
}

export function needsSharedOwnerReview(
  sessionState: WorkbookAgentBundleSessionStateRef,
  bundle: Pick<WorkbookAgentCommandBundle, 'riskClass'>,
): boolean {
  return requiresWorkbookAgentOwnerReview({
    scope: sessionState.scope,
    riskClass: bundle.riskClass,
  })
}

export function createPendingSharedReviewState(ownerUserId: string): WorkbookAgentSharedReviewState {
  return {
    ownerUserId,
    status: 'pending',
    decidedByUserId: null,
    decidedAtUnixMs: null,
    recommendations: [],
  }
}

export function normalizeSharedReviewState(
  bundle: WorkbookAgentCommandBundle,
  sessionState: WorkbookAgentBundleSessionStateRef,
): WorkbookAgentSharedReviewState | null {
  if (!needsSharedOwnerReview(sessionState, bundle)) {
    return null
  }
  if (bundle.sharedReview && bundle.sharedReview.ownerUserId === sessionState.storageActorUserId) {
    return {
      ...bundle.sharedReview,
      recommendations: [...(bundle.sharedReview.recommendations ?? [])],
    }
  }
  return createPendingSharedReviewState(sessionState.storageActorUserId)
}

export function attachSharedReviewState(
  bundle: WorkbookAgentCommandBundle,
  sessionState: WorkbookAgentBundleSessionStateRef,
): WorkbookAgentCommandBundle {
  return {
    ...bundle,
    sharedReview: normalizeSharedReviewState(bundle, sessionState),
  }
}
