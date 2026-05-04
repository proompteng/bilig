import type { WorkbookAgentCommandBundle, WorkbookAgentReviewQueueItem, WorkbookAgentSharedReviewState } from '@bilig/agent-api'
import { describeWorkbookAgentBundle, toWorkbookAgentCommandBundle, toWorkbookAgentReviewQueueItem } from '@bilig/agent-api'
import {
  createPendingSharedReviewState,
  createBundleRangeCitations,
  needsSharedOwnerReview,
  normalizeSharedReviewState,
} from './workbook-agent-bundle-state.js'
import { createSystemEntry } from './workbook-agent-session-model.js'
import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import { upsertEntry, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function getCurrentWorkbookAgentReviewItem(sessionState: WorkbookAgentThreadState): WorkbookAgentReviewQueueItem | null {
  return sessionState.durable.reviewQueueItems[0] ?? null
}

export function replaceCurrentWorkbookAgentReviewItem(
  sessionState: WorkbookAgentThreadState,
  reviewItem: WorkbookAgentReviewQueueItem | null,
): void {
  sessionState.durable.reviewQueueItems = reviewItem ? [reviewItem] : []
}

export function createWorkbookAgentReviewQueueItem(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly bundle: WorkbookAgentCommandBundle
}): WorkbookAgentReviewQueueItem {
  return toWorkbookAgentReviewQueueItem({
    bundle: input.bundle,
    reviewMode: needsSharedOwnerReview(input.sessionState, input.bundle) ? 'ownerReview' : 'manual',
    sharedReview: input.bundle.sharedReview ?? null,
  })
}

export function stageWorkbookAgentReviewBundle(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly turnId: string
  readonly bundle: WorkbookAgentCommandBundle
}): void {
  replaceCurrentWorkbookAgentReviewItem(
    input.sessionState,
    createWorkbookAgentReviewQueueItem({
      sessionState: input.sessionState,
      bundle: input.bundle,
    }),
  )
  input.sessionState.durable.entries = upsertEntry(
    input.sessionState.durable.entries,
    createSystemEntry(
      `system-preview:${input.bundle.id}`,
      input.turnId,
      describeWorkbookAgentBundle(input.bundle),
      createBundleRangeCitations(input.bundle),
    ),
  )
}

export function requireWorkbookAgentReviewItem(input: {
  readonly reviewItem: WorkbookAgentReviewQueueItem | null
  readonly reviewItemId: string
  readonly notFoundMessage: string
}): WorkbookAgentReviewQueueItem {
  if (!input.reviewItem || input.reviewItem.id !== input.reviewItemId) {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_REVIEW_ITEM_NOT_FOUND',
      message: input.notFoundMessage,
      statusCode: 404,
      retryable: false,
    })
  }
  return input.reviewItem
}

export function createWorkbookAgentDismissReviewEntry(input: { readonly reviewItem: WorkbookAgentReviewQueueItem; readonly now: number }) {
  return createSystemEntry(
    `system-dismiss:${input.reviewItem.id}:${input.now}`,
    input.reviewItem.turnId,
    `Cleared workbook review item: ${input.reviewItem.summary}`,
    createBundleRangeCitations(toWorkbookAgentCommandBundle(input.reviewItem)),
  )
}

export function transitionWorkbookAgentSharedReview(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly reviewItem: WorkbookAgentReviewQueueItem
  readonly decision: 'approved' | 'rejected'
  readonly reviewerUserId: string
  readonly now: number
}): {
  readonly reviewedBundle: WorkbookAgentCommandBundle
  readonly nextReviewItem: WorkbookAgentReviewQueueItem
  readonly counter:
    | 'sharedReviewApprovedCount'
    | 'sharedReviewRejectedCount'
    | 'sharedRecommendationApprovedCount'
    | 'sharedRecommendationRejectedCount'
  readonly entryText: string
} {
  if (input.reviewItem.reviewMode !== 'ownerReview') {
    throw createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_SHARED_REVIEW_NOT_REQUIRED',
      message: 'Shared review is only required for medium/high-risk bundles.',
      statusCode: 409,
      retryable: false,
    })
  }

  const bundle = toWorkbookAgentCommandBundle(input.reviewItem)
  const sharedReview =
    normalizeSharedReviewState(bundle, input.sessionState) ?? createPendingSharedReviewState(input.sessionState.storageActorUserId)
  const isOwnerReviewer = input.sessionState.storageActorUserId === input.reviewerUserId
  const nextSharedReview: WorkbookAgentSharedReviewState = isOwnerReviewer
    ? {
        ...sharedReview,
        status: input.decision,
        decidedByUserId: input.reviewerUserId,
        decidedAtUnixMs: input.now,
      }
    : {
        ...sharedReview,
        recommendations: [
          ...sharedReview.recommendations.filter((recommendation) => recommendation.userId !== input.reviewerUserId),
          {
            userId: input.reviewerUserId,
            decision: input.decision,
            decidedAtUnixMs: input.now,
          },
        ].toSorted((left, right) => left.userId.localeCompare(right.userId)),
      }

  const reviewedBundle = {
    ...bundle,
    sharedReview: nextSharedReview,
  } satisfies WorkbookAgentCommandBundle

  return {
    reviewedBundle,
    nextReviewItem: toWorkbookAgentReviewQueueItem({
      bundle: reviewedBundle,
      reviewMode: 'ownerReview',
      sharedReview: nextSharedReview,
    }),
    counter: isOwnerReviewer
      ? input.decision === 'approved'
        ? 'sharedReviewApprovedCount'
        : 'sharedReviewRejectedCount'
      : input.decision === 'approved'
        ? 'sharedRecommendationApprovedCount'
        : 'sharedRecommendationRejectedCount',
    entryText: isOwnerReviewer
      ? `${input.decision === 'approved' ? 'Approved' : 'Returned'} shared review item: ${reviewedBundle.summary}`
      : `${input.reviewerUserId} shared a ${input.decision === 'approved' ? 'ready-to-apply' : 'return-for-edit'} review recommendation: ${reviewedBundle.summary}`,
  }
}
