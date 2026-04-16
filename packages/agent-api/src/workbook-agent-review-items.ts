import type {
  WorkbookAgentCommand,
  WorkbookAgentBundleScope,
  WorkbookAgentCommandBundle,
  WorkbookAgentContextRef,
  WorkbookAgentPreviewRange,
  WorkbookAgentRiskClass,
  WorkbookAgentSharedReviewRecommendation,
  WorkbookAgentSharedReviewState,
  WorkbookAgentSharedReviewStatus,
} from './workbook-agent-bundles.js'

export type WorkbookAgentReviewMode = 'manual' | 'ownerReview'

export interface WorkbookAgentReviewQueueItem {
  id: string
  documentId: string
  threadId: string
  turnId: string
  goalText: string
  summary: string
  scope: WorkbookAgentBundleScope
  riskClass: WorkbookAgentRiskClass
  reviewMode: WorkbookAgentReviewMode
  ownerUserId: string | null
  status: WorkbookAgentSharedReviewStatus
  decidedByUserId: string | null
  decidedAtUnixMs: number | null
  recommendations: WorkbookAgentSharedReviewRecommendation[]
  baseRevision: number
  createdAtUnixMs: number
  context: WorkbookAgentContextRef | null
  commands: WorkbookAgentCommand[]
  affectedRanges: WorkbookAgentPreviewRange[]
  estimatedAffectedCells: number | null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSharedReviewRecommendation(value: unknown): value is WorkbookAgentSharedReviewRecommendation {
  return (
    isRecord(value) &&
    typeof value['userId'] === 'string' &&
    (value['decision'] === 'approved' || value['decision'] === 'rejected') &&
    typeof value['decidedAtUnixMs'] === 'number'
  )
}

export function toWorkbookAgentReviewQueueItem(input: {
  bundle: WorkbookAgentCommandBundle
  reviewMode: WorkbookAgentReviewMode
  sharedReview?: WorkbookAgentSharedReviewState | null
}): WorkbookAgentReviewQueueItem {
  const sharedReview = input.sharedReview ?? input.bundle.sharedReview ?? null
  return {
    id: input.bundle.id,
    documentId: input.bundle.documentId,
    threadId: input.bundle.threadId,
    turnId: input.bundle.turnId,
    goalText: input.bundle.goalText,
    summary: input.bundle.summary,
    scope: input.bundle.scope,
    riskClass: input.bundle.riskClass,
    reviewMode: input.reviewMode,
    ownerUserId: sharedReview?.ownerUserId ?? null,
    status: sharedReview?.status ?? 'pending',
    decidedByUserId: sharedReview?.decidedByUserId ?? null,
    decidedAtUnixMs: sharedReview?.decidedAtUnixMs ?? null,
    recommendations: [...(sharedReview?.recommendations ?? [])],
    baseRevision: input.bundle.baseRevision,
    createdAtUnixMs: input.bundle.createdAtUnixMs,
    context: input.bundle.context,
    commands: input.bundle.commands.map((command) => structuredClone(command)),
    affectedRanges: input.bundle.affectedRanges.map((range) => structuredClone(range)),
    estimatedAffectedCells: input.bundle.estimatedAffectedCells,
  }
}

export function toWorkbookAgentCommandBundle(reviewItem: WorkbookAgentReviewQueueItem): WorkbookAgentCommandBundle {
  return {
    id: reviewItem.id,
    documentId: reviewItem.documentId,
    threadId: reviewItem.threadId,
    turnId: reviewItem.turnId,
    goalText: reviewItem.goalText,
    summary: reviewItem.summary,
    scope: reviewItem.scope,
    riskClass: reviewItem.riskClass,
    baseRevision: reviewItem.baseRevision,
    createdAtUnixMs: reviewItem.createdAtUnixMs,
    context: reviewItem.context,
    commands: reviewItem.commands.map((command) => structuredClone(command)),
    affectedRanges: reviewItem.affectedRanges.map((range) => structuredClone(range)),
    estimatedAffectedCells: reviewItem.estimatedAffectedCells,
    sharedReview:
      reviewItem.reviewMode === 'ownerReview'
        ? {
            ownerUserId: reviewItem.ownerUserId ?? '',
            status: reviewItem.status,
            decidedByUserId: reviewItem.decidedByUserId,
            decidedAtUnixMs: reviewItem.decidedAtUnixMs,
            recommendations: [...reviewItem.recommendations],
          }
        : null,
  }
}

export function isWorkbookAgentReviewQueueItem(value: unknown): value is WorkbookAgentReviewQueueItem {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['documentId'] === 'string' &&
    typeof value['threadId'] === 'string' &&
    typeof value['turnId'] === 'string' &&
    typeof value['goalText'] === 'string' &&
    typeof value['summary'] === 'string' &&
    (value['scope'] === 'selection' || value['scope'] === 'sheet' || value['scope'] === 'workbook') &&
    (value['riskClass'] === 'low' || value['riskClass'] === 'medium' || value['riskClass'] === 'high') &&
    (value['reviewMode'] === 'manual' || value['reviewMode'] === 'ownerReview') &&
    (value['ownerUserId'] === null || typeof value['ownerUserId'] === 'string') &&
    (value['status'] === 'pending' || value['status'] === 'approved' || value['status'] === 'rejected') &&
    (value['decidedByUserId'] === null || typeof value['decidedByUserId'] === 'string') &&
    (value['decidedAtUnixMs'] === null || typeof value['decidedAtUnixMs'] === 'number') &&
    Array.isArray(value['recommendations']) &&
    value['recommendations'].every((entry) => isSharedReviewRecommendation(entry)) &&
    typeof value['baseRevision'] === 'number' &&
    typeof value['createdAtUnixMs'] === 'number' &&
    Array.isArray(value['commands']) &&
    Array.isArray(value['affectedRanges']) &&
    (value['estimatedAffectedCells'] === null || typeof value['estimatedAffectedCells'] === 'number')
  )
}
