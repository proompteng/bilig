import {
  isWorkbookAgentReviewQueueItem,
  type WorkbookAgentReviewQueueItem,
} from "@bilig/agent-api";

export function decodeWorkbookAgentReviewItems(
  reviewQueueItems: unknown,
): WorkbookAgentReviewQueueItem[] {
  if (!Array.isArray(reviewQueueItems)) {
    return [];
  }
  return reviewQueueItems.flatMap((item) => (isWorkbookAgentReviewQueueItem(item) ? [item] : []));
}

export function resolvePrimaryWorkbookAgentReviewItem(
  reviewQueueItems: readonly WorkbookAgentReviewQueueItem[],
): WorkbookAgentReviewQueueItem | null {
  return reviewQueueItems[0] ?? null;
}

export function resolveWorkbookAgentReviewOwnerUserId(input: {
  readonly reviewItem: WorkbookAgentReviewQueueItem | null;
  readonly sessionScope: "private" | "shared";
  readonly activeThreadOwnerUserId: string | null;
}): string | null {
  if (!input.reviewItem || input.reviewItem.reviewMode !== "ownerReview") {
    return null;
  }
  return input.reviewItem.ownerUserId ?? input.activeThreadOwnerUserId ?? null;
}
