import {
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
} from "@bilig/agent-api";

export function decodeWorkbookAgentReviewItem(
  pendingBundle: unknown,
): WorkbookAgentCommandBundle | null {
  return isWorkbookAgentCommandBundle(pendingBundle) ? pendingBundle : null;
}

export function resolveWorkbookAgentReviewOwnerUserId(input: {
  readonly reviewItem: WorkbookAgentCommandBundle | null;
  readonly sessionScope: "private" | "shared";
  readonly activeThreadOwnerUserId: string | null;
}): string | null {
  if (input.sessionScope !== "shared" || input.reviewItem?.riskClass === "low") {
    return null;
  }
  return input.reviewItem?.sharedReview?.ownerUserId ?? input.activeThreadOwnerUserId ?? null;
}
