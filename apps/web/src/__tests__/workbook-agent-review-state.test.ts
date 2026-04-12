import { describe, expect, it } from "vitest";
import {
  decodeWorkbookAgentReviewItem,
  resolveWorkbookAgentReviewOwnerUserId,
} from "../workbook-agent-review-state.js";

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    sessionId: "agent-session-1",
    documentId: "doc-1",
    threadId: "thr-1",
    executionPolicy: "autoApplyAll",
    scope: "private",
    status: "idle",
    activeTurnId: null,
    lastError: null,
    context: null,
    entries: [],
    pendingBundle: null,
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  };
}

describe("workbook agent review state", () => {
  it("decodes the current review item from the session snapshot", () => {
    const snapshot = createSnapshot({
      pendingBundle: {
        id: "bundle-1",
        documentId: "doc-1",
        threadId: "thr-1",
        turnId: "turn-1",
        goalText: "Bold the current cell",
        summary: "Format Sheet1!A1",
        scope: "selection",
        riskClass: "low",
        approvalMode: "auto",
        baseRevision: 1,
        createdAtUnixMs: 1,
        context: null,
        commands: [],
        affectedRanges: [],
        estimatedAffectedCells: 1,
      },
    });

    expect(decodeWorkbookAgentReviewItem(snapshot.pendingBundle)?.id).toBe("bundle-1");
  });

  it("derives owner review state for shared high-risk work", () => {
    const reviewItem = {
      riskClass: "high",
      sharedReview: {
        ownerUserId: "alex@example.com",
        status: "pending",
        decidedByUserId: null,
        decidedAtUnixMs: null,
        recommendations: [],
      },
    } as const;

    expect(
      resolveWorkbookAgentReviewOwnerUserId({
        reviewItem,
        sessionScope: "shared",
        activeThreadOwnerUserId: "casey@example.com",
      }),
    ).toBe("alex@example.com");
  });
});
