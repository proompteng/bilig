import { describe, expect, it } from "vitest";
import {
  describeWorkbookAgentExecutionPolicy,
  isWorkbookAgentBundleAutoApplyEligible,
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentReviewDisposition,
} from "../workbook-agent-execution-policy.js";

describe("workbook agent execution policy", () => {
  it("routes private auto-apply-all sessions straight to execution", () => {
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: "private",
        executionPolicy: "autoApplyAll",
        riskClass: "high",
      }),
    ).toBe(true);
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: "private",
        executionPolicy: "autoApplyAll",
        riskClass: "high",
      }),
    ).toBe("applyAutomatically");
  });

  it("routes private safe sessions by risk class", () => {
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: "private",
        executionPolicy: "autoApplySafe",
        riskClass: "low",
      }),
    ).toBe(true);
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: "private",
        executionPolicy: "autoApplySafe",
        riskClass: "medium",
      }),
    ).toBe(false);
  });

  it("routes shared medium and high risk work through owner review", () => {
    expect(
      requiresWorkbookAgentOwnerReview({
        scope: "shared",
        riskClass: "medium",
      }),
    ).toBe(true);
    expect(
      requiresWorkbookAgentOwnerReview({
        scope: "shared",
        riskClass: "low",
      }),
    ).toBe(false);
  });

  it("describes execution policies for user-facing summaries", () => {
    expect(describeWorkbookAgentExecutionPolicy("autoApplySafe")).toBe("auto-apply safe changes");
    expect(describeWorkbookAgentExecutionPolicy("autoApplyAll")).toBe("auto-apply all changes");
    expect(describeWorkbookAgentExecutionPolicy("ownerReview")).toBe("owner review");
  });
});
