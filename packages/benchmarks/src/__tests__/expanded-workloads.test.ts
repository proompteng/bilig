import { describe, expect, it } from "vitest";
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  type ExpandedComparativeBenchmarkWorkload,
} from "../benchmark-workpaper-vs-hyperformula-expanded.js";

describe("expanded comparative benchmark workloads", () => {
  it("includes the new workload families without duplicates", () => {
    const expectedNewWorkloads: ExpandedComparativeBenchmarkWorkload[] = [
      "build-parser-cache-row-templates",
      "rebuild-and-recalculate",
      "rebuild-config-toggle",
      "partial-recompute-mixed-frontier",
      "batch-suspended-single-column",
      "batch-suspended-multi-column",
      "structural-insert-rows",
      "aggregate-overlapping-ranges",
      "conditional-aggregation-reused-ranges",
      "lookup-with-column-index-after-column-write",
      "lookup-approximate-sorted-after-column-write",
    ];

    expect(new Set(EXPANDED_COMPARATIVE_WORKLOADS).size).toBe(
      EXPANDED_COMPARATIVE_WORKLOADS.length,
    );
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expect.arrayContaining(expectedNewWorkloads));
  });
});
