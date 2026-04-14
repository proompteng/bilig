import { describe, expect, it } from "vitest";
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  type ExpandedComparativeBenchmarkWorkload,
} from "../benchmark-workpaper-vs-hyperformula-expanded.js";

describe("expanded comparative benchmark workloads", () => {
  it("includes the new workload families without duplicates", () => {
    const expectedNewWorkloads: ExpandedComparativeBenchmarkWorkload[] = [
      "build-parser-cache-row-templates",
      "build-parser-cache-mixed-templates",
      "rebuild-and-recalculate",
      "rebuild-config-toggle",
      "rebuild-runtime-from-snapshot",
      "partial-recompute-mixed-frontier",
      "batch-edit-single-column-with-undo",
      "batch-suspended-single-column",
      "batch-suspended-multi-column",
      "structural-insert-rows",
      "structural-delete-rows",
      "structural-move-rows",
      "structural-insert-columns",
      "structural-delete-columns",
      "structural-move-columns",
      "aggregate-overlapping-ranges",
      "aggregate-overlapping-sliding-window",
      "conditional-aggregation-reused-ranges",
      "conditional-aggregation-criteria-cell-edit",
      "lookup-with-column-index-after-column-write",
      "lookup-with-column-index-after-batch-write",
      "lookup-approximate-sorted-after-column-write",
    ];

    expect(new Set(EXPANDED_COMPARATIVE_WORKLOADS).size).toBe(
      EXPANDED_COMPARATIVE_WORKLOADS.length,
    );
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expect.arrayContaining(expectedNewWorkloads));
  });
});
