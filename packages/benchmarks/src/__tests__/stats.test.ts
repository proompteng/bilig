import { describe, expect, it } from "vitest";
import { quantile, summarizeNumbers } from "../stats.js";

describe("benchmark stats", () => {
  it("computes nearest-rank percentiles deterministically", () => {
    const values = [10, 20, 30, 40, 50];
    expect(quantile(values, 0)).toBe(10);
    expect(quantile(values, 0.5)).toBe(30);
    expect(quantile(values, 0.95)).toBe(50);
    expect(quantile(values, 1)).toBe(50);
  });

  it("summarizes samples into min/median/p95/max/mean", () => {
    const summary = summarizeNumbers([7, 1, 5, 3, 9]);
    expect(summary.samples).toEqual([1, 3, 5, 7, 9]);
    expect(summary.min).toBe(1);
    expect(summary.median).toBe(5);
    expect(summary.p95).toBe(9);
    expect(summary.max).toBe(9);
    expect(summary.mean).toBe(5);
  });
});
