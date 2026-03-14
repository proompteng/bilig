import { describe, expect, it } from "vitest";
import { CycleDetector, detectFormulaCycles } from "../cycle-detection.js";

describe("CycleDetector", () => {
  it("assigns deterministic group ids for separate strongly connected components", () => {
    const graph = new Map<number, number[]>([
      [1, [2]],
      [2, [1]],
      [3, [4]],
      [4, [3]]
    ]);

    const detector = new CycleDetector();
    const result = detector.detect(
      [1, 2, 3, 4],
      8,
      (cellIndex) => graph.get(cellIndex) ?? [],
      (cellIndex) => graph.has(cellIndex)
    );

    expect(result.cycleMemberCount).toBe(4);
    expect(result.cycleGroups[1]).toBe(result.cycleGroups[2]);
    expect(result.cycleGroups[3]).toBe(result.cycleGroups[4]);
    expect(result.cycleGroups[1]).not.toBe(result.cycleGroups[3]);
  });

  it("reuses packed buffers across detections while clearing stale memberships", () => {
    const detector = new CycleDetector();

    const first = detector.detect(
      [1, 2],
      4,
      (cellIndex) => (cellIndex === 1 ? [2] : cellIndex === 2 ? [1] : []),
      (cellIndex) => cellIndex === 1 || cellIndex === 2
    );
    const second = detector.detect([3], 4, () => [3], (cellIndex) => cellIndex === 3);

    expect(second.cycleMembers).toBe(first.cycleMembers);
    expect(second.cycleMemberCount).toBe(1);
    expect(second.cycleGroups[1]).toBe(-1);
    expect(second.cycleGroups[2]).toBe(-1);
    expect(second.cycleGroups[3]).toBe(0);
  });
});

describe("detectFormulaCycles", () => {
  it("preserves the compatibility wrapper for set/map consumers", () => {
    const result = detectFormulaCycles(
      [1, 2],
      4,
      (cellIndex) => (cellIndex === 1 ? [2] : cellIndex === 2 ? [1] : []),
      (cellIndex) => cellIndex === 1 || cellIndex === 2
    );

    expect(result.inCycle).toEqual(new Set([1, 2]));
    expect(result.cycleGroups.get(1)).toBe(0);
    expect(result.cycleGroups.get(2)).toBe(0);
  });
});
