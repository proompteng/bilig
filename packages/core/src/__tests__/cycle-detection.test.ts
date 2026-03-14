import { describe, expect, it } from "vitest";
import { CycleDetector } from "../cycle-detection.js";

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
      (cellIndex, fn) => {
        const dependencies = graph.get(cellIndex) ?? [];
        for (let index = 0; index < dependencies.length; index += 1) {
          fn(dependencies[index]!);
        }
      },
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
      (cellIndex, fn) => {
        if (cellIndex === 1) fn(2);
        if (cellIndex === 2) fn(1);
      },
      (cellIndex) => cellIndex === 1 || cellIndex === 2
    );
    const second = detector.detect(
      [3],
      4,
      (cellIndex, fn) => {
        if (cellIndex === 3) fn(3);
      },
      (cellIndex) => cellIndex === 3
    );

    expect(second.cycleMembers).toBe(first.cycleMembers);
    expect(second.cycleMemberCount).toBe(1);
    expect(second.cycleGroups[1]).toBe(-1);
    expect(second.cycleGroups[2]).toBe(-1);
    expect(second.cycleGroups[3]).toBe(0);
  });
  it("returns packed cycle members and group ids directly", () => {
    const detector = new CycleDetector();
    const result = detector.detect(
      [1, 2],
      4,
      (cellIndex, fn) => {
        if (cellIndex === 1) fn(2);
        if (cellIndex === 2) fn(1);
      },
      (cellIndex) => cellIndex === 1 || cellIndex === 2
    );

    expect(Array.from(result.cycleMembers.slice(0, result.cycleMemberCount))).toEqual([2, 1]);
    expect(result.cycleGroups[1]).toBe(0);
    expect(result.cycleGroups[2]).toBe(0);
  });
});
