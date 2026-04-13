import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { ValueTag, type CellSnapshot, type Viewport } from "@bilig/protocol";
import { runProperty } from "@bilig/test-fuzz";
import { selectProjectedViewportKeysToEvict } from "../projected-viewport-cache-pruning.js";

describe("viewport cache pruning fuzz", () => {
  it("matches the cache eviction specification across random cache states", async () => {
    await runProperty({
      suite: "web/projected-viewport/cache-pruning-spec",
      arbitrary: cacheStateArbitrary,
      predicate: async (state) => {
        expect(selectProjectedViewportKeysToEvict(state)).toEqual(referenceEviction(state));
      },
    });
  });
});

// Helpers

const cellPositionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 4 }),
    col: fc.integer({ min: 0, max: 4 }),
  })
  .map(({ row, col }) => ({
    key: `Sheet1!${String.fromCharCode(65 + col)}${row + 1}`,
    snapshot: {
      sheetName: "Sheet1",
      address: `${String.fromCharCode(65 + col)}${row + 1}`,
      value: { tag: ValueTag.Number, value: row * 10 + col },
      flags: 0,
      version: 1,
    } satisfies CellSnapshot,
    row,
    col,
  }));

const cacheStateArbitrary = fc
  .uniqueArray(cellPositionArbitrary, {
    minLength: 1,
    maxLength: 8,
    selector: (entry) => entry.key,
  })
  .chain((entries) =>
    fc
      .record({
        maxCachedCellsPerSheet: fc.integer({ min: 1, max: 6 }),
        missingKeys: fc.subarray(entries.map((entry) => entry.key)),
        pinnedKeys: fc.subarray(entries.map((entry) => entry.key)),
        activeViewports: fc.array(
          fc.record({
            sheetName: fc.constant("Sheet1"),
            rowStart: fc.integer({ min: 0, max: 4 }),
            rowEnd: fc.integer({ min: 0, max: 4 }),
            colStart: fc.integer({ min: 0, max: 4 }),
            colEnd: fc.integer({ min: 0, max: 4 }),
          }),
          { maxLength: 3 },
        ),
        accessTicks: fc.dictionary(
          fc.constantFrom(...entries.map((entry) => entry.key)),
          fc.integer({ min: 0, max: 50 }),
        ),
      })
      .map((raw) => {
        const cellSnapshots = new Map<string, CellSnapshot>();
        entries.forEach((entry) => {
          if (!raw.missingKeys.includes(entry.key)) {
            cellSnapshots.set(entry.key, entry.snapshot);
          }
        });
        return {
          sheetCellKeys: entries.map((entry) => entry.key),
          cellSnapshots,
          cellAccessTicks: new Map(Object.entries(raw.accessTicks)),
          pinnedKeys: new Set(raw.pinnedKeys),
          activeViewports: raw.activeViewports.map(normalizeViewport),
          maxCachedCellsPerSheet: raw.maxCachedCellsPerSheet,
        };
      }),
  );

function referenceEviction(
  args: Parameters<typeof selectProjectedViewportKeysToEvict>[0],
): string[] {
  const overflow = args.sheetCellKeys.length - args.maxCachedCellsPerSheet;
  if (overflow <= 0) {
    return [];
  }

  const missingSnapshotKeys: string[] = [];
  const intactCandidates: Array<{ key: string; accessTick: number }> = [];
  for (const key of args.sheetCellKeys) {
    const snapshot = args.cellSnapshots.get(key);
    if (!snapshot) {
      missingSnapshotKeys.push(key);
      continue;
    }
    if (args.pinnedKeys.has(key)) {
      continue;
    }
    const parsed = parseAddress(snapshot.address);
    const insideViewport = args.activeViewports.some(
      (viewport) =>
        parsed.row >= viewport.rowStart &&
        parsed.row <= viewport.rowEnd &&
        parsed.col >= viewport.colStart &&
        parsed.col <= viewport.colEnd,
    );
    if (insideViewport) {
      continue;
    }
    intactCandidates.push({
      key,
      accessTick: args.cellAccessTicks.get(key) ?? 0,
    });
  }

  intactCandidates.sort(
    (left, right) => left.accessTick - right.accessTick || left.key.localeCompare(right.key),
  );
  return [...missingSnapshotKeys, ...intactCandidates.map((candidate) => candidate.key)].slice(
    0,
    overflow,
  );
}

function normalizeViewport(viewport: Viewport): Viewport {
  return {
    ...viewport,
    rowStart: Math.min(viewport.rowStart, viewport.rowEnd),
    rowEnd: Math.max(viewport.rowStart, viewport.rowEnd),
    colStart: Math.min(viewport.colStart, viewport.colEnd),
    colEnd: Math.max(viewport.colStart, viewport.colEnd),
  };
}

function parseAddress(address: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) {
    throw new Error(`Invalid address: ${address}`);
  }
  return {
    col: match[1].charCodeAt(0) - 65,
    row: Number(match[2]) - 1,
  };
}
