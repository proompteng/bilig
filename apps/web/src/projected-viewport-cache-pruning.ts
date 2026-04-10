import { parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, Viewport } from "@bilig/protocol";

export function selectProjectedViewportKeysToEvict(args: {
  sheetCellKeys: readonly string[];
  cellSnapshots: ReadonlyMap<string, CellSnapshot>;
  cellAccessTicks: ReadonlyMap<string, number>;
  pinnedKeys: ReadonlySet<string>;
  activeViewports: readonly Viewport[];
  maxCachedCellsPerSheet: number;
}): string[] {
  const overflow = args.sheetCellKeys.length - args.maxCachedCellsPerSheet;
  if (overflow <= 0) {
    return [];
  }

  const missingSnapshotKeys: string[] = [];
  const candidates: Array<{ key: string; accessTick: number }> = [];

  for (const key of args.sheetCellKeys) {
    const snapshot = args.cellSnapshots.get(key);
    if (!snapshot) {
      missingSnapshotKeys.push(key);
      continue;
    }
    if (args.pinnedKeys.has(key)) {
      continue;
    }
    const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
    const insideActiveViewport = args.activeViewports.some((viewport) => {
      return (
        parsed.row >= viewport.rowStart &&
        parsed.row <= viewport.rowEnd &&
        parsed.col >= viewport.colStart &&
        parsed.col <= viewport.colEnd
      );
    });
    if (insideActiveViewport) {
      continue;
    }
    candidates.push({
      key,
      accessTick: args.cellAccessTicks.get(key) ?? 0,
    });
  }

  candidates.sort(
    (left, right) => left.accessTick - right.accessTick || left.key.localeCompare(right.key),
  );
  return [...missingSnapshotKeys, ...candidates.map((candidate) => candidate.key)].slice(
    0,
    overflow,
  );
}
