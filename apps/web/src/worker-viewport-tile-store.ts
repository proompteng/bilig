import type { WorkbookLocalStore, WorkbookLocalViewportBase } from "@bilig/storage-browser";
import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import type { ViewportPatchSubscription } from "@bilig/worker-transport";

const VIEWPORT_TILE_COLUMN_COUNT = 128;
const VIEWPORT_TILE_ROW_COUNT = 32;
const DEFAULT_MAX_CACHED_TILES = 96;

type ViewportBounds = Pick<
  ViewportPatchSubscription,
  "rowStart" | "rowEnd" | "colStart" | "colEnd"
>;

interface CachedViewportTile {
  readonly key: string;
  readonly sheetId: number;
  readonly sheetName: string;
  readonly bounds: ViewportBounds;
  readonly viewport: WorkbookLocalViewportBase;
  accessTick: number;
}

function buildTileKey(sheetId: number, bounds: ViewportBounds): string {
  return `${sheetId}:${bounds.rowStart}:${bounds.colStart}`;
}

function sortViewportCells(
  left: WorkbookLocalViewportBase["cells"][number],
  right: WorkbookLocalViewportBase["cells"][number],
): number {
  return left.row - right.row || left.col - right.col;
}

function normalizeViewportBounds(viewport: ViewportBounds): ViewportBounds {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, viewport.rowStart));
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, viewport.rowEnd));
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, viewport.colStart));
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, viewport.colEnd));
  return {
    rowStart,
    rowEnd,
    colStart,
    colEnd,
  };
}

export function listViewportTileBounds(viewport: ViewportBounds): ViewportBounds[] {
  const normalized = normalizeViewportBounds(viewport);
  const rowTileStart = Math.floor(normalized.rowStart / VIEWPORT_TILE_ROW_COUNT);
  const rowTileEnd = Math.floor(normalized.rowEnd / VIEWPORT_TILE_ROW_COUNT);
  const colTileStart = Math.floor(normalized.colStart / VIEWPORT_TILE_COLUMN_COUNT);
  const colTileEnd = Math.floor(normalized.colEnd / VIEWPORT_TILE_COLUMN_COUNT);
  const bounds: ViewportBounds[] = [];

  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT;
      const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT;
      bounds.push({
        rowStart,
        rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
        colStart,
        colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
      });
    }
  }

  return bounds;
}

function mergeViewportTiles(
  viewport: ViewportBounds,
  tiles: readonly WorkbookLocalViewportBase[],
): WorkbookLocalViewportBase {
  const firstTile = tiles[0];
  if (!firstTile) {
    throw new Error("Cannot merge an empty tile set");
  }

  const cells = new Map<string, WorkbookLocalViewportBase["cells"][number]>();
  const rowAxisEntries = new Map<number, WorkbookLocalViewportBase["rowAxisEntries"][number]>();
  const columnAxisEntries = new Map<
    number,
    WorkbookLocalViewportBase["columnAxisEntries"][number]
  >();
  const styles = new Map<string, WorkbookLocalViewportBase["styles"][number]>();

  tiles.forEach((tile) => {
    tile.cells.forEach((cell) => {
      if (
        cell.row < viewport.rowStart ||
        cell.row > viewport.rowEnd ||
        cell.col < viewport.colStart ||
        cell.col > viewport.colEnd
      ) {
        return;
      }
      cells.set(cell.snapshot.address, cell);
    });
    tile.rowAxisEntries.forEach((entry) => {
      if (entry.index < viewport.rowStart || entry.index > viewport.rowEnd) {
        return;
      }
      rowAxisEntries.set(entry.index, entry);
    });
    tile.columnAxisEntries.forEach((entry) => {
      if (entry.index < viewport.colStart || entry.index > viewport.colEnd) {
        return;
      }
      columnAxisEntries.set(entry.index, entry);
    });
    tile.styles.forEach((style) => {
      styles.set(style.id, style);
    });
  });

  return {
    sheetId: firstTile.sheetId,
    sheetName: firstTile.sheetName,
    cells: [...cells.values()].toSorted(sortViewportCells),
    rowAxisEntries: [...rowAxisEntries.values()].toSorted(
      (left, right) => left.index - right.index,
    ),
    columnAxisEntries: [...columnAxisEntries.values()].toSorted(
      (left, right) => left.index - right.index,
    ),
    styles: [...styles.values()],
  };
}

export class WorkerViewportTileStore {
  private readonly tiles = new Map<string, CachedViewportTile>();
  private readonly tileKeysBySheetId = new Map<number, Set<string>>();
  private readonly sheetIdsByName = new Map<string, number>();
  private nextAccessTick = 1;

  constructor(private readonly maxCachedTiles = DEFAULT_MAX_CACHED_TILES) {}

  reset(): void {
    this.tiles.clear();
    this.tileKeysBySheetId.clear();
    this.sheetIdsByName.clear();
    this.nextAccessTick = 1;
  }

  invalidateSheet(sheetId: number): void {
    const keys = this.tileKeysBySheetId.get(sheetId);
    if (keys) {
      keys.forEach((key) => {
        this.tiles.delete(key);
      });
    }
    this.tileKeysBySheetId.delete(sheetId);
    for (const [sheetName, mappedSheetId] of this.sheetIdsByName) {
      if (mappedSheetId === sheetId) {
        this.sheetIdsByName.delete(sheetName);
      }
    }
  }

  hasTile(sheetId: number, bounds: ViewportBounds): boolean {
    return this.tiles.has(buildTileKey(sheetId, bounds));
  }

  readViewport(input: {
    readonly localStore: Pick<WorkbookLocalStore, "readViewportProjection">;
    readonly sheetName: string;
    readonly viewport: ViewportBounds;
  }): WorkbookLocalViewportBase | null {
    const tiles = listViewportTileBounds(input.viewport);
    const resolvedTiles: WorkbookLocalViewportBase[] = [];

    for (const bounds of tiles) {
      let cachedTile: CachedViewportTile | undefined;
      const knownSheetId = this.sheetIdsByName.get(input.sheetName);
      if (knownSheetId !== undefined) {
        cachedTile = this.tiles.get(buildTileKey(knownSheetId, bounds));
      }

      if (!cachedTile) {
        const viewport = input.localStore.readViewportProjection(input.sheetName, bounds);
        if (!viewport) {
          return null;
        }
        this.updateSheetNameMapping(input.sheetName, viewport.sheetName, viewport.sheetId);
        cachedTile = this.storeTile(bounds, viewport);
      } else {
        this.touch(cachedTile);
      }

      resolvedTiles.push(cachedTile.viewport);
    }

    if (resolvedTiles.length === 0) {
      return null;
    }
    return mergeViewportTiles(normalizeViewportBounds(input.viewport), resolvedTiles);
  }

  private touch(tile: CachedViewportTile): void {
    tile.accessTick = this.nextAccessTick++;
  }

  private storeTile(
    bounds: ViewportBounds,
    viewport: WorkbookLocalViewportBase,
  ): CachedViewportTile {
    const key = buildTileKey(viewport.sheetId, bounds);
    const cached: CachedViewportTile = {
      key,
      sheetId: viewport.sheetId,
      sheetName: viewport.sheetName,
      bounds,
      viewport,
      accessTick: this.nextAccessTick++,
    };
    this.tiles.set(key, cached);
    const sheetTileKeys = this.tileKeysBySheetId.get(viewport.sheetId) ?? new Set<string>();
    sheetTileKeys.add(key);
    this.tileKeysBySheetId.set(viewport.sheetId, sheetTileKeys);
    this.evictLeastRecentlyUsedTiles();
    return cached;
  }

  private updateSheetNameMapping(
    requestedSheetName: string,
    actualSheetName: string,
    sheetId: number,
  ): void {
    const requestedSheetId = this.sheetIdsByName.get(requestedSheetName);
    if (requestedSheetId !== undefined && requestedSheetId !== sheetId) {
      this.invalidateSheet(requestedSheetId);
    }
    const actualSheetId = this.sheetIdsByName.get(actualSheetName);
    if (actualSheetId !== undefined && actualSheetId !== sheetId) {
      this.invalidateSheet(actualSheetId);
    }
    this.sheetIdsByName.set(actualSheetName, sheetId);
    this.sheetIdsByName.set(requestedSheetName, sheetId);
  }

  private evictLeastRecentlyUsedTiles(): void {
    while (this.tiles.size > this.maxCachedTiles) {
      let oldestTile: CachedViewportTile | undefined;
      for (const tile of this.tiles.values()) {
        if (!oldestTile || tile.accessTick < oldestTile.accessTick) {
          oldestTile = tile;
        }
      }
      if (!oldestTile) {
        return;
      }
      this.tiles.delete(oldestTile.key);
      const sheetKeys = this.tileKeysBySheetId.get(oldestTile.sheetId);
      sheetKeys?.delete(oldestTile.key);
      if (!sheetKeys || sheetKeys.size === 0) {
        this.tileKeysBySheetId.delete(oldestTile.sheetId);
      }
    }
  }
}
