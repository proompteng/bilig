import type { WorkbookLocalStore, WorkbookLocalViewportBase } from '@bilig/storage-browser'
import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import type { ViewportPatchSubscription } from '@bilig/worker-transport'
import { MAX_TILE_SHEET_ORDINAL, packTileKey53 } from '../../../packages/grid/src/renderer-v3/tile-key.js'
import { TileResidencyV3, type TileEntryV3 } from '../../../packages/grid/src/renderer-v3/tile-residency.js'
const DEFAULT_MAX_CACHED_TILES = 96
const DEFAULT_MAX_CACHED_TILE_BYTES = 128 * 1024 * 1024

type ViewportBounds = Pick<ViewportPatchSubscription, 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>

interface CachedViewportTile {
  readonly key: number
  readonly sheetId: number
  readonly sheetName: string
  readonly bounds: ViewportBounds
  readonly viewport: WorkbookLocalViewportBase
}

function buildTileKey(sheetId: number, bounds: ViewportBounds): number {
  return packTileKey53({
    sheetOrdinal: sheetId,
    rowTile: Math.floor(bounds.rowStart / VIEWPORT_TILE_ROW_COUNT),
    colTile: Math.floor(bounds.colStart / VIEWPORT_TILE_COLUMN_COUNT),
    dprBucket: 0,
  })
}

function sortViewportCells(left: WorkbookLocalViewportBase['cells'][number], right: WorkbookLocalViewportBase['cells'][number]): number {
  return left.row - right.row || left.col - right.col
}

function normalizeViewportBounds(viewport: ViewportBounds): ViewportBounds {
  const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, viewport.rowStart))
  const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, viewport.rowEnd))
  const colStart = Math.max(0, Math.min(MAX_COLS - 1, viewport.colStart))
  const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, viewport.colEnd))
  return {
    rowStart,
    rowEnd,
    colStart,
    colEnd,
  }
}

export function listViewportTileBounds(viewport: ViewportBounds): ViewportBounds[] {
  const normalized = normalizeViewportBounds(viewport)
  const rowTileStart = Math.floor(normalized.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(normalized.rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(normalized.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(normalized.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  const bounds: ViewportBounds[] = []

  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      const rowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
      const colStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
      bounds.push({
        rowStart,
        rowEnd: Math.min(MAX_ROWS - 1, rowStart + VIEWPORT_TILE_ROW_COUNT - 1),
        colStart,
        colEnd: Math.min(MAX_COLS - 1, colStart + VIEWPORT_TILE_COLUMN_COUNT - 1),
      })
    }
  }

  return bounds
}

function mergeViewportTiles(viewport: ViewportBounds, tiles: readonly WorkbookLocalViewportBase[]): WorkbookLocalViewportBase {
  const firstTile = tiles[0]
  if (!firstTile) {
    throw new Error('Cannot merge an empty tile set')
  }

  const cells = new Map<string, WorkbookLocalViewportBase['cells'][number]>()
  const rowAxisEntries = new Map<number, WorkbookLocalViewportBase['rowAxisEntries'][number]>()
  const columnAxisEntries = new Map<number, WorkbookLocalViewportBase['columnAxisEntries'][number]>()
  const styles = new Map<string, WorkbookLocalViewportBase['styles'][number]>()

  tiles.forEach((tile) => {
    tile.cells.forEach((cell) => {
      if (cell.row < viewport.rowStart || cell.row > viewport.rowEnd || cell.col < viewport.colStart || cell.col > viewport.colEnd) {
        return
      }
      cells.set(cell.snapshot.address, cell)
    })
    tile.rowAxisEntries.forEach((entry) => {
      if (entry.index < viewport.rowStart || entry.index > viewport.rowEnd) {
        return
      }
      rowAxisEntries.set(entry.index, entry)
    })
    tile.columnAxisEntries.forEach((entry) => {
      if (entry.index < viewport.colStart || entry.index > viewport.colEnd) {
        return
      }
      columnAxisEntries.set(entry.index, entry)
    })
    tile.styles.forEach((style) => {
      styles.set(style.id, style)
    })
  })

  return {
    sheetId: firstTile.sheetId,
    sheetName: firstTile.sheetName,
    freezeRows: firstTile.freezeRows,
    freezeCols: firstTile.freezeCols,
    cells: [...cells.values()].toSorted(sortViewportCells),
    rowAxisEntries: [...rowAxisEntries.values()].toSorted((left, right) => left.index - right.index),
    columnAxisEntries: [...columnAxisEntries.values()].toSorted((left, right) => left.index - right.index),
    styles: [...styles.values()],
  }
}

export class WorkerViewportTileStore {
  private readonly residency = new TileResidencyV3<CachedViewportTile>()
  private readonly tileKeysBySheetId = new Map<number, Set<number>>()
  private readonly sheetIdsByName = new Map<string, number>()

  constructor(
    private readonly maxCachedTiles = DEFAULT_MAX_CACHED_TILES,
    private readonly maxCachedTileBytes = DEFAULT_MAX_CACHED_TILE_BYTES,
  ) {}

  reset(): void {
    this.residency.clear()
    this.tileKeysBySheetId.clear()
    this.sheetIdsByName.clear()
  }

  invalidateSheet(sheetId: number): void {
    const keys = this.tileKeysBySheetId.get(sheetId)
    if (keys) {
      keys.forEach((key) => {
        this.residency.delete(key)
      })
    }
    this.tileKeysBySheetId.delete(sheetId)
    for (const [sheetName, mappedSheetId] of this.sheetIdsByName) {
      if (mappedSheetId === sheetId) {
        this.sheetIdsByName.delete(sheetName)
      }
    }
  }

  hasTile(sheetId: number, bounds: ViewportBounds): boolean {
    return this.residency.getExact(buildTileKey(sheetId, bounds)) !== null
  }

  readViewport(input: {
    readonly localStore: Pick<WorkbookLocalStore, 'readViewportProjection'>
    readonly sheetName: string
    readonly viewport: ViewportBounds
  }): WorkbookLocalViewportBase | null {
    const tiles = listViewportTileBounds(input.viewport)
    const resolvedTiles: WorkbookLocalViewportBase[] = []

    for (const bounds of tiles) {
      let cachedTile: CachedViewportTile | undefined
      const knownSheetId = this.sheetIdsByName.get(input.sheetName)
      if (knownSheetId !== undefined) {
        cachedTile = this.residency.getExact(buildTileKey(knownSheetId, bounds))?.packet ?? undefined
      }

      if (!cachedTile) {
        const viewport = input.localStore.readViewportProjection(input.sheetName, bounds)
        if (!viewport) {
          return null
        }
        const sheetId = this.resolveViewportSheetId(input.sheetName, viewport)
        if (sheetId === null) {
          cachedTile = {
            bounds,
            key: -1,
            sheetId: viewport.sheetId,
            sheetName: viewport.sheetName,
            viewport,
          }
        } else {
          this.updateSheetNameMapping(input.sheetName, viewport.sheetName, sheetId)
          cachedTile = this.storeTile(bounds, viewport, sheetId)
        }
      }

      resolvedTiles.push(cachedTile.viewport)
    }

    if (resolvedTiles.length === 0) {
      return null
    }
    return mergeViewportTiles(normalizeViewportBounds(input.viewport), resolvedTiles)
  }

  private storeTile(bounds: ViewportBounds, viewport: WorkbookLocalViewportBase, sheetId: number): CachedViewportTile {
    const key = buildTileKey(sheetId, bounds)
    const cached: CachedViewportTile = {
      key,
      sheetId,
      sheetName: viewport.sheetName,
      bounds,
      viewport,
    }
    this.residency.upsert({
      axisSeqX: 0,
      axisSeqY: 0,
      byteSizeCpu: estimateViewportBytes(viewport),
      byteSizeGpu: 0,
      colTile: Math.floor(bounds.colStart / VIEWPORT_TILE_COLUMN_COUNT),
      dprBucket: 0,
      freezeSeq: 0,
      key,
      packet: cached,
      rectSeq: 0,
      rowTile: Math.floor(bounds.rowStart / VIEWPORT_TILE_ROW_COUNT),
      sheetOrdinal: sheetId,
      state: 'ready',
      styleSeq: 0,
      textSeq: 0,
      valueSeq: 0,
    })
    const sheetTileKeys = this.tileKeysBySheetId.get(sheetId) ?? new Set<number>()
    sheetTileKeys.add(key)
    this.tileKeysBySheetId.set(sheetId, sheetTileKeys)
    this.evictLeastRecentlyUsedTiles()
    return cached
  }

  private resolveViewportSheetId(requestedSheetName: string, viewport: WorkbookLocalViewportBase): number | null {
    const sheetId = Number(viewport.sheetId)
    if (Number.isSafeInteger(sheetId) && sheetId >= 0 && sheetId <= MAX_TILE_SHEET_ORDINAL) {
      return sheetId
    }
    const knownSheetId = this.sheetIdsByName.get(viewport.sheetName) ?? this.sheetIdsByName.get(requestedSheetName)
    if (knownSheetId !== undefined) {
      return knownSheetId
    }
    return null
  }

  private updateSheetNameMapping(requestedSheetName: string, actualSheetName: string, sheetId: number): void {
    const requestedSheetId = this.sheetIdsByName.get(requestedSheetName)
    if (requestedSheetId !== undefined && requestedSheetId !== sheetId) {
      this.invalidateSheet(requestedSheetId)
    }
    const actualSheetId = this.sheetIdsByName.get(actualSheetName)
    if (actualSheetId !== undefined && actualSheetId !== sheetId) {
      this.invalidateSheet(actualSheetId)
    }
    this.sheetIdsByName.set(actualSheetName, sheetId)
    this.sheetIdsByName.set(requestedSheetName, sheetId)
  }

  private evictLeastRecentlyUsedTiles(): void {
    const removeTileIndex = (entry: TileEntryV3<CachedViewportTile>) => {
      const tile = entry.packet
      if (tile) {
        this.removeTileKeyFromSheet(tile.sheetId, entry.key)
      }
    }
    this.residency.evictToBudgets({
      maxCpuBytes: this.maxCachedTileBytes,
      maxGpuBytes: Number.POSITIVE_INFINITY,
      onEvict: removeTileIndex,
    })
    this.residency.evictToSize(this.maxCachedTiles, removeTileIndex)
  }

  private removeTileKeyFromSheet(sheetId: number, key: number): void {
    const sheetKeys = this.tileKeysBySheetId.get(sheetId)
    sheetKeys?.delete(key)
    if (!sheetKeys || sheetKeys.size === 0) {
      this.tileKeysBySheetId.delete(sheetId)
    }
  }
}

function estimateViewportBytes(viewport: WorkbookLocalViewportBase): number {
  let bytes = viewport.cells.length * 64
  bytes += (viewport.rowAxisEntries.length + viewport.columnAxisEntries.length) * 32
  bytes += viewport.styles.length * 128
  for (const cell of viewport.cells) {
    bytes += cell.snapshot.address.length * 2
    if (typeof cell.snapshot.input === 'string') {
      bytes += cell.snapshot.input.length * 2
    }
  }
  return bytes
}
