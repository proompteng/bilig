import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import { MAX_TILE_COLUMN_INDEX, MAX_TILE_ROW_INDEX, packTileKey53, unpackTileKey53, type TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

export interface GridRenderTileInterestRuntimeInput {
  readonly dprBucket: number
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly visibleViewport: Viewport
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

export interface RenderTileDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly freezeCols: number
  readonly freezeRows: number
  readonly renderTileSource: unknown
  readonly residentViewport: Viewport | undefined
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
  readonly sheetName: string
  readonly viewport: Viewport
}

export interface WorkbookDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly renderTileSource: unknown
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
}

export interface LocalInvalidationConnectionIdentity {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

export interface RetainedFixedRenderTileDataPanesCompatibility {
  readonly dprBucket: number
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly sceneRevision: number
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly visibleViewport: Viewport
}

export interface GridRenderTilePaneRetainedCompatibilityInput extends RetainedFixedRenderTileDataPanesCompatibility {}

interface RenderTileSheetIdentity {
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

export function isResidentGridRenderTile(value: unknown): value is GridRenderTile {
  if (!value || typeof value !== 'object') {
    return false
  }
  const candidate = value as Partial<GridRenderTile>
  return (
    typeof candidate.tileId === 'number' &&
    candidate.coord !== undefined &&
    typeof candidate.coord.sheetId === 'number' &&
    typeof candidate.coord.sheetOrdinal === 'number' &&
    candidate.bounds !== undefined &&
    candidate.rectInstances instanceof Float32Array &&
    candidate.textMetrics instanceof Float32Array &&
    Array.isArray(candidate.textRuns)
  )
}

export function resolveGridRenderTileInputSheetOrdinal(input: {
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}): number {
  return input.sheetOrdinal ?? input.sheetId ?? 0
}

export function resolveRenderTileInterestTileKeys(input: GridRenderTileInterestRuntimeInput): readonly TileKey53[] {
  return resolveTileKeysForRenderTileViewports(input, resolveRenderTileInterestViewports(input))
}

export function resolveRenderTileResidentTileKeys(input: GridRenderTileInterestRuntimeInput): readonly TileKey53[] {
  return resolveTileKeysForRenderTileViewports(input, resolveRenderTileResidentViewports(input))
}

function resolveTileKeysForRenderTileViewports(
  input: GridRenderTileInterestRuntimeInput,
  viewports: readonly Viewport[],
): readonly TileKey53[] {
  const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
  const keys = new Set<number>()
  const result: number[] = []
  for (const viewport of viewports) {
    for (const key of input.gridRuntimeHost.viewportTileKeys({ dprBucket: input.dprBucket, sheetOrdinal, viewport })) {
      if (keys.has(key)) {
        continue
      }
      keys.add(key)
      result.push(key)
    }
  }
  return result
}

export function resolveRenderTileInterestViewports(input: GridRenderTileInterestRuntimeInput): readonly Viewport[] {
  return resolveDisjointRenderTileViewports(input, input.visibleViewport)
}

function resolveRenderTileResidentViewports(input: GridRenderTileInterestRuntimeInput): readonly Viewport[] {
  return resolveDisjointRenderTileViewports(input, input.residentViewport ?? input.renderTileViewport)
}

function resolveDisjointRenderTileViewports(input: GridRenderTileInterestRuntimeInput, bodyViewport: Viewport): readonly Viewport[] {
  const freezeRows = Math.max(0, input.freezeRows ?? 0)
  const freezeCols = Math.max(0, input.freezeCols ?? 0)
  const viewports: Viewport[] = []
  addRenderTileInterestViewport(viewports, bodyViewport)
  if (freezeRows > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: bodyViewport.colEnd,
      colStart: bodyViewport.colStart,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  if (freezeCols > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: bodyViewport.rowEnd,
      rowStart: bodyViewport.rowStart,
    })
  }
  if (freezeRows > 0 && freezeCols > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  return viewports
}

function addRenderTileInterestViewport(viewports: Viewport[], viewport: Viewport): void {
  if (viewport.rowEnd < viewport.rowStart || viewport.colEnd < viewport.colStart) {
    return
  }
  viewports.push(viewport)
}

export function resolveWarmRenderTileKeys(
  input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number },
  visibleTileKeys = resolveRenderTileInterestTileKeys(input),
): readonly TileKey53[] {
  const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
  if (visibleTileKeys.length === 0) {
    return []
  }
  const visibleSet = new Set(visibleTileKeys)
  const warmSet = new Set<number>()
  const warmTileKeys: number[] = []
  for (const viewport of resolveRenderTileInterestViewports(input)) {
    let minRowTile = Number.POSITIVE_INFINITY
    let maxRowTile = Number.NEGATIVE_INFINITY
    let minColTile = Number.POSITIVE_INFINITY
    let maxColTile = Number.NEGATIVE_INFINITY
    for (const key of input.gridRuntimeHost.viewportTileKeys({ dprBucket: input.dprBucket, sheetOrdinal, viewport })) {
      const fields = unpackTileKey53(key)
      minRowTile = Math.min(minRowTile, fields.rowTile)
      maxRowTile = Math.max(maxRowTile, fields.rowTile)
      minColTile = Math.min(minColTile, fields.colTile)
      maxColTile = Math.max(maxColTile, fields.colTile)
    }
    if (!Number.isFinite(minRowTile)) {
      continue
    }
    for (let rowTile = Math.max(0, minRowTile - 1); rowTile <= Math.min(MAX_TILE_ROW_INDEX, maxRowTile + 1); rowTile += 1) {
      for (let colTile = Math.max(0, minColTile - 1); colTile <= Math.min(MAX_TILE_COLUMN_INDEX, maxColTile + 1); colTile += 1) {
        const key = packTileKey53({
          colTile,
          dprBucket: input.dprBucket,
          rowTile,
          sheetOrdinal,
        })
        if (!visibleSet.has(key) && !warmSet.has(key)) {
          warmSet.add(key)
          warmTileKeys.push(key)
        }
      }
    }
  }
  return warmTileKeys
}

export function sameRenderTileDeltaConnectionIdentity(
  left: RenderTileDeltaConnectionIdentity,
  right: RenderTileDeltaConnectionIdentity,
): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.freezeCols === right.freezeCols &&
    left.freezeRows === right.freezeRows &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetName === right.sheetName &&
    left.sheetOrdinal === right.sheetOrdinal &&
    sameOptionalViewportIdentity(left.residentViewport, right.residentViewport) &&
    sameViewportIdentity(left.viewport, right.viewport)
  )
}

export function sameWorkbookDeltaConnectionIdentity(
  left: WorkbookDeltaConnectionIdentity,
  right: WorkbookDeltaConnectionIdentity,
): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetOrdinal === right.sheetOrdinal
  )
}

export function sameLocalInvalidationConnectionIdentity(
  left: LocalInvalidationConnectionIdentity,
  right: LocalInvalidationConnectionIdentity,
): boolean {
  return (
    left.engine === right.engine &&
    left.needsLocalCellInvalidation === right.needsLocalCellInvalidation &&
    left.sheetName === right.sheetName &&
    sameStringListIdentity(left.visibleAddresses, right.visibleAddresses)
  )
}

function sameStringListIdentity(left: readonly string[], right: readonly string[]): boolean {
  if (left === right) {
    return true
  }
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

export function shouldForceLocalTilesForWorkbookDelta(batch: WorkbookDeltaBatchLikeV3): boolean {
  return batch.source === 'localOptimistic' || batch.source === 'workerAuthoritative'
}

function sameViewportIdentity(left: Viewport, right: Viewport): boolean {
  return (
    left.colEnd === right.colEnd && left.colStart === right.colStart && left.rowEnd === right.rowEnd && left.rowStart === right.rowStart
  )
}

function sameOptionalViewportIdentity(left: Viewport | undefined, right: Viewport | undefined): boolean {
  if (left === undefined || right === undefined) {
    return left === right
  }
  return sameViewportIdentity(left, right)
}

export function matchesRenderTileSheetIdentity(tile: RenderTileSheetIdentity, expected: RenderTileSheetIdentity): boolean {
  if (expected.sheetId !== undefined && expected.sheetOrdinal !== undefined) {
    return tile.sheetId === expected.sheetId && tile.sheetOrdinal === expected.sheetOrdinal
  }
  if (expected.sheetId !== undefined) {
    return tile.sheetId === expected.sheetId
  }
  if (expected.sheetOrdinal !== undefined) {
    return tile.sheetOrdinal === expected.sheetOrdinal
  }
  return false
}

export function buildRetainedFixedRenderTileDataPanesCompatibility(
  input: GridRenderTilePaneRetainedCompatibilityInput,
): RetainedFixedRenderTileDataPanesCompatibility {
  return {
    dprBucket: input.dprBucket,
    freezeCols: input.freezeCols,
    freezeRows: input.freezeRows,
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    hostClientHeight: input.hostClientHeight,
    hostClientWidth: input.hostClientWidth,
    renderTileViewport: input.renderTileViewport,
    residentViewport: input.residentViewport,
    sceneRevision: input.sceneRevision,
    sheetId: input.sheetId,
    sheetOrdinal: input.sheetOrdinal,
    visibleViewport: input.visibleViewport,
  }
}

export function sameRetainedFixedRenderTileDataPanesCompatibility(
  left: RetainedFixedRenderTileDataPanesCompatibility,
  right: RetainedFixedRenderTileDataPanesCompatibility,
): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.freezeCols === right.freezeCols &&
    left.freezeRows === right.freezeRows &&
    left.frozenColumnWidth === right.frozenColumnWidth &&
    left.frozenRowHeight === right.frozenRowHeight &&
    left.hostClientHeight === right.hostClientHeight &&
    left.hostClientWidth === right.hostClientWidth &&
    left.sceneRevision === right.sceneRevision &&
    left.sheetId === right.sheetId &&
    left.sheetOrdinal === right.sheetOrdinal &&
    sameViewportIdentity(left.renderTileViewport, right.renderTileViewport) &&
    sameViewportIdentity(left.residentViewport, right.residentViewport) &&
    sameViewportIdentity(left.visibleViewport, right.visibleViewport)
  )
}

export function upsertRenderTileIntoHost(gridRuntimeHost: GridRuntimeHost, tile: GridRenderTile): void {
  gridRuntimeHost.tiles.upsertTile({
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    byteSizeCpu: estimateTileCpuBytes(tile),
    byteSizeGpu: tile.rectInstances.byteLength + tile.textMetrics.byteLength,
    colTile: tile.coord.colTile,
    dprBucket: tile.coord.dprBucket,
    freezeSeq: tile.version.freeze,
    key: tile.tileId,
    packet: tile,
    rectSeq: tile.version.styles,
    rowTile: tile.coord.rowTile,
    sheetOrdinal: tile.coord.sheetOrdinal,
    state: 'ready',
    styleSeq: tile.version.styles,
    textSeq: tile.version.text,
    valueSeq: tile.version.values,
  })
}

export function resolveRenderTileViewportUnion(tiles: readonly GridRenderTile[]): Viewport | null {
  const first = tiles[0]
  if (!first) {
    return null
  }
  let colStart = first.bounds.colStart
  let colEnd = first.bounds.colEnd
  let rowStart = first.bounds.rowStart
  let rowEnd = first.bounds.rowEnd
  for (let index = 1; index < tiles.length; index += 1) {
    const tile = tiles[index]!
    colStart = Math.min(colStart, tile.bounds.colStart)
    colEnd = Math.max(colEnd, tile.bounds.colEnd)
    rowStart = Math.min(rowStart, tile.bounds.rowStart)
    rowEnd = Math.max(rowEnd, tile.bounds.rowEnd)
  }
  return { colEnd, colStart, rowEnd, rowStart }
}

function estimateTileCpuBytes(tile: GridRenderTile): number {
  let textBytes = 0
  for (const run of tile.textRuns) {
    textBytes += run.text.length * 2
    textBytes += run.font.length * 2
    textBytes += run.color.length * 2
  }
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength + textBytes
}
