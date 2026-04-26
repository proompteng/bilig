import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { packTileKey53, unpackTileKey53, type TileKey53 } from './tile-key.js'

export enum DirtyMaskV3 {
  Value = 1 << 0,
  Style = 1 << 1,
  Text = 1 << 2,
  Rect = 1 << 3,
  Border = 1 << 4,
  AxisX = 1 << 5,
  AxisY = 1 << 6,
  Freeze = 1 << 7,
  Presence = 1 << 8,
}

export interface WorkbookDeltaDirtyRangesLikeV3 {
  readonly cellRanges: Uint32Array
  readonly axisX: Uint32Array
  readonly axisY: Uint32Array
}

export interface WorkbookDeltaBatchLikeV3 {
  readonly sheetOrdinal: number
  readonly dirty: WorkbookDeltaDirtyRangesLikeV3
}

export class DirtyTileIndexV3 {
  private readonly masks = new Map<TileKey53, number>()
  private readonly axisXMasks = new Map<TileKey53, number>()
  private readonly axisYMasks = new Map<TileKey53, number>()
  private readonly consumedAxisMasks = new Map<TileKey53, number>()

  markCellRange(input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
    readonly mask: number
  }): void {
    const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.rowStart))
    const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, input.rowEnd))
    const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.colStart))
    const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, input.colEnd))
    const rowTileStart = Math.floor(rowStart / VIEWPORT_TILE_ROW_COUNT)
    const rowTileEnd = Math.floor(rowEnd / VIEWPORT_TILE_ROW_COUNT)
    const colTileStart = Math.floor(colStart / VIEWPORT_TILE_COLUMN_COUNT)
    const colTileEnd = Math.floor(colEnd / VIEWPORT_TILE_COLUMN_COUNT)
    for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
      for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
        this.markTile(
          packTileKey53({
            sheetOrdinal: input.sheetOrdinal,
            rowTile,
            colTile,
            dprBucket: input.dprBucket,
          }),
          input.mask,
        )
      }
    }
  }

  markAxisX(input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly colStart: number
    readonly colEnd: number
    readonly mask: number
  }): void {
    this.consumedAxisMasks.clear()
    const colStart = Math.max(0, Math.min(MAX_COLS - 1, input.colStart))
    const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, input.colEnd))
    const colTileStart = Math.floor(colStart / VIEWPORT_TILE_COLUMN_COUNT)
    const colTileEnd = Math.floor(colEnd / VIEWPORT_TILE_COLUMN_COUNT)
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      const key = packTileKey53({
        colTile,
        dprBucket: input.dprBucket,
        rowTile: 0,
        sheetOrdinal: input.sheetOrdinal,
      })
      this.axisXMasks.set(key, (this.axisXMasks.get(key) ?? 0) | input.mask | DirtyMaskV3.AxisX)
    }
  }

  markAxisY(input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly mask: number
  }): void {
    this.consumedAxisMasks.clear()
    const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, input.rowStart))
    const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, input.rowEnd))
    const rowTileStart = Math.floor(rowStart / VIEWPORT_TILE_ROW_COUNT)
    const rowTileEnd = Math.floor(rowEnd / VIEWPORT_TILE_ROW_COUNT)
    for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
      const key = packTileKey53({
        colTile: 0,
        dprBucket: input.dprBucket,
        rowTile,
        sheetOrdinal: input.sheetOrdinal,
      })
      this.axisYMasks.set(key, (this.axisYMasks.get(key) ?? 0) | input.mask | DirtyMaskV3.AxisY)
    }
  }

  applyWorkbookDelta(batch: WorkbookDeltaBatchLikeV3, options: { readonly dprBucket: number }): void {
    markWorkbookDeltaDirtyTilesV3(this, batch, options)
  }

  markTile(key: TileKey53, mask: number): void {
    this.masks.set(key, (this.masks.get(key) ?? 0) | mask)
  }

  getMask(key: TileKey53): number {
    const fields = unpackTileKey53(key)
    const axisXKey = packTileKey53({
      colTile: fields.colTile,
      dprBucket: fields.dprBucket,
      rowTile: 0,
      sheetOrdinal: fields.sheetOrdinal,
    })
    const axisYKey = packTileKey53({
      colTile: 0,
      dprBucket: fields.dprBucket,
      rowTile: fields.rowTile,
      sheetOrdinal: fields.sheetOrdinal,
    })
    return (this.masks.get(key) ?? 0) | (this.axisXMasks.get(axisXKey) ?? 0) | (this.axisYMasks.get(axisYKey) ?? 0)
  }

  consumeVisible(visibleKeys: Iterable<TileKey53>): TileKey53[] {
    const dirty: number[] = []
    for (const key of visibleKeys) {
      const mask = this.getMask(key)
      const consumedAxisMask = this.consumedAxisMasks.get(key) ?? 0
      const exactMask = this.masks.get(key) ?? 0
      const nextMask = exactMask | (mask & ~consumedAxisMask)
      if (nextMask === 0) {
        continue
      }
      dirty.push(key)
      this.masks.delete(key)
      this.consumedAxisMasks.set(
        key,
        consumedAxisMask | (mask & (DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Rect | DirtyMaskV3.Text)),
      )
    }
    return dirty
  }

  peekWarm(warmKeys: Iterable<TileKey53>): TileKey53[] {
    const dirty: number[] = []
    for (const key of warmKeys) {
      if (this.getMask(key) !== 0) {
        dirty.push(key)
      }
    }
    return dirty
  }

  clear(): void {
    this.masks.clear()
    this.axisXMasks.clear()
    this.axisYMasks.clear()
    this.consumedAxisMasks.clear()
  }
}

export function markWorkbookDeltaDirtyTilesV3(
  index: DirtyTileIndexV3,
  batch: WorkbookDeltaBatchLikeV3,
  options: { readonly dprBucket: number },
): void {
  forEachDirtyCellRange(batch.dirty.cellRanges, (rowStart, rowEnd, colStart, colEnd, mask) => {
    index.markCellRange({
      colEnd,
      colStart,
      dprBucket: options.dprBucket,
      mask,
      rowEnd,
      rowStart,
      sheetOrdinal: batch.sheetOrdinal,
    })
  })
  forEachDirtyAxisRange(batch.dirty.axisX, (colStart, colEnd, mask) => {
    index.markAxisX({
      colEnd,
      colStart,
      dprBucket: options.dprBucket,
      mask,
      sheetOrdinal: batch.sheetOrdinal,
    })
  })
  forEachDirtyAxisRange(batch.dirty.axisY, (rowStart, rowEnd, mask) => {
    index.markAxisY({
      dprBucket: options.dprBucket,
      mask,
      rowEnd,
      rowStart,
      sheetOrdinal: batch.sheetOrdinal,
    })
  })
}

function forEachDirtyCellRange(
  ranges: Uint32Array,
  callback: (rowStart: number, rowEnd: number, colStart: number, colEnd: number, mask: number) => void,
): void {
  for (let offset = 0; offset + 4 < ranges.length; offset += 5) {
    callback(ranges[offset] ?? 0, ranges[offset + 1] ?? 0, ranges[offset + 2] ?? 0, ranges[offset + 3] ?? 0, ranges[offset + 4] ?? 0)
  }
}

function forEachDirtyAxisRange(ranges: Uint32Array, callback: (start: number, end: number, mask: number) => void): void {
  for (let offset = 0; offset + 2 < ranges.length; offset += 3) {
    callback(ranges[offset] ?? 0, ranges[offset + 1] ?? 0, ranges[offset + 2] ?? 0)
  }
}
