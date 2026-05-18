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
  readonly seq?: number | undefined
  readonly source?: string | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal: number
  readonly dirty: WorkbookDeltaDirtyRangesLikeV3
}

export interface DirtyTileLocalSpanV3 {
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
  readonly mask: number
}

export class DirtyTileIndexV3 {
  private readonly masks = new Map<TileKey53, number>()
  private readonly spans = new Map<TileKey53, DirtyTileLocalSpanV3[]>()
  private readonly axisXMasks = new Map<TileKey53, number>()
  private readonly axisXSpans = new Map<TileKey53, DirtyTileLocalSpanV3[]>()
  private readonly axisYMasks = new Map<TileKey53, number>()
  private readonly axisYSpans = new Map<TileKey53, DirtyTileLocalSpanV3[]>()
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
    if (rowStart === 0 && rowEnd === MAX_ROWS - 1) {
      this.markAxisX({
        colEnd,
        colStart,
        dprBucket: input.dprBucket,
        mask: input.mask,
        sheetOrdinal: input.sheetOrdinal,
      })
      return
    }
    if (colStart === 0 && colEnd === MAX_COLS - 1) {
      this.markAxisY({
        dprBucket: input.dprBucket,
        mask: input.mask,
        rowEnd,
        rowStart,
        sheetOrdinal: input.sheetOrdinal,
      })
      return
    }
    const rowTileStart = Math.floor(rowStart / VIEWPORT_TILE_ROW_COUNT)
    const rowTileEnd = Math.floor(rowEnd / VIEWPORT_TILE_ROW_COUNT)
    const colTileStart = Math.floor(colStart / VIEWPORT_TILE_COLUMN_COUNT)
    const colTileEnd = Math.floor(colEnd / VIEWPORT_TILE_COLUMN_COUNT)
    for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
      for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
        const tileRowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
        const tileRowEnd = Math.min(MAX_ROWS - 1, tileRowStart + VIEWPORT_TILE_ROW_COUNT - 1)
        const tileColStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
        const tileColEnd = Math.min(MAX_COLS - 1, tileColStart + VIEWPORT_TILE_COLUMN_COUNT - 1)
        const key = packTileKey53({
          sheetOrdinal: input.sheetOrdinal,
          rowTile,
          colTile,
          dprBucket: input.dprBucket,
        })
        this.markTile(key, input.mask)
        this.appendSpan(key, {
          colEnd: Math.min(colEnd, tileColEnd) - tileColStart,
          colStart: Math.max(colStart, tileColStart) - tileColStart,
          mask: input.mask,
          rowEnd: Math.min(rowEnd, tileRowEnd) - tileRowStart,
          rowStart: Math.max(rowStart, tileRowStart) - tileRowStart,
        })
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
      const mask = input.mask | DirtyMaskV3.AxisX
      const tileColStart = colTile * VIEWPORT_TILE_COLUMN_COUNT
      const tileColEnd = Math.min(MAX_COLS - 1, tileColStart + VIEWPORT_TILE_COLUMN_COUNT - 1)
      this.axisXMasks.set(key, (this.axisXMasks.get(key) ?? 0) | mask)
      const localColStart = Math.max(colStart, tileColStart) - tileColStart
      const localColEnd = (input.mask & DirtyMaskV3.AxisX) !== 0 ? tileColEnd - tileColStart : Math.min(colEnd, tileColEnd) - tileColStart
      this.appendSpan(
        key,
        {
          colEnd: localColEnd,
          colStart: localColStart,
          mask,
          rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
          rowStart: 0,
        },
        this.axisXSpans,
      )
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
      const mask = input.mask | DirtyMaskV3.AxisY
      const tileRowStart = rowTile * VIEWPORT_TILE_ROW_COUNT
      const tileRowEnd = Math.min(MAX_ROWS - 1, tileRowStart + VIEWPORT_TILE_ROW_COUNT - 1)
      this.axisYMasks.set(key, (this.axisYMasks.get(key) ?? 0) | mask)
      const localRowStart = Math.max(rowStart, tileRowStart) - tileRowStart
      const localRowEnd = (input.mask & DirtyMaskV3.AxisY) !== 0 ? tileRowEnd - tileRowStart : Math.min(rowEnd, tileRowEnd) - tileRowStart
      this.appendSpan(
        key,
        {
          colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
          colStart: 0,
          mask,
          rowEnd: localRowEnd,
          rowStart: localRowStart,
        },
        this.axisYSpans,
      )
    }
  }

  applyWorkbookDelta(batch: WorkbookDeltaBatchLikeV3, options: { readonly dprBucket: number }): void {
    markWorkbookDeltaDirtyTilesV3(this, batch, options)
  }

  markTile(key: TileKey53, mask: number): void {
    this.masks.set(key, (this.masks.get(key) ?? 0) | mask)
  }

  getMask(key: TileKey53): number {
    return this.resolveMask(key, false)
  }

  getUnconsumedMask(key: TileKey53): number {
    return this.resolveMask(key, true)
  }

  getSpans(key: TileKey53): readonly DirtyTileLocalSpanV3[] {
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
    const consumedAxisMask = this.consumedAxisMasks.get(key) ?? 0
    return [
      ...(this.spans.get(key) ?? []),
      ...this.filterUnconsumedAxisSpans(this.axisXSpans.get(axisXKey), consumedAxisMask),
      ...this.filterUnconsumedAxisSpans(this.axisYSpans.get(axisYKey), consumedAxisMask),
    ]
  }

  private resolveMask(key: TileKey53, unconsumedOnly: boolean): number {
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
    const mask = (this.masks.get(key) ?? 0) | (this.axisXMasks.get(axisXKey) ?? 0) | (this.axisYMasks.get(axisYKey) ?? 0)
    if (!unconsumedOnly) {
      return mask
    }
    const consumedAxisMask = this.consumedAxisMasks.get(key) ?? 0
    const exactMask = this.masks.get(key) ?? 0
    return exactMask | (mask & ~consumedAxisMask)
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
      this.spans.delete(key)
      this.consumedAxisMasks.set(key, consumedAxisMask | mask)
    }
    return dirty
  }

  peekWarm(warmKeys: Iterable<TileKey53>): TileKey53[] {
    const dirty: number[] = []
    for (const key of warmKeys) {
      if (this.getUnconsumedMask(key) !== 0) {
        dirty.push(key)
      }
    }
    return dirty
  }

  clear(): void {
    this.masks.clear()
    this.spans.clear()
    this.axisXMasks.clear()
    this.axisXSpans.clear()
    this.axisYMasks.clear()
    this.axisYSpans.clear()
    this.consumedAxisMasks.clear()
  }

  private appendSpan(key: TileKey53, span: DirtyTileLocalSpanV3, spansByTile: Map<TileKey53, DirtyTileLocalSpanV3[]> = this.spans): void {
    const spans = spansByTile.get(key) ?? []
    spans.push(span)
    spansByTile.set(key, spans)
  }

  private filterUnconsumedAxisSpans(
    spans: readonly DirtyTileLocalSpanV3[] | undefined,
    consumedAxisMask: number,
  ): readonly DirtyTileLocalSpanV3[] {
    if (!spans || spans.length === 0) {
      return []
    }
    return spans.filter((span) => (span.mask & ~consumedAxisMask) !== 0)
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
