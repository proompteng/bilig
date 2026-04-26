import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { packTileKey53, type TileKey53 } from './tile-key.js'

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

export class DirtyTileIndexV3 {
  private readonly masks = new Map<TileKey53, number>()

  markCellRange(input: {
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
    readonly mask: number
  }): void {
    const rowTileStart = Math.floor(input.rowStart / VIEWPORT_TILE_ROW_COUNT)
    const rowTileEnd = Math.floor(input.rowEnd / VIEWPORT_TILE_ROW_COUNT)
    const colTileStart = Math.floor(input.colStart / VIEWPORT_TILE_COLUMN_COUNT)
    const colTileEnd = Math.floor(input.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
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

  markTile(key: TileKey53, mask: number): void {
    this.masks.set(key, (this.masks.get(key) ?? 0) | mask)
  }

  getMask(key: TileKey53): number {
    return this.masks.get(key) ?? 0
  }

  consumeVisible(visibleKeys: Iterable<TileKey53>): TileKey53[] {
    const dirty: number[] = []
    for (const key of visibleKeys) {
      if (!this.masks.has(key)) {
        continue
      }
      dirty.push(key)
      this.masks.delete(key)
    }
    return dirty
  }

  peekWarm(warmKeys: Iterable<TileKey53>): TileKey53[] {
    const dirty: number[] = []
    for (const key of warmKeys) {
      if (this.masks.has(key)) {
        dirty.push(key)
      }
    }
    return dirty
  }

  clear(): void {
    this.masks.clear()
  }
}
