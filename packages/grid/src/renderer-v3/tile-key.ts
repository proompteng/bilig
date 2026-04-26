import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'

export type TileKey53 = number

export interface TileKeyFields {
  readonly sheetOrdinal: number
  readonly rowTile: number
  readonly colTile: number
  readonly dprBucket: number
}

const DPR_BITS = 4
const COL_TILE_BITS = 7
const ROW_TILE_BITS = 15
const SHEET_ORDINAL_BITS = 27

const DPR_FACTOR = 2 ** DPR_BITS
const COL_TILE_FACTOR = 2 ** COL_TILE_BITS
const ROW_TILE_FACTOR = 2 ** ROW_TILE_BITS
const SHEET_ORDINAL_FACTOR = 2 ** SHEET_ORDINAL_BITS

export const MAX_TILE_DPR_BUCKET = DPR_FACTOR - 1
export const MAX_TILE_COLUMN_INDEX = Math.ceil(MAX_COLS / VIEWPORT_TILE_COLUMN_COUNT) - 1
export const MAX_TILE_ROW_INDEX = Math.ceil(MAX_ROWS / VIEWPORT_TILE_ROW_COUNT) - 1
export const MAX_TILE_SHEET_ORDINAL = SHEET_ORDINAL_FACTOR - 1

export function packTileKey53(fields: TileKeyFields): TileKey53 {
  assertIntegerInRange('sheetOrdinal', fields.sheetOrdinal, MAX_TILE_SHEET_ORDINAL)
  assertIntegerInRange('rowTile', fields.rowTile, MAX_TILE_ROW_INDEX)
  assertIntegerInRange('colTile', fields.colTile, MAX_TILE_COLUMN_INDEX)
  assertIntegerInRange('dprBucket', fields.dprBucket, MAX_TILE_DPR_BUCKET)
  return ((fields.sheetOrdinal * ROW_TILE_FACTOR + fields.rowTile) * COL_TILE_FACTOR + fields.colTile) * DPR_FACTOR + fields.dprBucket
}

export function unpackTileKey53(key: TileKey53): TileKeyFields {
  assertIntegerInRange('key', key, Number.MAX_SAFE_INTEGER)
  let remaining = key
  const dprBucket = remaining % DPR_FACTOR
  remaining = Math.floor(remaining / DPR_FACTOR)
  const colTile = remaining % COL_TILE_FACTOR
  remaining = Math.floor(remaining / COL_TILE_FACTOR)
  const rowTile = remaining % ROW_TILE_FACTOR
  const sheetOrdinal = Math.floor(remaining / ROW_TILE_FACTOR)
  return {
    colTile,
    dprBucket,
    rowTile,
    sheetOrdinal,
  }
}

export function tileKeyFromCell(input: {
  readonly sheetOrdinal: number
  readonly row: number
  readonly col: number
  readonly dprBucket: number
}): TileKey53 {
  return packTileKey53({
    sheetOrdinal: input.sheetOrdinal,
    rowTile: Math.floor(input.row / VIEWPORT_TILE_ROW_COUNT),
    colTile: Math.floor(input.col / VIEWPORT_TILE_COLUMN_COUNT),
    dprBucket: input.dprBucket,
  })
}

export function tileKeysForViewport(input: {
  readonly sheetOrdinal: number
  readonly viewport: Viewport
  readonly dprBucket: number
}): TileKey53[] {
  const rowTileStart = Math.floor(input.viewport.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const rowTileEnd = Math.floor(input.viewport.rowEnd / VIEWPORT_TILE_ROW_COUNT)
  const colTileStart = Math.floor(input.viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const colTileEnd = Math.floor(input.viewport.colEnd / VIEWPORT_TILE_COLUMN_COUNT)
  const keys: number[] = []
  for (let rowTile = rowTileStart; rowTile <= rowTileEnd; rowTile += 1) {
    for (let colTile = colTileStart; colTile <= colTileEnd; colTile += 1) {
      keys.push(
        packTileKey53({
          sheetOrdinal: input.sheetOrdinal,
          rowTile,
          colTile,
          dprBucket: input.dprBucket,
        }),
      )
    }
  }
  return keys
}

export function debugTileKey(key: TileKey53): string {
  const fields = unpackTileKey53(key)
  return `s${fields.sheetOrdinal}:r${fields.rowTile}:c${fields.colTile}:d${fields.dprBucket}`
}

function assertIntegerInRange(name: string, value: number, max: number): void {
  if (!Number.isInteger(value) || value < 0 || value > max) {
    throw new RangeError(`${name} must be an integer between 0 and ${max}`)
  }
}
