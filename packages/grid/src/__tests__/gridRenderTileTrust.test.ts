import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { hasCompleteRenderTileGrid, tileSelectedTextNeedsLocalRefresh } from '../runtime/gridRenderTileTrust.js'

describe('grid render tile trust predicates', () => {
  it('rejects remote grid payloads when rect buffers are shorter than rectCount', () => {
    const tile = createTile({ rectCount: 4, rectInstances: createGridBorderRectInstances(2) })

    expect(hasCompleteRenderTileGrid(tile)).toBe(false)
  })

  it('accepts a complete set of visible border rects for the tile bounds', () => {
    const tile = createTile({ rectCount: 6, rectInstances: createGridBorderRectInstances(6) })

    expect(hasCompleteRenderTileGrid(tile)).toBe(true)
  })

  it('requires local refresh when the selected remote text disagrees with the authoritative snapshot', () => {
    const tile = createTile({
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 0,
          clipY: 0,
          color: '#111827',
          col: 1,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 2,
          strike: false,
          text: 'old value',
          underline: false,
          width: 100,
          x: 0,
          y: 0,
        },
      ],
    })

    expect(tileSelectedTextNeedsLocalRefresh(tile, [1, 2], stringSnapshot('B3', 'new value'))).toBe(true)
    expect(tileSelectedTextNeedsLocalRefresh(tile, [1, 2], stringSnapshot('B3', 'old value'))).toBe(false)
  })

  it('requires local refresh when an authoritative empty snapshot clears stale selected text', () => {
    const tile = createTile({
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 0,
          clipY: 0,
          color: '#111827',
          col: 1,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 2,
          strike: false,
          text: 'stale value',
          underline: false,
          width: 100,
          x: 0,
          y: 0,
        },
      ],
    })

    expect(tileSelectedTextNeedsLocalRefresh(tile, [1, 2], emptySnapshot('B3'))).toBe(true)
  })

  it('keeps remote selected text when the selected snapshot is only an unhydrated empty placeholder', () => {
    const tile = createTile({
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 0,
          clipY: 0,
          color: '#111827',
          col: 1,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 2,
          strike: false,
          text: 'remote value',
          underline: false,
          width: 100,
          x: 0,
          y: 0,
        },
      ],
    })

    expect(tileSelectedTextNeedsLocalRefresh(tile, [1, 2], defaultPlaceholderEmptySnapshot('B3'))).toBe(false)
    expect(tileSelectedTextNeedsLocalRefresh(tile, [1, 2], null)).toBe(false)
  })
})

function createGridBorderRectInstances(rectCount: number): Float32Array {
  const rectInstances = new Float32Array(rectCount * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    rectInstances[offset + 2] = index % 2 === 0 ? 100 : 1
    rectInstances[offset + 3] = index % 2 === 0 ? 1 : 20
    rectInstances[offset + 11] = 1
    rectInstances[offset + 13] = 1
  }
  return rectInstances
}

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  return {
    bounds: { colEnd: 2, colStart: 0, rowEnd: 2, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
      sheetOrdinal: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 6,
    rectInstances: createGridBorderRectInstances(6),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 1,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
    ...overrides,
  }
}

function emptySnapshot(address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value: { tag: ValueTag.Empty },
    version: 1,
  }
}

function defaultPlaceholderEmptySnapshot(address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.Empty },
    version: 0,
  }
}

function stringSnapshot(address: string, value: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: value,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.String, stringId: 0, value },
    version: 1,
  }
}
