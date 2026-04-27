import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { resolveGridRectTileSignatureV3, resolveGridTextTileSignatureV3 } from '../renderer-v3/typegpu-tile-buffer-pool.js'

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  const version = {
    axisX: 1,
    axisY: 1,
    freeze: 0,
    styles: 1,
    text: 1,
    values: 1,
    ...overrides.version,
  }
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 1,
    rectInstances: new Float32Array([0, 0, 104, 22, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200, 100]),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 101,
    version,
    ...overrides,
  }
}

function rectSignature(tile: GridRenderTile): string {
  return resolveGridRectTileSignatureV3({ tile })
}

describe('typegpu v3 resource cache signatures', () => {
  test('keeps equivalent V3 text tiles stable without relying on object identity', () => {
    const first = createTile({
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 22,
          clipWidth: 104,
          clipX: 0,
          clipY: 0,
          color: '#111111',
          font: '400 11px sans-serif',
          fontSize: 11,
          height: 22,
          strike: false,
          text: 'A1',
          underline: false,
          width: 104,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    })
    const second = createTile({
      textCount: first.textCount,
      textRuns: first.textRuns.map((run) => ({ ...run })),
      version: { ...first.version },
    })
    const changed = createTile({
      textCount: first.textCount,
      textRuns: first.textRuns.map((run) => ({ ...run, text: 'A2' })),
      version: { ...first.version, text: 2 },
    })

    expect(resolveGridTextTileSignatureV3(first)).toBe(resolveGridTextTileSignatureV3(second))
    expect(resolveGridTextTileSignatureV3(changed)).not.toBe(resolveGridTextTileSignatureV3(first))
  })

  test('keeps V3 resource signatures stable across camera sequence churn', () => {
    const base = createTile()
    const newerCameraWithSameContent = createTile({
      lastCameraSeq: 22,
      version: { ...base.version },
    })

    expect(rectSignature(newerCameraWithSameContent)).toBe(rectSignature(base))
    expect(resolveGridTextTileSignatureV3(newerCameraWithSameContent)).toBe(resolveGridTextTileSignatureV3(base))
  })

  test('tracks V3 tile revisions and decoration counts in resource signatures', () => {
    const base = createTile()
    const changedValues = createTile({ version: { ...base.version, values: 2 } })
    const changedBatch = createTile({ lastBatchId: 2, version: { ...base.version } })

    expect(rectSignature(changedValues)).not.toBe(rectSignature(base))
    expect(resolveGridTextTileSignatureV3(changedBatch)).not.toBe(resolveGridTextTileSignatureV3(base))
    expect(
      resolveGridRectTileSignatureV3({
        decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
        tile: base,
      }),
    ).not.toBe(rectSignature(base))
  })
})
