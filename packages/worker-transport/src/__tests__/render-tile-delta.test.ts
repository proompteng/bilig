import { describe, expect, it } from 'vitest'

import { decodeRenderTileDeltaBatch, encodeRenderTileDeltaBatch, type RenderTileDeltaBatch } from '../index.js'

function createBatch(): RenderTileDeltaBatch {
  const version = {
    axisX: 1,
    axisY: 2,
    values: 3,
    styles: 4,
    text: 5,
    freeze: 6,
  }
  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: 7,
    batchId: 11,
    cameraSeq: 13,
    mutations: [
      {
        kind: 'tileReplace',
        tileId: 17,
        coord: {
          sheetId: 7,
          paneKind: 'body',
          rowTile: 2,
          colTile: 3,
          dprBucket: 2,
        },
        version,
        bounds: {
          rowStart: 64,
          rowEnd: 95,
          colStart: 384,
          colEnd: 511,
        },
        rectInstances: Float32Array.from([1, 2, 3, 4]),
        rectCount: 1,
        textMetrics: Float32Array.from([5, 6, 7, 8]),
        glyphRefs: Uint32Array.from([10, 11, 12]),
        textRuns: [
          {
            text: 'Revenue',
            x: 12,
            y: 18,
            width: 88,
            height: 20,
            clipX: 10,
            clipY: 16,
            clipWidth: 90,
            clipHeight: 24,
            font: '400 11px sans-serif',
            fontSize: 11,
            color: '#202124',
            underline: true,
            strike: false,
          },
        ],
        textCount: 1,
        dirty: {
          rectSpans: [{ offset: 0, length: 4 }],
          textSpans: [{ offset: 0, length: 4 }],
          glyphSpans: [{ offset: 0, length: 3 }],
        },
      },
      {
        kind: 'cellRuns',
        tileId: 17,
        version: { ...version, values: 4 },
        runs: [
          {
            row: 66,
            colStart: 386,
            colEnd: 390,
            rectSpan: { offset: 20, length: 8 },
            textSpan: { offset: 10, length: 5 },
            glyphSpan: { offset: 7, length: 4 },
          },
        ],
      },
      {
        kind: 'axis',
        axis: 'col',
        changedStart: 4,
        changedEnd: 7,
        axisVersion: 9,
      },
      {
        kind: 'freeze',
        freezeRows: 1,
        freezeCols: 2,
        freezeVersion: 3,
      },
      {
        kind: 'invalidate',
        tileId: 17,
        reason: 'axis-version-mismatch',
      },
      {
        kind: 'overlay',
        overlayRevision: 19,
        dirtyBounds: {
          rowStart: 1,
          rowEnd: 5,
          colStart: 2,
          colEnd: 8,
        },
      },
    ],
  }
}

describe('render tile delta codec', () => {
  it('round-trips renderer-native tile delta batches through the binary codec', () => {
    const batch = createBatch()
    const bytes = encodeRenderTileDeltaBatch(batch)
    const decoded = decodeRenderTileDeltaBatch(bytes)

    expect(bytes[0]).not.toBe('{'.charCodeAt(0))
    expect(decoded).toEqual(batch)
  })
})
