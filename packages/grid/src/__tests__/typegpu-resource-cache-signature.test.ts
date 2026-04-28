import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'
import {
  resolveGridRectTileSignatureV3,
  resolveGridTextTileSignatureV3,
  shouldSyncGridRectTileResourceV3,
  shouldSyncGridTextTileResourceV3,
  type TypeGpuTileContentResourceEntryV3,
} from '../renderer-v3/typegpu-tile-buffer-pool.js'
import { writeTypeGpuVertexBufferSubrange } from '../renderer-v3/typegpu-primitives.js'

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

function contentEntry(overrides: Partial<TypeGpuTileContentResourceEntryV3> = {}): TypeGpuTileContentResourceEntryV3 {
  return {
    decorationRects: null,
    rectCount: 1,
    rectHandle: null,
    rectSignature: 'previous-rect',
    textCount: 1,
    textGlyphIds: null,
    textGlyphPageIds: null,
    textHandle: null,
    textRunGlyphIds: null,
    textRunCount: 1,
    textRunPayloads: null,
    textRunQuadSpans: null,
    textSignature: 'previous-text',
    ...overrides,
  }
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

  test('compares V3 text run counts separately from GPU quad counts', () => {
    const tile = createTile({
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
          text: 'AB',
          underline: false,
          width: 104,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    })

    expect(
      shouldSyncGridTextTileResourceV3({
        content: contentEntry({
          textCount: 2,
          textRunCount: 1,
          textSignature: resolveGridTextTileSignatureV3(tile),
        }),
        textSignature: resolveGridTextTileSignatureV3(tile),
        tile,
      }),
    ).toBe(false)
  })

  test('resyncs V3 text resources when logical text run count changes', () => {
    const tile = createTile({
      textCount: 2,
      textRuns: [createTextRun({ text: 'A' }), createTextRun({ text: 'B' })],
    })

    expect(
      shouldSyncGridTextTileResourceV3({
        content: contentEntry({
          textCount: 2,
          textRunCount: 1,
          textSignature: 'previous-text',
        }),
        textSignature: resolveGridTextTileSignatureV3(tile),
        tile,
      }),
    ).toBe(true)
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

  test('uses dirty masks to skip unrelated rect uploads for plain text updates', () => {
    const base = createTile({ rectCount: 0, rectInstances: new Float32Array() })
    const textOnlyUpdate = createTile({
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      rectCount: 0,
      rectInstances: new Float32Array(),
      version: { ...base.version, text: 2, values: 2 },
    })

    expect(
      shouldSyncGridTextTileResourceV3({
        content: contentEntry({ textSignature: resolveGridTextTileSignatureV3(base) }),
        textSignature: resolveGridTextTileSignatureV3(textOnlyUpdate),
        tile: textOnlyUpdate,
      }),
    ).toBe(true)
    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({ rectCount: 0, rectSignature: rectSignature(base) }),
        rectSignature: rectSignature(textOnlyUpdate),
        tile: textOnlyUpdate,
      }),
    ).toBe(false)
  })

  test('keeps rect uploads for decorated text updates and decoration removal', () => {
    const decorated = createTile({
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      rectCount: 0,
      rectInstances: new Float32Array(),
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
          underline: true,
          width: 104,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
      version: { ...createTile().version, text: 2, values: 2 },
    })
    const plainAfterDecoration = createTile({
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      rectCount: 0,
      rectInstances: new Float32Array(),
      version: { ...decorated.version, text: 3, values: 3 },
    })

    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({ rectCount: 0, rectSignature: 'old-rect' }),
        rectSignature: 'new-rect',
        tile: decorated,
      }),
    ).toBe(true)
    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({
          decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
          rectCount: 0,
          rectSignature: 'old-rect',
        }),
        rectSignature: 'new-rect',
        tile: plainAfterDecoration,
      }),
    ).toBe(true)
  })

  test('writes V3 vertex buffer subranges with byte offsets instead of full payload bytes', () => {
    const writes: { readonly bytes: number; readonly startOffset?: number; readonly endOffset?: number }[] = []
    const buffer = {
      write(source: ArrayBuffer, options?: { readonly startOffset?: number; readonly endOffset?: number }) {
        writes.push({
          bytes: source.byteLength,
          endOffset: options?.endOffset,
          startOffset: options?.startOffset,
        })
      },
    }
    const floats = new Float32Array(80)

    writeTypeGpuVertexBufferSubrange({
      buffer,
      floatCount: 20,
      floats,
      label: 'test-subrange',
      startFloat: 40,
    })

    expect(writes).toEqual([
      {
        bytes: 20 * Float32Array.BYTES_PER_ELEMENT,
        endOffset: 60 * Float32Array.BYTES_PER_ELEMENT,
        startOffset: 40 * Float32Array.BYTES_PER_ELEMENT,
      },
    ])
  })
})

function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  return {
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
    ...overrides,
  }
}
