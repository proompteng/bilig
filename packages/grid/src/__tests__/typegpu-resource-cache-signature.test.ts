import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'
import {
  areGridRectTileRevisionKeysEqualV3,
  areGridTextTileRevisionKeysEqualV3,
  resolveGridRectTileRevisionKeyV3,
  resolveGridTextTileRevisionKeyV3,
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

function rectRevisionKey(tile: GridRenderTile): ReturnType<typeof resolveGridRectTileRevisionKeyV3> {
  return resolveGridRectTileRevisionKeyV3({ tile })
}

function contentEntry(overrides: Partial<TypeGpuTileContentResourceEntryV3> = {}): TypeGpuTileContentResourceEntryV3 {
  return {
    decorationRects: null,
    rectCount: 1,
    rectHandle: null,
    rectRevisionKey: null,
    textCount: 1,
    textAtlasGeometryVersion: 1,
    textGlyphIds: null,
    textGlyphPageIds: null,
    textHandle: null,
    textRunGlyphIds: null,
    textRunCount: 1,
    textRunPayloads: null,
    textRunQuadSpans: null,
    textRevisionKey: null,
    ...overrides,
  }
}

describe('typegpu v3 resource cache revision keys', () => {
  test('keeps equivalent V3 text tile revision keys stable without relying on object identity', () => {
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

    expect(areGridTextTileRevisionKeysEqualV3(resolveGridTextTileRevisionKeyV3(first), resolveGridTextTileRevisionKeyV3(second))).toBe(true)
    expect(areGridTextTileRevisionKeysEqualV3(resolveGridTextTileRevisionKeyV3(changed), resolveGridTextTileRevisionKeyV3(first))).toBe(
      false,
    )
  })

  test('keeps V3 resource revision keys stable across camera sequence churn', () => {
    const base = createTile()
    const newerCameraWithSameContent = createTile({
      lastCameraSeq: 22,
      version: { ...base.version },
    })

    expect(areGridRectTileRevisionKeysEqualV3(rectRevisionKey(newerCameraWithSameContent), rectRevisionKey(base))).toBe(true)
    expect(
      areGridTextTileRevisionKeysEqualV3(
        resolveGridTextTileRevisionKeyV3(newerCameraWithSameContent),
        resolveGridTextTileRevisionKeyV3(base),
      ),
    ).toBe(true)
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
          textRevisionKey: resolveGridTextTileRevisionKeyV3(tile),
        }),
        textRevisionKey: resolveGridTextTileRevisionKeyV3(tile),
        tile,
      }),
    ).toBe(false)
  })

  test('resyncs V3 text resources only when atlas glyph geometry changes', () => {
    const tile = createTile({ textCount: 1, textRuns: [createTextRun({ text: 'A' })] })
    const revisionKey = resolveGridTextTileRevisionKeyV3(tile)

    expect(
      shouldSyncGridTextTileResourceV3({
        atlasGeometryVersion: 1,
        content: contentEntry({
          textAtlasGeometryVersion: 1,
          textCount: 1,
          textRunCount: 1,
          textRevisionKey: revisionKey,
        }),
        textRevisionKey: revisionKey,
        tile,
      }),
    ).toBe(false)
    expect(
      shouldSyncGridTextTileResourceV3({
        atlasGeometryVersion: 2,
        content: contentEntry({
          textAtlasGeometryVersion: 1,
          textCount: 1,
          textRunCount: 1,
          textRevisionKey: revisionKey,
        }),
        textRevisionKey: revisionKey,
        tile,
      }),
    ).toBe(true)
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
          textRevisionKey: resolveGridTextTileRevisionKeyV3(createTile()),
        }),
        textRevisionKey: resolveGridTextTileRevisionKeyV3(tile),
        tile,
      }),
    ).toBe(true)
  })

  test('tracks V3 tile revisions and decoration counts in resource revision keys', () => {
    const base = createTile()
    const changedValues = createTile({ version: { ...base.version, values: 2 } })
    const changedBatch = createTile({ lastBatchId: 2, version: { ...base.version } })

    expect(areGridRectTileRevisionKeysEqualV3(rectRevisionKey(changedValues), rectRevisionKey(base))).toBe(false)
    expect(areGridTextTileRevisionKeysEqualV3(resolveGridTextTileRevisionKeyV3(changedBatch), resolveGridTextTileRevisionKeyV3(base))).toBe(
      false,
    )
    expect(
      areGridRectTileRevisionKeysEqualV3(
        resolveGridRectTileRevisionKeyV3({
          decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
          tile: base,
        }),
        rectRevisionKey(base),
      ),
    ).toBe(false)
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
        content: contentEntry({ textRevisionKey: resolveGridTextTileRevisionKeyV3(base) }),
        textRevisionKey: resolveGridTextTileRevisionKeyV3(textOnlyUpdate),
        tile: textOnlyUpdate,
      }),
    ).toBe(true)
    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({ rectCount: 0, rectRevisionKey: rectRevisionKey(base) }),
        rectRevisionKey: rectRevisionKey(textOnlyUpdate),
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
        content: contentEntry({ rectCount: 0, rectRevisionKey: rectRevisionKey(createTile()) }),
        rectRevisionKey: rectRevisionKey(decorated),
        tile: decorated,
      }),
    ).toBe(true)
    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({
          decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
          rectCount: 0,
          rectRevisionKey: rectRevisionKey(createTile()),
        }),
        rectRevisionKey: rectRevisionKey(plainAfterDecoration),
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
