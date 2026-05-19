import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { buildTextQuadsFromRunsWithSpans } from '../renderer-v3/line-text-quad-buffer.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'
import {
  areGridRectTileRevisionKeysEqualV3,
  areGridTextTileRevisionKeysEqualV3,
  resolveGridRectTileRevisionKeyV3,
  resolveGridTextTileRevisionKeyV3,
  resolveMissingTextGlyphRunSpansV3,
  shouldFullWriteTileRectPayloadV3,
  shouldSyncGridRectTileResourceV3,
  shouldSyncGridTextTileResourceV3,
  shouldAttemptAxisOnlyTileTextGeometryResourceSync,
  syncAxisOnlyTileTextGeometryResource,
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
    decorationCellKeys: null,
    decorationRects: null,
    rectBaseCount: overrides.rectBaseCount ?? overrides.rectCount ?? 1,
    rectCount: 1,
    rectHandle: null,
    rectRevisionKey: null,
    textCount: 1,
    textAtlasDependencyVersion: 1,
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
    const changedWithoutSequenceBump = createTile({
      textCount: first.textCount,
      textRuns: first.textRuns.map((run) => ({ ...run, text: 'A2' })),
      version: { ...first.version },
    })

    expect(areGridTextTileRevisionKeysEqualV3(resolveGridTextTileRevisionKeyV3(first), resolveGridTextTileRevisionKeyV3(second))).toBe(true)
    expect(areGridTextTileRevisionKeysEqualV3(resolveGridTextTileRevisionKeyV3(changed), resolveGridTextTileRevisionKeyV3(first))).toBe(
      false,
    )
    expect(
      areGridTextTileRevisionKeysEqualV3(
        resolveGridTextTileRevisionKeyV3(changedWithoutSequenceBump),
        resolveGridTextTileRevisionKeyV3(first),
      ),
    ).toBe(false)
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

  test('detects V3 rect payload changes without relying only on sequence bumps', () => {
    const base = createTile({
      rectSignature: 'base-grid-rects',
    })
    const changedWithoutSequenceBump = createTile({
      rectCount: base.rectCount,
      rectInstances: new Float32Array(base.rectInstances),
      rectSignature: 'changed-grid-rects',
      version: { ...base.version },
    })

    expect(areGridRectTileRevisionKeysEqualV3(rectRevisionKey(changedWithoutSequenceBump), rectRevisionKey(base))).toBe(false)
  })

  test('full-writes V3 rect payloads when same-count fills change without dirty spans', () => {
    expect(
      shouldFullWriteTileRectPayloadV3({
        canWritePartialPayload: true,
        contentRectCount: 1,
        dirtySpans: [],
        rectPayloadCount: 1,
      }),
    ).toBe(true)

    expect(
      shouldFullWriteTileRectPayloadV3({
        canWritePartialPayload: true,
        contentRectCount: 2,
        dirtySpans: [{ length: 1, offset: 1 }],
        rectPayloadCount: 2,
      }),
    ).toBe(false)
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

  test('resyncs V3 text resources when a glyph dependency is missing', () => {
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
        missingGlyphDependencies: true,
        textRevisionKey: revisionKey,
        tile,
      }),
    ).toBe(true)
  })

  test('maps missing glyph dependencies to dirty text runs', () => {
    const atlas = {
      resolveGlyphRecord(glyphId: number) {
        if (glyphId === 2) {
          return null
        }
        return {
          glyphId,
          pageId: glyphId * 10,
          refCount: 1,
          u0: 0,
          u1: 1,
          v0: 0,
          v1: 1,
        }
      },
    }

    expect(
      resolveMissingTextGlyphRunSpansV3({
        atlas,
        content: contentEntry({
          textGlyphIds: [1, 2],
          textGlyphPageIds: [10, 20],
          textRunCount: 2,
          textRunGlyphIds: [[1], [2]],
        }),
      }),
    ).toEqual([{ length: 1, offset: 1 }])
    expect(
      resolveMissingTextGlyphRunSpansV3({
        atlas,
        content: contentEntry({
          textGlyphIds: [1],
          textGlyphPageIds: [20],
          textRunCount: 1,
          textRunGlyphIds: [[1]],
        }),
      }),
    ).toEqual([{ length: 1, offset: 0 }])
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
          decorationCellKeys: new Set(['0:0']),
          decorationRects: [{ color: '#111111', height: 1, width: 20, x: 4, y: 18 }],
          rectCount: 0,
          rectRevisionKey: rectRevisionKey(createTile()),
        }),
        rectRevisionKey: rectRevisionKey(plainAfterDecoration),
        tile: plainAfterDecoration,
      }),
    ).toBe(true)
  })

  test('skips rect uploads when a plain text edit shares a tile with unrelated decorated text', () => {
    const base = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 2,
      textRuns: [
        {
          align: 'left',
          clipHeight: 22,
          clipWidth: 104,
          clipX: 0,
          clipY: 0,
          color: '#111111',
          col: 1,
          font: '400 11px sans-serif',
          fontSize: 11,
          height: 22,
          row: 1,
          strike: false,
          text: 'Decorated',
          underline: true,
          width: 104,
          wrap: false,
          x: 104,
          y: 22,
        },
        {
          align: 'left',
          clipHeight: 22,
          clipWidth: 104,
          clipX: 0,
          clipY: 0,
          color: '#111111',
          col: 5,
          font: '400 11px sans-serif',
          fontSize: 11,
          height: 22,
          row: 5,
          strike: false,
          text: 'Plain',
          underline: false,
          width: 104,
          wrap: false,
          x: 520,
          y: 110,
        },
      ],
    })
    const plainEdit = createTile({
      dirtyLocalCols: new Uint32Array([5, 5]),
      dirtyLocalRows: new Uint32Array([5, 5]),
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      rectCount: base.rectCount,
      textCount: base.textCount,
      textRuns: base.textRuns.map((run) => (run.row === 5 && run.col === 5 ? { ...run, text: 'Plain edited' } : run)),
      version: { ...base.version, text: 2, values: 2 },
    })

    expect(
      shouldSyncGridRectTileResourceV3({
        content: contentEntry({
          decorationCellKeys: new Set(['1:1']),
          decorationRects: [{ color: '#111111', height: 1, width: 52, x: 104, y: 40 }],
          rectCount: base.rectCount,
          rectRevisionKey: rectRevisionKey(base),
        }),
        rectRevisionKey: rectRevisionKey(plainEdit),
        tile: plainEdit,
      }),
    ).toBe(false)
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

  test('writes axis-only text geometry subranges from source zero for dirty shifted runs', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ text: 'A', x: 0, clipX: 0 }),
      createTextRun({ text: 'B', x: 104, clipX: 104 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })

    const shiftedRuns: GridRenderTile['textRuns'] = [
      baseRuns[0],
      {
        ...baseRuns[1],
        clipX: 116,
        x: 116,
      },
    ]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 1, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: shiftedRuns.length,
      textRuns: shiftedRuns,
      version: {
        ...baseTile.version,
        axisX: 2,
      },
    })
    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    const dirtyRunSpan = fullPayload.runSpans[1]
    if (!dirtyRunSpan) {
      throw new Error('Expected a text quad span for the shifted run')
    }
    const expectedStartFloat = dirtyRunSpan.offset * 16
    const expectedFloatCount = dirtyRunSpan.length * 16
    expect(didSync).toBe(true)
    expect(writes).toHaveLength(1)
    expect(writes[0]?.startOffset).toBe(expectedStartFloat * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.endOffset).toBe((expectedStartFloat + expectedFloatCount) * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats).toEqual(Array.from(fullPayload.floats.subarray(expectedStartFloat, expectedStartFloat + expectedFloatCount)))
  })

  test('translates left-aligned axis-only text when clip width changes', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ text: 'A', x: 0, clipX: 0 }),
      createTextRun({ clipWidth: 104, text: 'B', width: 104, x: 104, clipX: 104 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedRuns: GridRenderTile['textRuns'] = [
      baseRuns[0],
      {
        ...baseRuns[1],
        clipWidth: 148,
        width: 148,
      },
    ]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 1, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      textRuns: shiftedRuns,
      version: {
        ...baseTile.version,
        axisX: 2,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    const dirtyRunSpan = fullPayload.runSpans[1]
    if (!dirtyRunSpan) {
      throw new Error('Expected a text quad span for the shifted run')
    }
    const expectedStartFloat = dirtyRunSpan.offset * 16
    const expectedFloatCount = dirtyRunSpan.length * 16
    expect(didSync).toBe(true)
    expect(writes).toHaveLength(1)
    expect(writes[0]?.startOffset).toBe(expectedStartFloat * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.endOffset).toBe((expectedStartFloat + expectedFloatCount) * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats).toEqual(Array.from(fullPayload.floats.subarray(expectedStartFloat, expectedStartFloat + expectedFloatCount)))
  })

  test('translates right-aligned axis-only text when only clip height changes', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ align: 'right', clipHeight: 22, height: 22, row: 1, text: '123', y: 22, clipY: 22 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedRuns: GridRenderTile['textRuns'] = [
      {
        ...baseRuns[0],
        clipHeight: 44,
        height: 44,
      },
    ]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      textRuns: shiftedRuns,
      version: {
        ...baseTile.version,
        axisY: 2,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    expect(didSync).toBe(true)
    expect(writes).toHaveLength(1)
    expect(writes[0]?.floats).toEqual(Array.from(fullPayload.floats))
  })

  test('translates authoritative axis-only text tile revisions without dirty span metadata', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ row: 0, text: 'A', y: 0, clipY: 0 }),
      createTextRun({ row: 1, text: 'B', y: 22, clipY: 22 }),
      createTextRun({ row: 2, text: 'C', y: 44, clipY: 44 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedRuns: GridRenderTile['textRuns'] = [
      Object.assign({}, baseRuns[0], { clipY: baseRuns[0].clipY + 8, y: baseRuns[0].y + 8 }),
      Object.assign({}, baseRuns[1], { clipY: baseRuns[1].clipY + 8, y: baseRuns[1].y + 8 }),
      Object.assign({}, baseRuns[2], { clipY: baseRuns[2].clipY + 8, y: baseRuns[2].y + 8 }),
    ]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: undefined,
      dirtyLocalCols: undefined,
      dirtyLocalRows: undefined,
      dirtyMasks: undefined,
      textCount: shiftedRuns.length,
      textRuns: shiftedRuns,
      version: {
        ...baseTile.version,
        axisY: 2,
        styles: 9,
        text: 9,
        values: 9,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    expect(didSync).toBe(true)
    expect(writes).toHaveLength(1)
    expect(writes[0]?.startOffset).toBe(0)
    expect(writes[0]?.endOffset).toBe(fullPayload.floats.length * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats).toEqual(Array.from(fullPayload.floats))
    expect(content.textRevisionKey).toEqual(resolveGridTextTileRevisionKeyV3(shiftedTile))
  })

  test('uploads only changed runs for authoritative full-tile axis-only revisions', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ row: 0, text: 'A', y: 0, clipY: 0 }),
      createTextRun({ row: 1, text: 'B', y: 22, clipY: 22 }),
      createTextRun({ row: 2, text: 'C', y: 44, clipY: 44 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedRuns: GridRenderTile['textRuns'] = [
      baseRuns[0],
      Object.assign({}, baseRuns[1], { clipX: baseRuns[1].clipX + 11, x: baseRuns[1].x + 11 }),
      baseRuns[2],
    ]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: undefined,
      dirtyLocalCols: undefined,
      dirtyLocalRows: undefined,
      dirtyMasks: undefined,
      textCount: shiftedRuns.length,
      textRuns: shiftedRuns,
      version: {
        ...baseTile.version,
        axisX: 2,
        styles: 9,
        text: 9,
        values: 9,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    const changedSpan = fullPayload.runSpans[1]
    if (!changedSpan) {
      throw new Error('Expected a changed text quad span')
    }
    const expectedStartFloat = changedSpan.offset * 16
    const expectedFloatCount = changedSpan.length * 16
    expect(didSync).toBe(true)
    expect(writes).toHaveLength(1)
    expect(writes[0]?.startOffset).toBe(expectedStartFloat * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.endOffset).toBe((expectedStartFloat + expectedFloatCount) * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats).toEqual(Array.from(fullPayload.floats.subarray(expectedStartFloat, expectedStartFloat + expectedFloatCount)))
    expect(Array.from(content.textRunPayloads?.[0]?.floats ?? [])).toEqual(Array.from(basePayload.runPayloads[0]?.floats ?? []))
    expect(Array.from(content.textRunPayloads?.[2]?.floats ?? [])).toEqual(Array.from(basePayload.runPayloads[2]?.floats ?? []))
  })

  test('rejects authoritative axis-only translation when a run visual signature changed', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [
      createTextRun({ row: 0, text: 'A', y: 0, clipY: 0 }),
      createTextRun({ row: 1, text: 'B', y: 22, clipY: 22 }),
    ]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const changedTile = createTile({
      ...baseTile,
      dirty: undefined,
      dirtyLocalCols: undefined,
      dirtyLocalRows: undefined,
      dirtyMasks: undefined,
      textRuns: [
        {
          ...baseRuns[0],
          clipY: 8,
          y: 8,
        },
        {
          ...baseRuns[1],
          clipY: 30,
          text: 'B changed',
          y: 30,
        },
      ],
      version: {
        ...baseTile.version,
        axisY: 2,
        text: 2,
        values: 2,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: changedTile.dirty?.textSpans,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(changedTile),
      tile: changedTile,
    })

    expect(didSync).toBe(false)
    expect(writes).toHaveLength(0)
  })

  test('does not use axis-only text translation when glyph dependencies are missing', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [createTextRun({ text: 'A', x: 0, clipX: 0 })]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      textRuns: [{ ...baseRuns[0], clipX: 16, x: 16 }],
      version: {
        ...baseTile.version,
        axisX: 2,
      },
    })

    const didSync = syncAxisOnlyTileTextGeometryResource({
      content,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      hasMissingGlyphDependencies: true,
      handle,
      label: 'tile-text:test',
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    })

    expect(didSync).toBe(false)
    expect(writes).toHaveLength(0)
  })

  test('keeps axis-only text geometry byte-equivalent to full rebuild across resize geometry matrix', () => {
    const cases: readonly {
      readonly dirtyMasks: Uint32Array
      readonly name: string
      readonly baseRun: GridRenderTile['textRuns'][number]
      readonly shiftedRun: GridRenderTile['textRuns'][number]
    }[] = [
      {
        baseRun: createTextRun({ clipX: 12.5, text: 'Left', x: 12.5 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'fractional left x shift',
        shiftedRun: createTextRun({ clipX: 18.25, text: 'Left', x: 18.25 }),
      },
      {
        baseRun: createTextRun({ align: 'center', clipWidth: 104, text: 'Center', width: 104 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'center width growth',
        shiftedRun: createTextRun({ align: 'center', clipWidth: 139.5, text: 'Center', width: 139.5 }),
      },
      {
        baseRun: createTextRun({ align: 'right', clipWidth: 104, text: 'Right', width: 104 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'right width shrink',
        shiftedRun: createTextRun({ align: 'right', clipWidth: 87.25, text: 'Right', width: 87.25 }),
      },
      {
        baseRun: createTextRun({ clipHeight: 22, clipY: 64, height: 22, text: 'Tall', y: 64 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'height growth keeps vertical centering',
        shiftedRun: createTextRun({ clipHeight: 48.5, clipY: 64, height: 48.5, text: 'Tall', y: 64 }),
      },
      {
        baseRun: createTextRun({ clipHeight: 22, clipWidth: 104, clipX: 260, clipY: 88, text: 'Both', x: 260, y: 88 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'both axes and clip dimensions',
        shiftedRun: createTextRun({
          clipHeight: 31.75,
          clipWidth: 126.5,
          clipX: 271.25,
          clipY: 93.5,
          height: 31.75,
          text: 'Both',
          width: 126.5,
          x: 271.25,
          y: 93.5,
        }),
      },
      {
        baseRun: createTextRun({ clipHeight: 1, clipWidth: 1, text: 'Z', width: 1 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'near-zero visible clip',
        shiftedRun: createTextRun({ clipHeight: 1.5, clipWidth: 1.25, text: 'Z', width: 1.25 }),
      },
      {
        baseRun: createTextRun({ clipX: 1_000_000, clipY: 2_000_000, text: 'Large', x: 1_000_000, y: 2_000_000 }),
        dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        name: 'large world coordinates',
        shiftedRun: createTextRun({ clipX: 1_000_128, clipY: 2_000_064, text: 'Large', x: 1_000_128, y: 2_000_064 }),
      },
    ]

    for (const entry of cases) {
      expectAxisOnlySyncMatchesFullTextPayload(entry)
    }
  })

  test('routes missing glyph axis-only candidates through full dirty rebuild instead of stale payload reuse', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [createTextRun({ text: 'Glyph', x: 0, clipX: 0 })]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const shiftedRuns: GridRenderTile['textRuns'] = [{ ...baseRuns[0], clipX: 16, x: 16 }]
    const shiftedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      textRuns: shiftedRuns,
      version: { ...baseTile.version, axisX: 2 },
    })

    expect(
      shouldAttemptAxisOnlyTileTextGeometryResourceSync({
        contentRevisionKey: content.textRevisionKey,
        dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
        textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
        tile: shiftedTile,
      }),
    ).toBe(true)
    expect(
      syncAxisOnlyTileTextGeometryResource({
        content,
        dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
        hasMissingGlyphDependencies: true,
        handle,
        label: 'tile-text:test',
        textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
        tile: shiftedTile,
      }),
    ).toBe(false)
    expect(writes).toHaveLength(0)
    expect(content.textRevisionKey).toEqual(resolveGridTextTileRevisionKeyV3(baseTile))

    const fallbackPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas, undefined, {
      dirtyRunSpans: shiftedTile.dirty?.textSpans,
      forceRebuildDirtyRunSpans: true,
      previousRunPayloads: content.textRunPayloads,
    })
    const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)
    expect(fallbackPayload.diagnostics.rebuiltRunPayloads).toBe(1)
    expect(fallbackPayload.diagnostics.reusedRunPayloads).toBe(0)
    expect(Array.from(fallbackPayload.floats)).toEqual(Array.from(fullPayload.floats))
  })

  test('rebuilds dirty text after axis-only signature rejection so stale cache cannot be promoted', () => {
    const writes: RecordedVertexWrite[] = []
    const atlas = createTestAtlas()
    const handle = createRecordedTextHandle(writes)
    const baseRuns: GridRenderTile['textRuns'] = [createTextRun({ text: 'Old', x: 0, clipX: 0 })]
    const baseTile = createTile({
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })
    const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
    const content = contentEntry({
      rectBaseCount: 0,
      rectCount: 0,
      textCount: basePayload.quadCount,
      textGlyphIds: basePayload.glyphIds,
      textGlyphPageIds: basePayload.pageIds,
      textRunCount: baseRuns.length,
      textRunGlyphIds: basePayload.runGlyphIds,
      textRunPayloads: basePayload.runPayloads,
      textRunQuadSpans: basePayload.runSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
    })
    const changedRuns: GridRenderTile['textRuns'] = [{ ...baseRuns[0], clipX: 16, text: 'New value', x: 16 }]
    const changedTile = createTile({
      ...baseTile,
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: 1 }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
      textRuns: changedRuns,
      version: { ...baseTile.version, axisX: 2, text: 2, values: 2 },
    })

    expect(
      shouldAttemptAxisOnlyTileTextGeometryResourceSync({
        contentRevisionKey: content.textRevisionKey,
        dirtyTextRunSpans: changedTile.dirty?.textSpans,
        textRevisionKey: resolveGridTextTileRevisionKeyV3(changedTile),
        tile: changedTile,
      }),
    ).toBe(true)
    expect(
      syncAxisOnlyTileTextGeometryResource({
        content,
        dirtyTextRunSpans: changedTile.dirty?.textSpans,
        handle,
        label: 'tile-text:test',
        textRevisionKey: resolveGridTextTileRevisionKeyV3(changedTile),
        tile: changedTile,
      }),
    ).toBe(false)
    expect(writes).toHaveLength(0)

    const fallbackPayload = buildTextQuadsFromRunsWithSpans(changedRuns, atlas, undefined, {
      dirtyRunSpans: changedTile.dirty?.textSpans,
      previousRunPayloads: content.textRunPayloads,
    })
    const fullPayload = buildTextQuadsFromRunsWithSpans(changedRuns, atlas)
    expect(fallbackPayload.diagnostics.rebuiltRunPayloads).toBe(1)
    expect(fallbackPayload.diagnostics.reusedRunPayloads).toBe(0)
    expect(Array.from(fallbackPayload.floats)).toEqual(Array.from(fullPayload.floats))
    expect(fallbackPayload.runPayloads[0]?.contentSignature).toBe(fullPayload.runPayloads[0]?.contentSignature)
    expect(fallbackPayload.runPayloads[0]?.contentSignature).not.toBe(basePayload.runPayloads[0]?.contentSignature)
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

interface RecordedVertexWrite {
  readonly endOffset?: number | undefined
  readonly floats: readonly number[]
  readonly startOffset?: number | undefined
}

function createRecordedTextHandle(
  writes: RecordedVertexWrite[],
): NonNullable<Parameters<typeof syncAxisOnlyTileTextGeometryResource>[0]['handle']> {
  return {
    buffer: {
      write(source: ArrayBuffer, options?: { readonly startOffset?: number | undefined; readonly endOffset?: number | undefined }) {
        writes.push({
          endOffset: options?.endOffset,
          floats: Array.from(new Float32Array(source)),
          startOffset: options?.startOffset,
        })
      },
    },
    capacityBytes: 4096,
    classId: 4,
    layout: 'textRuns',
    usedBytes: 0,
  }
}

function expectAxisOnlySyncMatchesFullTextPayload(input: {
  readonly dirtyMasks: Uint32Array
  readonly name: string
  readonly baseRun: GridRenderTile['textRuns'][number]
  readonly shiftedRun: GridRenderTile['textRuns'][number]
}): void {
  const writes: RecordedVertexWrite[] = []
  const atlas = createTestAtlas()
  const handle = createRecordedTextHandle(writes)
  const baseRuns: GridRenderTile['textRuns'] = [input.baseRun]
  const baseTile = createTile({
    rectCount: 0,
    rectInstances: new Float32Array(),
    textCount: baseRuns.length,
    textRuns: baseRuns,
  })
  const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
  const content = contentEntry({
    rectBaseCount: 0,
    rectCount: 0,
    textCount: basePayload.quadCount,
    textGlyphIds: basePayload.glyphIds,
    textGlyphPageIds: basePayload.pageIds,
    textRunCount: baseRuns.length,
    textRunGlyphIds: basePayload.runGlyphIds,
    textRunPayloads: basePayload.runPayloads,
    textRunQuadSpans: basePayload.runSpans,
    textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
  })
  const shiftedRuns: GridRenderTile['textRuns'] = [input.shiftedRun]
  const shiftedTile = createTile({
    ...baseTile,
    dirty: {
      glyphSpans: [],
      rectSpans: [],
      textSpans: [{ offset: 0, length: 1 }],
    },
    dirtyMasks: input.dirtyMasks,
    textRuns: shiftedRuns,
    version: {
      ...baseTile.version,
      axisX: 2,
      axisY: 2,
    },
  })
  expect(
    shouldAttemptAxisOnlyTileTextGeometryResourceSync({
      contentRevisionKey: content.textRevisionKey,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    }),
    input.name,
  ).toBe(true)

  const didSync = syncAxisOnlyTileTextGeometryResource({
    content,
    dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
    handle,
    label: 'tile-text:test',
    textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
    tile: shiftedTile,
  })
  const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)

  expect(didSync, input.name).toBe(true)
  if (areNumberArraysEqual(Array.from(basePayload.floats), Array.from(fullPayload.floats))) {
    expect(writes, input.name).toHaveLength(0)
  } else {
    expect(writes, input.name).toHaveLength(1)
    expect(writes[0]?.startOffset, input.name).toBe(0)
    expect(writes[0]?.endOffset, input.name).toBe(fullPayload.floats.length * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats, input.name).toEqual(Array.from(fullPayload.floats))
  }
  expect(Array.from(content.textRunPayloads?.[0]?.floats ?? []), input.name).toEqual(Array.from(fullPayload.runPayloads[0]?.floats ?? []))
}

function areNumberArraysEqual(left: readonly number[], right: readonly number[]): boolean {
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

function createTestAtlas(): Parameters<typeof buildTextQuadsFromRunsWithSpans>[1] {
  return {
    getGlyphGeometryVersion: () => 1,
    getVersion: () => 1,
    intern(font: string, glyph: string) {
      const advance = Math.max(0, glyph.length * 8)
      return {
        advance,
        baseline: 10,
        font,
        glyph,
        glyphId: glyph.codePointAt(0) ?? 0,
        height: 12,
        key: `atlas:${glyph}`,
        originOffsetX: 0,
        pageId: 1,
        u0: 0,
        u1: 1,
        v0: 0,
        v1: 1,
        width: advance,
        x: 0,
        y: 0,
      }
    },
  }
}
