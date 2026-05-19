import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'
import { materializeGridRenderTileV3 } from '../renderer-v3/grid-tile-materializer.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { GridVisibleTextRefreshCache } from '../runtime/gridVisibleTextRefreshCache.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'

const TEST_GRID_METRICS = {
  ...getGridMetrics(),
  columnWidth: 100,
  rowHeight: 20,
}

describe('GridVisibleTextRefreshCache', () => {
  it('rejects unanchored remote text runs instead of letting ghost text render', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 1,
      textRuns: [
        {
          ...createTextRun({ text: 'ghost text' }),
          col: undefined,
          row: undefined,
        },
      ],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine() }))).toBe(true)
  })

  it('rejects inconsistent remote text run counts', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 2,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'A1' }) }))).toBe(true)
  })

  it('rejects duplicate visible text runs for the same cell instead of allowing double text artifacts', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 2,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' }), createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'A1' }) }))).toBe(true)
  })

  it('accepts remote text only when the visible cell text matches the workbook state', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 1,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'A1' }) }))).toBe(false)
    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'fresh A1' }) }))).toBe(true)
  })

  it('rejects matching remote text when the visible font styling is stale', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 1,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-bold', value: 'A1' },
            },
            {
              'style-bold': { id: 'style-bold', font: { bold: true } },
            },
          ),
        }),
      ),
    ).toBe(true)
  })

  it('rejects style-stale visible fills even when text is unchanged', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      textCount: 1,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 1,
        text: 7,
        values: 7,
      },
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-green', value: 'A1' },
            },
            {
              'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
            },
            7,
          ),
        }),
      ),
    ).toBe(true)
  })

  it('rejects current-revision remote tiles that are missing the visible fill color', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      textCount: 1,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-green', value: 'A1' },
            },
            {
              'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
            },
            7,
          ),
        }),
      ),
    ).toBe(true)
  })

  it('rejects stale fill rects after visible cell fills are cleared', () => {
    const cache = new GridVisibleTextRefreshCache()
    const fillRects = new Float32Array(GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
    fillRects[2] = 100
    fillRects[3] = 20
    fillRects[5] = 1
    fillRects[7] = 1
    fillRects[13] = 0
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: fillRects,
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 1,
        text: 7,
        values: 7,
      },
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({}, {}, 7) }))).toBe(true)
  })

  it('rejects current-revision stale fill rects after visible cell fills are cleared', () => {
    const cache = new GridVisibleTextRefreshCache()
    const fillRects = new Float32Array(GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
    fillRects[2] = 100
    fillRects[3] = 20
    fillRects[5] = 1
    fillRects[7] = 1
    fillRects[13] = 0
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: fillRects,
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({}, {}, 7) }))).toBe(true)
  })

  it('rejects current-revision partial stale fill slivers when projected tiles have no rect signature', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: createFillRectInstances([
        {
          height: TEST_GRID_METRICS.rowHeight,
          width: 12,
        },
      ]),
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({}, {}, 7) }))).toBe(true)
  })

  it('rejects current-revision same-color extra fill covering a now-unfilled visible cell', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: createFillRectInstances([
        {
          width: TEST_GRID_METRICS.columnWidth * 2,
        },
      ]),
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-green' },
            },
            {
              'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
            },
            7,
          ),
        }),
      ),
    ).toBe(true)
  })

  it('rejects current-revision same-color fill at the wrong visible location', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: createFillRectInstances([
        {
          x: TEST_GRID_METRICS.columnWidth,
        },
      ]),
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-green' },
            },
            {
              'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
            },
            7,
          ),
        }),
      ),
    ).toBe(true)
  })

  it('accepts coalesced adjacent same-color fills when every covered visible cell expects that fill', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      lastBatchId: 7,
      rectCount: 1,
      rectInstances: createFillRectInstances([
        {
          width: TEST_GRID_METRICS.columnWidth * 2,
        },
      ]),
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    })

    expect(
      cache.needsLocalRefresh(
        tile.tileId,
        tile,
        createInput({
          engine: createEngine(
            {
              A1: { styleId: 'style-green' },
              B1: { styleId: 'style-green' },
            },
            {
              'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
            },
            7,
          ),
        }),
      ),
    ).toBe(false)
  })

  it('accepts a materialized current rect signature for visible fills, borders, and gridlines', () => {
    const cache = new GridVisibleTextRefreshCache()
    const engine = createEngine(
      {
        A1: { styleId: 'style-green-border', value: 'A1' },
      },
      {
        'style-green-border': {
          borders: { bottom: { color: '#ff0000', style: 'solid' } },
          fill: { backgroundColor: '#00ff00' },
          id: 'style-green-border',
        },
      },
      7,
    )
    const tile = createMaterializedTile(engine, 7)

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine }))).toBe(false)
  })

  it('rejects current-revision stale border or checkbox rect payload signatures', () => {
    const cache = new GridVisibleTextRefreshCache()
    const engine = createEngine(
      {
        A1: { styleId: 'style-border', value: 'A1' },
      },
      {
        'style-border': {
          borders: { bottom: { color: '#ff0000', style: 'solid' } },
          id: 'style-border',
        },
      },
      7,
    )
    const tile = createMaterializedTile(engine, 7)
    const staleTile = {
      ...tile,
      rectSignature: 'stale-border-signature',
    }

    expect(cache.needsLocalRefresh(staleTile.tileId, staleTile, createInput({ engine }))).toBe(true)
  })

  it('accepts merged-fill rect signatures instead of permanently localizing merged ranges', () => {
    const cache = new GridVisibleTextRefreshCache()
    const engine = createEngine(
      {
        A1: { styleId: 'style-green', value: 'merged title' },
      },
      {
        'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
      },
      7,
      {
        A1: { endAddress: 'B1', sheetName: 'Sheet1', startAddress: 'A1' },
        B1: { endAddress: 'B1', sheetName: 'Sheet1', startAddress: 'A1' },
      },
    )
    const tile = createMaterializedTile(engine, 7)

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine }))).toBe(false)
  })

  it('rejects stale authored border rects after visible cell borders are cleared', () => {
    const cache = new GridVisibleTextRefreshCache()
    const staleBorderRectCount = 32 + 128 + 1
    const staleBorderRects = new Float32Array(staleBorderRectCount * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
    staleBorderRects[11] = 1
    staleBorderRects[13] = 1
    const tile = createTile({
      lastBatchId: 7,
      rectCount: staleBorderRectCount,
      rectInstances: staleBorderRects,
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 1,
        text: 7,
        values: 7,
      },
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({}, {}, 7) }))).toBe(true)
  })
})

function createInput(overrides: Partial<Parameters<GridVisibleTextRefreshCache['needsLocalRefresh']>[2]> = {}) {
  return {
    columnWidths: {},
    engine: createEngine(),
    gridMetrics: TEST_GRID_METRICS,
    rowHeights: {},
    sceneRevision: 1,
    sheetName: 'Sheet1',
    sortedColumnWidthOverrides: [],
    sortedRowHeightOverrides: [],
    visibleViewport: { colEnd: 2, colStart: 0, rowEnd: 2, rowStart: 0 },
    ...overrides,
  }
}

type EngineCellFixture = string | { readonly value?: string | undefined; readonly styleId?: string | undefined }

function createEngine(
  values: Record<string, EngineCellFixture> = {},
  styles: Record<string, CellStyleRecord> = {},
  projectedRevision = 1,
  mergedRanges: Record<string, { sheetName: string; startAddress: string; endAddress: string }> = {},
): GridEngineLike {
  return {
    getCell: (_sheetName, address) => createCellSnapshot(address, values[address] ?? ''),
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    getMergeRange: (_sheetName, address) => mergedRanges[address],
    getRenderRevisionSnapshot: () => ({
      authoritativeRevision: projectedRevision,
      localRevision: projectedRevision,
      projectedRevision,
      tileSceneCameraSeq: projectedRevision,
      tileSceneRevision: projectedRevision,
    }),
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  }
}

function createMaterializedTile(engine: GridEngineLike, projectedRevision: number): GridRenderTile {
  return materializeGridRenderTileV3({
    axisSeqX: projectedRevision,
    axisSeqY: projectedRevision,
    cameraSeq: projectedRevision,
    columnWidths: {},
    dirtyMasks: undefined,
    dprBucket: 1,
    engine,
    freezeSeq: projectedRevision,
    gridMetrics: TEST_GRID_METRICS,
    materializedAtSeq: projectedRevision,
    packetSeq: projectedRevision,
    rectSeq: projectedRevision,
    rowHeights: {},
    sheetId: 1,
    sheetName: 'Sheet1',
    sheetOrdinal: 1,
    sortedColumnWidthOverrides: [],
    sortedRowHeightOverrides: [],
    styleSeq: projectedRevision,
    textSeq: projectedRevision,
    valueSeq: projectedRevision,
    viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
  })
}

function createCellSnapshot(address: string, fixture: EngineCellFixture): CellSnapshot {
  const value = typeof fixture === 'string' ? fixture : (fixture.value ?? '')
  const styleId = typeof fixture === 'string' ? undefined : fixture.styleId
  if (value.length === 0) {
    return {
      address,
      flags: 0,
      sheetName: 'Sheet1',
      ...(styleId ? { styleId } : {}),
      value: { tag: ValueTag.Empty },
      version: 1,
    }
  }
  return {
    address,
    flags: 0,
    input: value,
    sheetName: 'Sheet1',
    ...(styleId ? { styleId } : {}),
    value: { tag: ValueTag.String, value, stringId: 1 },
    version: 1,
  }
}

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  return {
    bounds: { colEnd: 2, colStart: 0, rowEnd: 2, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 1,
      sheetOrdinal: 1,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 1,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 1,
      styles: 1,
      text: 1,
      values: 1,
    },
    ...overrides,
  }
}

function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  const fontSize = workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
  return {
    clipHeight: 20,
    clipWidth: 100,
    clipX: 0,
    clipY: 0,
    col: 0,
    color: '#1f2933',
    font: `400 ${fontSize}px ${WORKBOOK_FONT_SANS}`,
    fontSize,
    height: 20,
    row: 0,
    strike: false,
    text: '',
    underline: false,
    width: 100,
    x: 0,
    y: 0,
    ...overrides,
  }
}

function createFillRectInstances(
  rects: readonly {
    readonly a?: number | undefined
    readonly b?: number | undefined
    readonly g?: number | undefined
    readonly height?: number | undefined
    readonly r?: number | undefined
    readonly width?: number | undefined
    readonly x?: number | undefined
    readonly y?: number | undefined
  }[],
): Float32Array {
  const fillRects = new Float32Array(Math.max(1, rects.length) * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
  rects.forEach((rect, index) => {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    fillRects[offset + 0] = rect.x ?? 0
    fillRects[offset + 1] = rect.y ?? 0
    fillRects[offset + 2] = rect.width ?? TEST_GRID_METRICS.columnWidth
    fillRects[offset + 3] = rect.height ?? TEST_GRID_METRICS.rowHeight
    fillRects[offset + 4] = rect.r ?? 0
    fillRects[offset + 5] = rect.g ?? 1
    fillRects[offset + 6] = rect.b ?? 0
    fillRects[offset + 7] = rect.a ?? 1
    fillRects[offset + 13] = 0
  })
  return fillRects
}
