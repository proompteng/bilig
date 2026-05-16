import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { encodeViewportPatch, type ViewportPatch, type WorkerEngineClient } from '@bilig/worker-transport'
import { DEFAULT_MAX_CACHED_CELLS_PER_SHEET } from '../projected-viewport-cell-cache.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createPatch(styleId?: string): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 7,
      colStart: 2,
      colEnd: 4,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [
      {
        row: 4,
        col: 3,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'D5',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
          ...(styleId ? { styleId } : {}),
        },
        displayText: '',
        copyText: '',
        editorText: '',
        formatId: 0,
        styleId: styleId ?? 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

function createColumnPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    columns: [{ index: 0, size, hidden }],
  }
}

function createRowPatch(size: number, hidden = false): ViewportPatch {
  return {
    ...createPatch(),
    cells: [],
    rows: [{ index: 0, size, hidden }],
  }
}

function columnLabel(columnIndex: number): string {
  let index = columnIndex + 1
  let label = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    index = Math.floor((index - 1) / 26)
  }
  return label
}

function createLargePatch(rowCount: number, columnCount: number): ViewportPatch {
  return {
    version: 1,
    full: false,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: rowCount - 1,
      colStart: 0,
      colEnd: columnCount - 1,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: Array.from({ length: rowCount * columnCount }, (_, index) => {
      const row = Math.floor(index / columnCount)
      const col = index % columnCount
      return {
        row,
        col,
        snapshot: {
          sheetName: 'Sheet1',
          address: `${columnLabel(col)}${row + 1}`,
          value: { tag: ValueTag.Number, value: index },
          flags: 0,
          version: 1,
        },
        displayText: String(index),
        copyText: String(index),
        editorText: String(index),
        formatId: 0,
        styleId: 'style-0',
      }
    }),
    columns: [],
    rows: [],
  }
}

function countSheetCells(cache: ProjectedViewportStore, sheetName: string): number {
  let count = 0
  cache.workbook.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1
  })
  return count
}

function createNoopWorkerEngineClient(): WorkerEngineClient {
  return {
    dispose: vi.fn(),
    invoke: vi.fn(async () => undefined),
    ready: vi.fn(async () => undefined),
    subscribe: vi.fn(() => () => undefined),
    subscribeBatches: vi.fn(() => () => undefined),
    subscribeRenderTileDeltas: vi.fn(() => () => undefined),
    subscribeViewportPatches: vi.fn(() => () => undefined),
    subscribeWorkbookDeltas: vi.fn(() => () => undefined),
  }
}

describe('ProjectedViewportStore', () => {
  it('accepts equal-version empty snapshots that clear stale styling', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-red')

    cache.applyViewportPatch(createPatch())

    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('exposes authoritative, projected, and tile-scene render revisions for visible proof checks', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 17,
      metrics: {
        ...TEST_METRICS,
        batchId: 23,
      },
    })

    expect(cache.getRenderRevisionSnapshot()).toEqual({
      authoritativeRevision: 17,
      projectedRevision: 23,
      tileSceneCameraSeq: null,
      tileSceneRevision: null,
    })
  })

  it('accepts partial reset-empty patches that clear newer local input after structural deletes', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'stale-tail', stringId: 1 },
      input: 'stale-tail',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 0,
      full: false,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5').value).toEqual({ tag: ValueTag.Empty })
    expect(cache.getCell('Sheet1', 'D5').input).toBeUndefined()
  })

  it('keeps newer local input when a full reset-empty patch arrives during hydration', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'local-edit', stringId: 1 },
      input: 'local-edit',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'local-edit' },
      input: 'local-edit',
      version: 7,
    })
  })

  it('keeps an equal-version local formula snapshot when a later patch drops the formula', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 3,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a newer eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a direct cell refresh drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: 0,
      version: 3,
    })

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: false },
      flags: 0,
      version: 4,
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('accepts a newer literal snapshot when the source input is present', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            input: false,
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const snapshot = cache.getCell('Sheet1', 'D5')
    expect(snapshot).toMatchObject({
      value: { tag: ValueTag.Boolean, value: false },
      input: false,
      version: 4,
    })
    expect('formula' in snapshot).toBe(false)
  })

  it('reports damage when a style record changes without a newer cell snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#c9daf8' } }],
    })

    const damage = cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.getCellStyle('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#a4c2f4' },
    })
  })

  it('tracks freeze pane metadata from viewport patches', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 2,
      freezeCols: 1,
    })

    expect(cache.getFreezeRows('Sheet1')).toBe(2)
    expect(cache.getFreezeCols('Sheet1')).toBe(1)
  })

  it('does not notify freeze subscribers when a patch confirms the default unfrozen state', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 0,
      freezeCols: 0,
    })

    expect(freezeListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
  })

  it('clears stale viewport cells on full patches without dropping cells outside the viewport', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.String, value: 'pinned', stringId: 1 },
      flags: 0,
      version: 1,
    })
    cache.applyViewportPatch({ ...createPatch(), full: true })

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.String,
      value: 'pinned',
      stringId: 1,
    })
  })

  it('drops stale sheet cache entries when sheets disappear', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch())
    expect(cache.peekCell('Sheet1', 'D5')).toBeDefined()

    cache.setKnownSheets(['Sheet2'])

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
  })

  it('resets same-sheet projected state before installing a replacement authoritative snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    cache.setColumnWidth('Sheet1', 0, 68)
    cache.setRowHeight('Sheet1', 0, 240)

    cache.resetProjectionState()

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.workbook.getSheet('Sheet1')).toBeDefined()
    expect(cache.getColumnWidths('Sheet1')[0]).toBeUndefined()
    expect(cache.getRowHeights('Sheet1')[0]).toBeUndefined()
  })

  it('publishes sparse local column axis deltas without materializing default axis sizes', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'columnWidths', () => {
      events.push(`axis:${cache.getColumnWidths('Sheet1')[2]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${cache.getColumnWidths('Sheet1')[2]}:${[...batch.dirty.axisX].join(':')}:${batch.dirty.axisY.length}`)
    })

    cache.setColumnWidth('Sheet1', 2, 144)

    expect(events).toEqual(['axis:144', 'delta:144:2:2:44:0'])
    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })

  it('publishes sparse local row axis deltas without materializing default axis sizes', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'rowHeights', () => {
      events.push(`axis:${cache.getRowHeights('Sheet1')[3]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${cache.getRowHeights('Sheet1')[3]}:${batch.dirty.axisX.length}:${[...batch.dirty.axisY].join(':')}`)
    })

    cache.setRowHeight('Sheet1', 3, 48)

    expect(events).toEqual(['axis:48', 'delta:48:0:3:3:76'])
    expect(cache.getRowHeights('Sheet1')).toEqual({ 3: 48 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })

  it('rejects invalid local axis mutations before state and dirty deltas diverge', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()

    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    expect(() => cache.setColumnWidth('Sheet1', -1, 144)).toThrow('Invalid projected column index: -1')
    expect(() => cache.setRowHeight('Sheet1', 1.5, 48)).toThrow('Invalid projected row index: 1.5')
    expect(() => cache.setColumnHidden('Sheet1', 0, true, Number.NaN)).toThrow('Invalid projected column size: NaN')
    expect(() => cache.rollbackRowHidden('Sheet1', 0, { hidden: false, size: Number.POSITIVE_INFINITY })).toThrow(
      'Invalid projected row size: Infinity',
    )

    expect(cache.getColumnWidths('Sheet1')).toEqual({})
    expect(cache.getRowHeights('Sheet1')).toEqual({})
    expect(listener).not.toHaveBeenCalled()

    unsubscribeDeltas()
  })

  it('drops stale sheet identities when runtime state publishes a replacement sheet list', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()

    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 0 }])
    cache.setSheetIdentities([{ id: 8, name: 'Sheet2', order: 0 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    cache.setCellSnapshot({
      address: 'A1',
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 17 },
      version: 12,
    })

    expect(listener).not.toHaveBeenCalled()

    cache.setCellSnapshot({
      address: 'A1',
      flags: 0,
      sheetName: 'Sheet2',
      value: { tag: ValueTag.Number, value: 18 },
      version: 13,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetId: 8,
        sheetOrdinal: 0,
        source: 'localOptimistic',
      }),
    )

    unsubscribeDeltas()
  })

  it('clears a pending local column width once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.applyViewportPatch(createColumnPatch(68))
    cache.applyViewportPatch(createColumnPatch(93))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
  })

  it('keeps pending local column width visible across older authoritative patches', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.applyViewportPatch(createColumnPatch(93))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(68)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)

    cache.applyViewportPatch(createColumnPatch(68))
    cache.applyViewportPatch(createColumnPatch(104))

    expect(cache.getColumnWidths('Sheet1')[0] ?? 104).toBe(104)
    expect(cache.getColumnSizes('Sheet1')[0] ?? 104).toBe(104)
  })

  it('keeps default authoritative axis sizes sparse after a local resize', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [],
      columns: Array.from({ length: 8 }, (_, index) => ({ index, size: 104, hidden: false })),
      rows: Array.from({ length: 8 }, (_, index) => ({ index, size: 22, hidden: false })),
    })
    cache.setColumnWidth('Sheet1', 2, 144)
    cache.setRowHeight('Sheet1', 3, 48)

    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })
    expect(cache.getRowHeights('Sheet1')).toEqual({ 3: 48 })
  })

  it('rolls back a failed local column width mutation without leaving a pending width behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setColumnWidth('Sheet1', 0, 68)
    cache.rollbackColumnWidth('Sheet1', 0, undefined)
    cache.applyViewportPatch(createColumnPatch(104))

    expect(cache.getColumnWidths('Sheet1')[0]).toBeUndefined()
  })

  it('clears a pending local row height once the authoritative patch matches it', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.applyViewportPatch(createRowPatch(30))
    cache.applyViewportPatch(createRowPatch(44))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
  })

  it('keeps pending local row height visible across older authoritative patches', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.applyViewportPatch(createRowPatch(22))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(30)
    expect(cache.getRowSizes('Sheet1')[0] ?? 22).toBe(22)

    cache.applyViewportPatch(createRowPatch(30))
    cache.applyViewportPatch(createRowPatch(44))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
  })

  it('rolls back a failed local row height mutation without leaving a pending height behind', () => {
    const cache = new ProjectedViewportStore()

    cache.setRowHeight('Sheet1', 0, 30)
    cache.rollbackRowHeight('Sheet1', 0, undefined)
    cache.applyViewportPatch(createRowPatch(22))

    expect(cache.getRowHeights('Sheet1')[0]).toBeUndefined()
  })

  it('preserves hidden column metadata and collapses hidden columns from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93, true))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createColumnPatch(93, false))

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('preserves hidden row metadata and collapses hidden rows from the visible axis map', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44, true))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.applyViewportPatch(createRowPatch(44, false))

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic column hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createColumnPatch(93))
    cache.setColumnHidden('Sheet1', 0, true, 93)

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)

    cache.rollbackColumnHidden('Sheet1', 0, { hidden: false, size: 93 })

    expect(cache.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBeUndefined()
  })

  it('supports optimistic row hide and rollback using preserved raw sizes', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createRowPatch(44))
    cache.setRowHidden('Sheet1', 0, true, 44)

    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)

    cache.rollbackRowHidden('Sheet1', 0, { hidden: false, size: 44 })

    expect(cache.getRowHeights('Sheet1')[0]).toBe(44)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBeUndefined()
  })

  it('prunes back to the cache cap after the last viewport unsubscribes', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)
  })

  it('supports a configurable viewport cache cap', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: 1000,
      },
    )
    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(6010)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(1000)
  })

  it('supports fractional cache caps in viewport-store options', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: 15.4,
      },
    )
    const unsubscribe = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 9 }, () => undefined)

    cache.applyViewportPatch(createLargePatch(3, 10))

    expect(countSheetCells(cache, 'Sheet1')).toBe(30)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(15)
  })

  it('falls back to the default cache cap for non-finite values', () => {
    const cache = new ProjectedViewportStore(
      {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches: () => () => undefined,
        dispose: () => undefined,
      },
      {
        maxCachedCellsPerSheet: Number.NaN,
      },
    )
    const unsubscribe = cache.subscribeViewport(
      'Sheet1',
      {
        rowStart: 0,
        rowEnd: DEFAULT_MAX_CACHED_CELLS_PER_SHEET,
        colStart: 0,
        colEnd: 0,
      },
      () => undefined,
    )

    cache.applyViewportPatch(createLargePatch(DEFAULT_MAX_CACHED_CELLS_PER_SHEET + 1, 1))

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET + 1)

    unsubscribe()

    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)
  })

  it('keeps pinned cell subscriptions while pruning after viewport teardown', () => {
    const cache = new ProjectedViewportStore({
      invoke: async () => undefined,
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      dispose: () => undefined,
    })

    const unsubscribeViewport = cache.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 600, colStart: 0, colEnd: 9 }, () => undefined)
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['A1'], () => undefined)

    cache.applyViewportPatch(createLargePatch(601, 10))
    unsubscribeViewport()

    expect(cache.peekCell('Sheet1', 'A1')).toBeDefined()
    expect(countSheetCells(cache, 'Sheet1')).toBe(DEFAULT_MAX_CACHED_CELLS_PER_SHEET)

    unsubscribeCell()
  })

  it('allows auxiliary selection viewports to skip duplicate initial hydration', () => {
    const typedPatch: ViewportPatch = {
      ...createPatch(),
      full: true,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'future selection update', stringId: 1 },
            input: 'future selection update',
            flags: 0,
            version: 2,
          },
          displayText: 'future selection update',
          copyText: 'future selection update',
          editorText: 'future selection update',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    }
    const encodedPatch = encodeViewportPatch(typedPatch)
    let publishPatch: ((bytes: Uint8Array) => void) | null = null
    const subscribeViewportPatches = vi.fn((subscription, listener: (bytes: Uint8Array) => void) => {
      expect(subscription).toMatchObject({ initialPatch: 'none' })
      publishPatch = listener
      return () => undefined
    })
    const cache = new ProjectedViewportStore({
      invoke: vi.fn(),
      ready: async () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      subscribeViewportPatches,
      dispose: () => undefined,
    })
    const cellListener = vi.fn()

    const unsubscribeCell = cache.subscribeCell('Sheet1', 'D5', cellListener)
    const unsubscribeViewport = cache.subscribeAuxiliaryViewport(
      'Sheet1',
      { rowStart: 4, rowEnd: 4, colStart: 3, colEnd: 3 },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cellListener).not.toHaveBeenCalled()

    publishPatch?.(encodedPatch)

    expect(cache.getCell('Sheet1', 'D5').input).toBe('future selection update')
    expect(cellListener).toHaveBeenCalledTimes(1)

    unsubscribeViewport()
    unsubscribeCell()
  })

  it('notifies selected-cell listeners only for the addressed cell', () => {
    const cache = new ProjectedViewportStore()
    const selectedCellListener = vi.fn()
    const unrelatedCellListener = vi.fn()

    const unsubscribeSelected = cache.subscribeCell('Sheet1', 'A1', selectedCellListener)
    const unsubscribeUnrelated = cache.subscribeCell('Sheet1', 'B2', unrelatedCellListener)

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.Number, value: 42 },
      flags: 0,
      version: 1,
    })

    expect(selectedCellListener).toHaveBeenCalledTimes(1)
    expect(unrelatedCellListener).not.toHaveBeenCalled()

    unsubscribeSelected()
    unsubscribeUnrelated()
  })

  it('notifies only the subscribed sheet-axis channels for viewport axis patches', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()
    const columnsListener = vi.fn()
    const rowsListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)
    const unsubscribeColumns = cache.subscribeSheetChannel('Sheet1', 'columnWidths', columnsListener)
    const unsubscribeRows = cache.subscribeSheetChannel('Sheet1', 'rowHeights', rowsListener)

    cache.applyViewportPatch({
      viewport: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 },
      full: false,
      cells: [],
      styles: [],
      columns: [{ index: 1, size: 140, hidden: false }],
      rows: [],
      freezeRows: 1,
      freezeCols: 1,
    })

    expect(freezeListener).toHaveBeenCalledTimes(1)
    expect(columnsListener).toHaveBeenCalledTimes(1)
    expect(rowsListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
    unsubscribeColumns()
    unsubscribeRows()
  })

  it('exposes viewport merge metadata through the grid engine interface', () => {
    const cache = new ProjectedViewportStore()
    const mergeListener = vi.fn()
    const cellListener = vi.fn()
    const unsubscribeMerge = cache.subscribeSheetChannel('Sheet1', 'merges', mergeListener)
    const unsubscribeCell = cache.subscribeCells('Sheet1', ['C5'], cellListener)

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      cells: [],
      viewport: { sheetName: 'Sheet1', rowStart: 4, rowEnd: 4, colStart: 2, colEnd: 3 },
      merges: [{ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'D5' }],
    })

    expect(damage).toEqual([{ cell: [2, 4] }, { cell: [3, 4] }])
    expect(cache.getMergeRange('Sheet1', 'C5')).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'C5',
      endAddress: 'D5',
    })
    expect(cache.getMergeRange('Sheet1', 'D5')).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'C5',
      endAddress: 'D5',
    })
    expect(cache.listMergeRanges('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        startAddress: 'C5',
        endAddress: 'D5',
      },
    ])
    expect(mergeListener).toHaveBeenCalledTimes(1)
    expect(cellListener).toHaveBeenCalledTimes(1)

    unsubscribeMerge()
    unsubscribeCell()
  })
})
