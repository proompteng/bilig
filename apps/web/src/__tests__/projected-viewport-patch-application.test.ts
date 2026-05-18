import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type RecalcMetrics } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import { applyProjectedViewportPatch, type ProjectedViewportPatchState } from '../projected-viewport-patch-application.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

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

function createPatchState(): ProjectedViewportPatchState {
  return {
    cellSnapshots: new Map(),
    cellKeysBySheet: new Map(),
    cellStyles: new Map([['style-0', { id: 'style-0' }]]),
    columnSizesBySheet: new Map(),
    columnWidthsBySheet: new Map(),
    pendingColumnWidthsBySheet: new Map(),
    pendingHiddenColumnsBySheet: new Map(),
    rowSizesBySheet: new Map(),
    rowHeightsBySheet: new Map(),
    pendingRowHeightsBySheet: new Map(),
    pendingHiddenRowsBySheet: new Map(),
    hiddenColumnsBySheet: new Map(),
    hiddenRowsBySheet: new Map(),
    freezeRowsBySheet: new Map(),
    freezeColsBySheet: new Map(),
    mergeRangesBySheet: new Map(),
    knownSheets: new Set(),
  }
}

describe('applyProjectedViewportPatch', () => {
  it('reports damage when a style record changes without a newer cell snapshot', () => {
    const state = createPatchState()

    applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch('style-fill'),
        styles: [{ id: 'style-fill', fill: { backgroundColor: '#c9daf8' } }],
      },
    })

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch('style-fill'),
        styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
      },
    })

    expect(result.damage).toEqual([{ cell: [3, 4] }])
    expect(state.cellStyles.get('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#a4c2f4' },
    })
  })

  it('preserves top-level patch style ids for styled empty cells', () => {
    const state = createPatchState()

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        styles: [{ id: 'style-fill', fill: { backgroundColor: '#00ff00' } }],
        cells: createPatch().cells.map((cell) => Object.assign({}, cell, { styleId: 'style-fill' })),
      },
    })

    expect(result.damage).toEqual([{ cell: [3, 4] }])
    expect(state.cellSnapshots.get('Sheet1!D5')?.styleId).toBe('style-fill')
    expect(state.cellStyles.get('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#00ff00' },
    })

    const clearResult = applyProjectedViewportPatch({
      state,
      patch: createPatch(),
    })

    expect(clearResult.damage).toEqual([{ cell: [3, 4] }])
    expect(state.cellSnapshots.get('Sheet1!D5')?.styleId).toBeUndefined()
  })

  it('clears stale viewport cells on full patches without dropping cells outside the viewport', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!A1', {
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.String, value: 'pinned', stringId: 1 },
      flags: 0,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!A1']))

    applyProjectedViewportPatch({
      state,
      patch: { ...createPatch(), full: true },
    })

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        full: true,
        cells: [],
      },
    })

    expect(result.damage).toEqual([{ cell: [3, 4] }])
    expect(state.cellSnapshots.get('Sheet1!D5')).toBeUndefined()
    expect(state.cellSnapshots.get('Sheet1!A1')?.value).toEqual({
      tag: ValueTag.String,
      value: 'pinned',
      stringId: 1,
    })
  })

  it('tracks freeze metadata and axis changes from a viewport patch', () => {
    const state = createPatchState()

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        freezeRows: 2,
        freezeCols: 1,
        cells: [],
        columns: [{ index: 0, size: 93, hidden: true }],
        rows: [{ index: 0, size: 44, hidden: false }],
      },
    })

    expect(result.freezeChanged).toBe(true)
    expect(result.axisChanged).toBe(true)
    expect(state.freezeRowsBySheet.get('Sheet1')).toBe(2)
    expect(state.freezeColsBySheet.get('Sheet1')).toBe(1)
    expect(state.columnWidthsBySheet.get('Sheet1')?.[0]).toBe(0)
    expect(state.hiddenColumnsBySheet.get('Sheet1')?.[0]).toBe(true)
    expect(state.rowHeightsBySheet.get('Sheet1')?.[0]).toBe(44)
  })

  it('tracks merge metadata and damages the exact merged cells in the viewport', () => {
    const state = createPatchState()

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        cells: [],
        merges: [{ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'D5' }],
      },
    })

    expect(result.mergesChanged).toBe(true)
    expect(result.damage).toEqual([{ cell: [2, 4] }, { cell: [3, 4] }])
    expect(state.mergeRangesBySheet.get('Sheet1')?.size).toBe(1)
  })

  it('clears intersecting merge metadata when an unmerge patch sends an explicit empty set', () => {
    const state = createPatchState()
    applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        cells: [],
        merges: [{ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'D5' }],
      },
    })

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        version: 2,
        cells: [],
        merges: [],
      },
    })

    expect(result.mergesChanged).toBe(true)
    expect(result.damage).toEqual([{ cell: [2, 4] }, { cell: [3, 4] }])
    expect(state.mergeRangesBySheet.get('Sheet1')).toBeUndefined()
  })

  it('does not let stale viewport patches override an optimistic hidden row', () => {
    const state = createPatchState()
    state.rowSizesBySheet.set('Sheet1', { 1: 22 })
    state.rowHeightsBySheet.set('Sheet1', { 1: 0 })
    state.hiddenRowsBySheet.set('Sheet1', { 1: true })
    state.pendingHiddenRowsBySheet.set('Sheet1', { 1: true })

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        cells: [],
        rows: [{ index: 1, size: 22, hidden: false }],
      },
    })

    expect(result.rowsChanged).toBe(false)
    expect(state.rowSizesBySheet.get('Sheet1')?.[1]).toBe(22)
    expect(state.rowHeightsBySheet.get('Sheet1')?.[1]).toBe(0)
    expect(state.hiddenRowsBySheet.get('Sheet1')?.[1]).toBe(true)
    expect(state.pendingHiddenRowsBySheet.get('Sheet1')?.[1]).toBe(true)
  })

  it('clears pending hidden row state when the authoritative patch catches up', () => {
    const state = createPatchState()
    state.rowSizesBySheet.set('Sheet1', { 1: 22 })
    state.rowHeightsBySheet.set('Sheet1', { 1: 0 })
    state.hiddenRowsBySheet.set('Sheet1', { 1: true })
    state.pendingHiddenRowsBySheet.set('Sheet1', { 1: true })

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        cells: [],
        rows: [{ index: 1, size: 22, hidden: true }],
      },
    })

    expect(result.rowsChanged).toBe(false)
    expect(state.rowHeightsBySheet.get('Sheet1')?.[1]).toBe(0)
    expect(state.hiddenRowsBySheet.get('Sheet1')?.[1]).toBe(true)
    expect(state.pendingHiddenRowsBySheet.get('Sheet1')).toBeUndefined()
  })

  it('does not report a freeze change when a patch confirms the default unfrozen state', () => {
    const state = createPatchState()

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        ...createPatch(),
        freezeRows: 0,
        freezeCols: 0,
      },
    })

    expect(result.freezeChanged).toBe(false)
    expect(state.freezeRowsBySheet.get('Sheet1')).toBe(0)
    expect(state.freezeColsBySheet.get('Sheet1')).toBe(0)
  })

  it('accepts reset empty snapshots that clear stale cached cells', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.String, value: 'left', stringId: 1 },
      flags: 16,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    })
  })

  it('keeps newer authored snapshots when stale reset empty patches arrive', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      formula: 'A1="HELLO"',
      value: { tag: ValueTag.String, value: '=A1="HELLO"', stringId: 1 },
      flags: 0,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      formula: 'A1="HELLO"',
      value: { tag: ValueTag.String, value: '=A1="HELLO"', stringId: 1 },
      version: 1,
    })
  })

  it('keeps optimistic value snapshots when lagging reset-empty patches arrive', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('accepts authoritative empty patches over stale optimistic value snapshots', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: 'deleted-by-history',
      value: { tag: ValueTag.String, value: 'deleted-by-history', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    })
  })

  it('keeps optimistic clear snapshots when lagging non-empty patches arrive', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: 'before-delete',
              value: { tag: ValueTag.String, value: 'before-delete', stringId: 1 },
              flags: 0,
              version: 7,
            },
            displayText: 'before-delete',
            copyText: 'before-delete',
            editorText: 'before-delete',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
  })

  it('does not resurrect stale content after a reset-empty patch confirms an optimistic clear', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 10,
        version: 10,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: 'before-delete',
              value: { tag: ValueTag.String, value: 'before-delete', stringId: 1 },
              flags: 0,
              version: 7,
            },
            displayText: 'before-delete',
            copyText: 'before-delete',
            editorText: 'before-delete',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 8,
    })
  })

  it('does not resurrect same-version stale content after a clear is confirmed', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              value: { tag: ValueTag.Empty },
              flags: 0,
              version: 9,
            },
            displayText: '',
            copyText: '',
            editorText: '',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: 'same-version-stale-content',
              value: { tag: ValueTag.String, value: 'same-version-stale-content', stringId: 1 },
              flags: 0,
              version: 9,
            },
            displayText: 'same-version-stale-content',
            copyText: 'same-version-stale-content',
            editorText: 'same-version-stale-content',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 9,
    })
  })

  it('keeps optimistic clears when newer non-empty patches arrive before clear confirmation', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: 'new-authoritative-value',
              value: { tag: ValueTag.String, value: 'new-authoritative-value', stringId: 1 },
              flags: 0,
              version: 9,
            },
            displayText: 'new-authoritative-value',
            copyText: 'new-authoritative-value',
            editorText: 'new-authoritative-value',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
  })

  it('releases optimistic clear protection after authoritative empty confirmation', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              value: { tag: ValueTag.Empty },
              flags: 0,
              version: 9,
            },
            displayText: '',
            copyText: '',
            editorText: '',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 9,
    })

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 10,
        version: 10,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: 'new-after-confirmed-clear',
              value: { tag: ValueTag.String, value: 'new-after-confirmed-clear', stringId: 1 },
              flags: 0,
              version: 10,
            },
            displayText: 'new-after-confirmed-clear',
            copyText: 'new-after-confirmed-clear',
            editorText: 'new-after-confirmed-clear',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      input: 'new-after-confirmed-clear',
      value: { tag: ValueTag.String, value: 'new-after-confirmed-clear', stringId: 1 },
      flags: 0,
      version: 10,
    })
  })

  it('releases optimistic clear protection when an authoritative full patch omits the cleared cell', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 9,
        version: 9,
        full: true,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 8,
    })
  })

  it('keeps optimistic invalid-formula errors when lagging reset-empty patches arrive', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: '=1+',
      value: { tag: ValueTag.Error, code: ErrorCode.Value },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
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
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      input: '=1+',
      value: { tag: ValueTag.Error, code: ErrorCode.Value },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('keeps optimistic value snapshots when full lagging patches omit the cell', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: true,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('removes optimistic value snapshots when an authoritative full patch omits the cell', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: 'undo-removed',
      value: { tag: ValueTag.String, value: 'undo-removed', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    const result = applyProjectedViewportPatch({
      state,
      patch: {
        authoritativeRevision: 2,
        version: 2,
        full: true,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.has('Sheet1!B2')).toBe(false)
    expect(state.cellKeysBySheet.get('Sheet1')?.has('Sheet1!B2')).toBe(false)
    expect(result.changedKeys.has('Sheet1!B2')).toBe(true)
    expect(result.damage).toEqual([{ cell: [1, 1] }])
  })

  it('accepts matching value patches while preserving optimistic protection', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              input: '12',
              value: { tag: ValueTag.String, value: '12', stringId: 2 },
              flags: 0,
              version: 5,
            },
            displayText: '12',
            copyText: '12',
            editorText: '12',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      input: '12',
      value: { tag: ValueTag.String, value: '12', stringId: 2 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 5,
    })
  })

  it('accepts newer empty patches over optimistic snapshots for structural clears', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      input: 'row-2',
      value: { tag: ValueTag.String, value: 'row-2', stringId: 1 },
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              value: { tag: ValueTag.Empty },
              flags: 0,
              version: 3,
            },
            displayText: '',
            copyText: '',
            editorText: '',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 3,
    })
  })

  it('accepts formula error snapshots over optimistic formula text', () => {
    const state = createPatchState()
    state.cellSnapshots.set('Sheet1!B2', {
      sheetName: 'Sheet1',
      address: 'B2',
      formula: '1+',
      value: { tag: ValueTag.String, value: '=1+', stringId: 1 },
      flags: 0,
      version: 1,
    })
    state.cellKeysBySheet.set('Sheet1', new Set(['Sheet1!B2']))

    applyProjectedViewportPatch({
      state,
      patch: {
        version: 2,
        full: false,
        freezeRows: 0,
        freezeCols: 0,
        viewport: {
          sheetName: 'Sheet1',
          rowStart: 1,
          rowEnd: 1,
          colStart: 1,
          colEnd: 1,
        },
        metrics: TEST_METRICS,
        styles: [],
        cells: [
          {
            row: 1,
            col: 1,
            snapshot: {
              sheetName: 'Sheet1',
              address: 'B2',
              value: { tag: ValueTag.Error, code: ErrorCode.Value },
              flags: 0,
              version: 2,
            },
            displayText: '#VALUE!',
            copyText: '#VALUE!',
            editorText: '#VALUE!',
            formatId: 0,
            styleId: 'style-0',
          },
        ],
        columns: [],
        rows: [],
      },
    })

    expect(state.cellSnapshots.get('Sheet1!B2')).toMatchObject({
      value: { tag: ValueTag.Error, code: ErrorCode.Value },
      version: 2,
    })
    expect(state.cellSnapshots.get('Sheet1!B2')).not.toHaveProperty('formula')
  })
})
