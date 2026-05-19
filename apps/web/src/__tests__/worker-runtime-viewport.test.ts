import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot, type CellStyleRecord, type EngineEvent, type RecalcMetrics } from '@bilig/protocol'
import { WorkerViewportPatchPublisher } from '../worker-runtime-viewport-publisher.js'
import type { ViewportSubscriptionState, WorkerEngine } from '../worker-runtime-support.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 1,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

const STYLE_ID = 'style-live'
const CELL: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'B2',
  value: { tag: ValueTag.Empty },
  flags: 0,
  styleId: STYLE_ID,
  version: 1,
}

function createStyle(backgroundColor: string): CellStyleRecord {
  return {
    id: STYLE_ID,
    fill: { backgroundColor },
  }
}

function createEvent(): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics: TEST_METRICS,
  }
}

function createSubscriptionState(): ViewportSubscriptionState {
  return {
    subscription: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 5,
      colStart: 0,
      colEnd: 5,
    },
    listener: vi.fn(),
    nextVersion: 1,
    knownStyleIds: new Set(),
    lastStyleSignatures: new Map(),
    lastCellSignatures: new Map(),
    lastColumnSignatures: new Map(),
    lastRowSignatures: new Map(),
    lastMergeSignatures: new Map(),
  }
}

describe('worker runtime viewport patches', () => {
  it('includes affected cells when a referenced style record changes under the same style id', () => {
    let currentStyle = createStyle('#00ff00')
    const engine: WorkerEngine = {
      workbook: {
        workbookName: 'viewport-style-proof',
        cellStore: {
          sheetIds: new Uint16Array(),
          rows: new Uint32Array(),
          cols: new Uint16Array(),
        },
        sheetsByName: new Map(),
        getSheet: () => ({ name: 'Sheet1', order: 0, grid: { forEachCellEntry: () => undefined } }),
        getSheetNameById: () => 'Sheet1',
        getQualifiedAddress: () => 'Sheet1!B2',
      },
      ready: async () => undefined,
      createSheet: () => undefined,
      subscribe: () => () => undefined,
      subscribeBatches: () => () => undefined,
      getLastMetrics: () => TEST_METRICS,
      getSyncState: () => 'local-only',
      getCell: () => CELL,
      getCellStyle: () => currentStyle,
      setRangeNumberFormat: () => undefined,
      clearRangeNumberFormat: () => undefined,
      clearRange: () => undefined,
      setCellValue: () => undefined,
      setCellFormula: () => undefined,
      setRangeStyle: () => undefined,
      clearRangeStyle: () => undefined,
      clearCell: () => undefined,
      undo: () => false,
      redo: () => false,
      canUndo: () => false,
      canRedo: () => false,
      renderCommit: () => undefined,
      fillRange: () => undefined,
      copyRange: () => undefined,
      moveRange: () => undefined,
      insertRows: () => undefined,
      deleteRows: () => undefined,
      insertColumns: () => undefined,
      deleteColumns: () => undefined,
      updateRowMetadata: () => undefined,
      updateColumnMetadata: () => undefined,
      setFreezePane: () => undefined,
      getFreezePane: () => undefined,
      mergeCells: () => undefined,
      unmergeCells: () => false,
      getMergeRange: () => undefined,
      listMergeRanges: () => [],
      exportSnapshot: () => ({
        version: 1,
        workbook: { name: 'viewport-style-proof' },
        sheets: [{ name: 'Sheet1', order: 0, cells: [] }],
      }),
      exportReplicaSnapshot: () => ({
        replica: {
          replicaId: 'viewport-style-proof',
          counter: 0,
          appliedBatchIds: [],
        },
        entityVersions: [],
        sheetDeleteVersions: [],
      }),
      importSnapshot: () => undefined,
      importReplicaSnapshot: () => undefined,
      getColumnAxisEntries: () => [],
      getRowAxisEntries: () => [],
    }
    const publisher = new WorkerViewportPatchPublisher({
      buildPatch: () => {
        throw new Error('test uses publisher.buildPatch directly')
      },
      getAuthoritativeRevision: () => 0,
      getCurrentMetrics: () => TEST_METRICS,
      getProjectionEngine: () => engine,
      hasProjectionEngine: () => true,
    })
    const state = createSubscriptionState()

    const initialPatch = publisher.buildPatch(state, null, TEST_METRICS, 0, null)
    expect(initialPatch.cells.map((cell) => cell.snapshot.address)).toContain('B2')
    expect(initialPatch.styles).toContainEqual(createStyle('#00ff00'))

    currentStyle = createStyle('#0000ff')
    const styleOnlyEventPatch = publisher.buildPatch(state, createEvent(), TEST_METRICS, 0, {
      changedCells: null,
      invalidatedRanges: [{ rowStart: 1, rowEnd: 1, colStart: 1, colEnd: 1 }],
      invalidatedRows: [],
      invalidatedColumns: [],
    })

    expect(styleOnlyEventPatch.styles).toEqual([createStyle('#0000ff')])
    expect(styleOnlyEventPatch.cells).toEqual([
      expect.objectContaining({
        row: 1,
        col: 1,
        snapshot: expect.objectContaining({
          address: 'B2',
          styleId: STYLE_ID,
        }),
        styleId: STYLE_ID,
      }),
    ])
  })
})
