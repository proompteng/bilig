import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { buildResidentPaneSceneCacheKey, buildWorkerResidentPaneScenes } from '../worker-runtime-render-scene.js'

const emptyCell: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'A1',
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
}

const engine = {
  workbook: {
    getSheet: () => undefined,
  },
  getCell: () => emptyCell,
  getCellStyle: () => undefined,
  getColumnAxisEntries: () => [],
  getRowAxisEntries: () => [],
  getFreezePane: () => undefined,
  subscribeCells: () => () => undefined,
  getLastMetrics: () => ({
    batchId: 3,
    changedInputCount: 0,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    compileMs: 0,
  }),
} as const

describe('worker-runtime-render-scene', () => {
  it('builds resident pane scenes with caller-supplied generation', () => {
    const scenes = buildWorkerResidentPaneScenes({
      engine,
      generation: 7,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 2,
        freezeCols: 2,
        selectedCell: { col: 8, row: 10 },
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(scenes.map((scene) => scene.paneId)).toEqual(['body', 'top', 'left', 'corner'])
    expect(new Set(scenes.map((scene) => scene.generation))).toEqual(new Set([7]))
  })

  it('keys the cache by viewport, freeze state, selection, and editing cell', () => {
    const baseKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 0, row: 0 },
      selectionRange: null,
      editingCell: null,
    })
    const changedKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 1, row: 0 },
      selectionRange: null,
      editingCell: null,
    })

    expect(changedKey).not.toBe(baseKey)
  })
})
