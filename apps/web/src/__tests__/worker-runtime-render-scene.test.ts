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
        selectedCellSnapshot: null,
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(scenes.map((scene) => scene.paneId)).toEqual(['body', 'top', 'left', 'corner'])
    expect(new Set(scenes.map((scene) => scene.generation))).toEqual(new Set([7]))
    expect(scenes.every((scene) => scene.packedScene?.rects instanceof Float32Array)).toBe(true)
    expect(scenes.every((scene) => scene.packedScene?.textMetrics instanceof Float32Array)).toBe(true)
  })

  it('keys the cache by viewport, freeze state, selection, and editing cell', () => {
    const baseKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    })
    const changedKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 1, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    })

    expect(changedKey).not.toBe(baseKey)
  })

  it('uses the selected cell snapshot when worker projection has not caught up', () => {
    const selectedCellSnapshot: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'I11',
      input: 123,
      value: { tag: ValueTag.Number, value: 123 },
      flags: 0,
      version: 42,
    }

    const scenes = buildWorkerResidentPaneScenes({
      engine,
      generation: 8,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 8, row: 10 },
        selectedCellSnapshot,
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(scenes.find((scene) => scene.paneId === 'body')?.textScene.items.some((item) => item.text === '123')).toBe(true)
    expect(
      buildResidentPaneSceneCacheKey({
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 8, row: 10 },
        selectedCellSnapshot,
        selectionRange: null,
        editingCell: null,
      }),
    ).not.toBe(
      buildResidentPaneSceneCacheKey({
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 8, row: 10 },
        selectedCellSnapshot: { ...selectedCellSnapshot, version: 41 },
        selectionRange: null,
        editingCell: null,
      }),
    )
  })
})
