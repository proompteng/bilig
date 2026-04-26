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
        cameraSeq: 12,
        dprBucket: 2,
        freezeRows: 2,
        freezeCols: 2,
        requestSeq: 11,
        selectedCell: { col: 8, row: 10 },
        selectedCellSnapshot: null,
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(scenes.map((scene) => scene.paneId)).toEqual(['body', 'top', 'left', 'corner'])
    expect(new Set(scenes.map((scene) => scene.generation))).toEqual(new Set([7]))
    expect(scenes.every((scene) => scene.packedScene.rects instanceof Float32Array)).toBe(true)
    expect(scenes.every((scene) => scene.packedScene.textMetrics instanceof Float32Array)).toBe(true)
    expect(scenes.every((scene) => scene.packedScene.requestSeq === 11)).toBe(true)
    expect(scenes.every((scene) => scene.packedScene.cameraSeq === 12)).toBe(true)
    expect(scenes.every((scene) => Number.isFinite(scene.packedScene.generatedAt))).toBe(true)
    expect(scenes.find((scene) => scene.paneId === 'body')?.packedScene.key).toMatchObject({
      dprBucket: 2,
      freezeVersion: expect.any(Number),
      styleVersion: 3,
      textEpoch: 0,
      valueVersion: 3,
    })
  })

  it('changes packet revision keys when axis or batch epochs change', () => {
    const axisEngine = {
      ...engine,
      getColumnAxisEntries: () => [{ id: 'col-1', index: 1, size: 144 }],
      getRowAxisEntries: () => [{ hidden: true, id: 'row-2', index: 2 }],
    }
    const first = buildWorkerResidentPaneScenes({
      engine: axisEngine,
      generation: 1,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 0, row: 0 },
        selectedCellSnapshot: null,
        selectionRange: null,
        editingCell: null,
      },
    })
    const second = buildWorkerResidentPaneScenes({
      engine: {
        ...axisEngine,
        getLastMetrics: () => ({ ...engine.getLastMetrics(), batchId: 4 }),
        getColumnAxisEntries: () => [{ id: 'col-1', index: 1, size: 160 }],
      },
      generation: 2,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 0, row: 0 },
        selectedCellSnapshot: null,
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(first[0]?.packedScene.key.axisVersionX).not.toBe(second[0]?.packedScene.key.axisVersionX)
    expect(first[0]?.packedScene.key.axisVersionY).toBe(second[0]?.packedScene.key.axisVersionY)
    expect(second[0]?.packedScene.key.valueVersion).toBe(4)
    expect(second[0]?.packedScene.key.styleVersion).toBe(4)
  })

  it('keys the worker scene cache by viewport, freeze state, selection snapshot, and editing cell', () => {
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
    const selectionOnlyKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 1, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: null,
    })
    const editingKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 1, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: null,
      editingCell: { col: 1, row: 0 },
    })
    const rangeSelectionKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      selectedCell: { col: 0, row: 0 },
      selectedCellSnapshot: null,
      selectionRange: { x: 0, y: 0, width: 2, height: 2 },
      editingCell: null,
    })

    expect(selectionOnlyKey).not.toBe(baseKey)
    expect(editingKey).not.toBe(baseKey)
    expect(rangeSelectionKey).toBe(baseKey)
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

  it('uses the engine selected-cell snapshot when it is newer than the request snapshot', () => {
    const staleSelectedCellSnapshot: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'I11',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 1,
    }
    const newerEngineSnapshot: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'I11',
      input: 'committed',
      value: { tag: ValueTag.String, value: 'committed', stringId: 1 },
      flags: 0,
      version: 2,
    }

    const scenes = buildWorkerResidentPaneScenes({
      engine: {
        ...engine,
        getCell: (_sheetName: string, address: string) => (address === 'I11' ? newerEngineSnapshot : emptyCell),
      },
      generation: 9,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 0,
        freezeCols: 0,
        selectedCell: { col: 8, row: 10 },
        selectedCellSnapshot: staleSelectedCellSnapshot,
        selectionRange: null,
        editingCell: null,
      },
    })

    expect(scenes.find((scene) => scene.paneId === 'body')?.textScene.items.some((item) => item.text === 'committed')).toBe(true)
  })

  it('keeps range selection geometry out of resident pane scenes', () => {
    const baseRequest = {
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 0,
      freezeCols: 0,
      selectedCell: { col: 1, row: 1 },
      selectedCellSnapshot: null,
      editingCell: null,
    } as const
    const unselected = buildWorkerResidentPaneScenes({
      engine,
      generation: 10,
      request: {
        ...baseRequest,
        selectionRange: null,
      },
    })
    const selected = buildWorkerResidentPaneScenes({
      engine,
      generation: 11,
      request: {
        ...baseRequest,
        selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      },
    })

    const unselectedBody = unselected.find((scene) => scene.paneId === 'body')
    const selectedBody = selected.find((scene) => scene.paneId === 'body')
    expect(selectedBody?.gpuScene.fillRects.length).toBe(unselectedBody?.gpuScene.fillRects.length)
    expect(selectedBody?.gpuScene.borderRects.length).toBe(unselectedBody?.gpuScene.borderRects.length)
  })
})
