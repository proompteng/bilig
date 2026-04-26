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
      },
    })

    expect(first[0]?.packedScene.key.axisVersionX).not.toBe(second[0]?.packedScene.key.axisVersionX)
    expect(first[0]?.packedScene.key.axisVersionY).toBe(second[0]?.packedScene.key.axisVersionY)
    expect(second[0]?.packedScene.key.valueVersion).toBe(4)
    expect(second[0]?.packedScene.key.styleVersion).toBe(4)
  })

  it('keys the worker scene cache by viewport, freeze state, dpr, and scene revision', () => {
    const baseKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
    })
    const dprKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      dprBucket: 2,
    })
    const revisionKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
      freezeRows: 1,
      freezeCols: 1,
      sceneRevision: 2,
    })
    const viewportKey = buildResidentPaneSceneCacheKey({
      sheetName: 'Sheet1',
      residentViewport: { rowStart: 0, rowEnd: 10, colStart: 1, colEnd: 11 },
      freezeRows: 1,
      freezeCols: 1,
    })

    expect(dprKey).not.toBe(baseKey)
    expect(revisionKey).not.toBe(baseKey)
    expect(viewportKey).not.toBe(baseKey)
  })

  it('renders resident text from the engine projection state', () => {
    const engineSnapshot: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'I11',
      input: 123,
      value: { tag: ValueTag.Number, value: 123 },
      flags: 0,
      version: 42,
    }

    const scenes = buildWorkerResidentPaneScenes({
      engine: {
        ...engine,
        getCell: (_sheetName: string, address: string) => (address === 'I11' ? engineSnapshot : emptyCell),
      },
      generation: 8,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 10, rowEnd: 20, colStart: 8, colEnd: 16 },
        freezeRows: 0,
        freezeCols: 0,
      },
    })

    expect(scenes.find((scene) => scene.paneId === 'body')?.packedScene.textRuns.some((item) => item.text === '123')).toBe(true)
  })

  it('keeps live selection geometry out of resident pane scenes', () => {
    const scenes = buildWorkerResidentPaneScenes({
      engine,
      generation: 10,
      request: {
        sheetName: 'Sheet1',
        residentViewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 10 },
        freezeRows: 0,
        freezeCols: 0,
      },
    })

    const body = scenes.find((scene) => scene.paneId === 'body')
    expect(body?.packedScene.fillRectCount).toBe(0)
  })
})
