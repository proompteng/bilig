import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { makeCellEntity, makeRangeEntity } from '../entity-ids.js'
import { RecalcScheduler } from '../scheduler.js'

describe('RecalcScheduler', () => {
  it('orders dirty formulas by topo rank while traversing range nodes', () => {
    const store = new CellStore()
    for (let index = 0; index < 4; index += 1) {
      store.allocate(1, 0, index)
    }
    store.topoRanks[1] = 5
    store.topoRanks[3] = 2

    const rangeEntity = makeRangeEntity(0)
    const graph = new Map<number, Uint32Array>([
      [makeCellEntity(0), Uint32Array.from([makeCellEntity(1), rangeEntity])],
      [rangeEntity, Uint32Array.from([makeCellEntity(3)])],
    ])

    const scheduler = new RecalcScheduler()
    const result = scheduler.collectDirty(
      [0],
      {
        getDependents: (entityId) => graph.get(entityId) ?? new Uint32Array(),
      },
      store,
      (cellIndex) => cellIndex === 1 || cellIndex === 3,
      1,
    )

    expect(result.orderedFormulaCount).toBe(2)
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([3, 1])
    expect(result.rangeNodeVisits).toBe(1)
  })

  it('returns zero dirty formulas without allocating a new result buffer', () => {
    const store = new CellStore()
    store.allocate(1, 0, 0)

    const scheduler = new RecalcScheduler()
    const first = scheduler.collectDirty([0], { getDependents: () => new Uint32Array() }, store, () => false, 0)
    const second = scheduler.collectDirty([0], { getDependents: () => new Uint32Array() }, store, () => false, 0)

    expect(first.orderedFormulaCount).toBe(0)
    expect(second.orderedFormulaCount).toBe(0)
    expect(second.orderedFormulaCellIndices).toBe(first.orderedFormulaCellIndices)
  })

  it('grows queue, range, and rank buffers while deduping repeated roots and ranges', () => {
    const store = new CellStore()
    for (let index = 0; index < 140; index += 1) {
      store.allocate(1, 0, index)
    }
    store.topoRanks[10] = 0
    store.topoRanks[131] = 3
    store.topoRanks[130] = 130

    const rangeEntity = makeRangeEntity(90)
    const graph = new Map<number, Uint32Array>([
      [makeCellEntity(0), Uint32Array.from([rangeEntity, rangeEntity, makeCellEntity(10)])],
      [rangeEntity, Uint32Array.from([makeCellEntity(131), makeCellEntity(130)])],
    ])

    const scheduler = new RecalcScheduler()
    const result = scheduler.collectDirty(
      [0, 0],
      {
        getDependents: (entityId) => graph.get(entityId) ?? new Uint32Array(),
      },
      store,
      (cellIndex) => cellIndex === 10 || cellIndex === 130 || cellIndex === 131,
      128,
    )

    expect(result.orderedFormulaCount).toBe(3)
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([10, 131, 130])
    expect(result.rangeNodeVisits).toBe(1)
  })
})
