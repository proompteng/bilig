import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { CalcChain } from '../scheduler/calc-chain.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('CalcChain', () => {
  it('rebuilds a persistent chain from topo ranks and orders dirty cells by chain position', () => {
    const store = new CellStore()
    for (let index = 0; index < 4; index += 1) {
      store.allocate(1, 0, index)
    }

    store.topoRanks[0] = 40
    store.topoRanks[1] = 10
    store.topoRanks[2] = 30
    store.topoRanks[3] = 20

    const chain = new CalcChain()
    chain.rebuild(Uint32Array.of(0, 1, 2, 3), store)

    const first = chain.orderDirty(Uint32Array.of(0, 3), 2)
    expect(Array.from(first.orderedFormulaCellIndices.subarray(0, first.orderedFormulaCount))).toEqual([3, 0])

    store.topoRanks[0] = 5
    store.topoRanks[3] = 50
    chain.rebuild(Uint32Array.of(0, 1, 2, 3), store)

    const second = chain.orderDirty(Uint32Array.of(0, 3), 2)
    expect(Array.from(second.orderedFormulaCellIndices.subarray(0, second.orderedFormulaCount))).toEqual([0, 3])
  })

  it('drops stale formula entries when the chain is rebuilt with a smaller active set', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index
    }

    const chain = new CalcChain()
    chain.rebuild(Uint32Array.of(0, 1, 2), store)
    chain.rebuild(Uint32Array.of(1, 2), store)

    const ordered = chain.orderDirty(Uint32Array.of(2, 1), 2)
    expect(Array.from(ordered.orderedFormulaCellIndices.subarray(0, ordered.orderedFormulaCount))).toEqual([1, 2])
  })

  it('orders sparse dirty cells without scanning the full chain', () => {
    const store = new CellStore()
    for (let index = 0; index < 100; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = 100 - index
    }
    const counters = createEngineCounters()
    const chain = new CalcChain(counters)
    chain.rebuild(
      Array.from({ length: 100 }, (_, index) => index),
      store,
    )

    const ordered = chain.orderDirty(Uint32Array.of(1, 90, 10), 3)

    expect(Array.from(ordered.orderedFormulaCellIndices.subarray(0, ordered.orderedFormulaCount))).toEqual([90, 10, 1])
    expect(counters.calcChainFullScans).toBe(0)
  })

  it('handles empty and complete dirty sets without sorting work', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index
    }

    const chain = new CalcChain()
    expect(chain.hasChainFor(0)).toBe(true)

    chain.rebuild(Uint32Array.of(0, 1, 2), store)
    expect(chain.hasChainFor(3)).toBe(true)
    expect(chain.orderDirty(Uint32Array.of(), 0).orderedFormulaCount).toBe(0)

    const allDirty = chain.orderDirty(Uint32Array.of(2, 1, 0), 3)
    expect(Array.from(allDirty.orderedFormulaCellIndices.subarray(0, allDirty.orderedFormulaCount))).toEqual([0, 1, 2])
  })

  it('uses full-chain scans for dense dirty sets and grows internal buffers', () => {
    const store = new CellStore()
    for (let index = 0; index < 200; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = 199 - index
    }
    const counters = createEngineCounters()
    const chain = new CalcChain(counters)
    chain.rebuild(
      Array.from({ length: 200 }, (_, index) => index),
      store,
    )

    const dirty = Uint32Array.from(Array.from({ length: 80 }, (_, offset) => offset * 2))
    const ordered = chain.orderDirty(dirty, dirty.length)

    expect(ordered.orderedFormulaCount).toBe(80)
    expect(Array.from(ordered.orderedFormulaCellIndices.subarray(0, 3))).toEqual([158, 156, 154])
    expect(Array.from(ordered.orderedFormulaCellIndices.subarray(77, 80))).toEqual([4, 2, 0])
    expect(counters.calcChainFullScans).toBe(1)
  })

  it('resets the dense-dirty epoch when the marker reaches the uint32 limit', () => {
    const store = new CellStore()
    for (let index = 0; index < 80; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index
    }
    const chain = new CalcChain()
    chain.rebuild(
      Array.from({ length: 80 }, (_, index) => index),
      store,
    )
    Reflect.set(chain, 'dirtyEpoch', 0xffff_fffe)

    const dirty = Uint32Array.from(Array.from({ length: 70 }, (_, index) => index))
    expect(chain.orderDirty(dirty, dirty.length).orderedFormulaCount).toBe(70)
    expect(chain.orderDirty(dirty, dirty.length).orderedFormulaCount).toBe(70)
  })
})
