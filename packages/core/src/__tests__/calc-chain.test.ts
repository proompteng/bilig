import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { CalcChain } from '../scheduler/calc-chain.js'

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
})
