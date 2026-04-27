import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { makeCellEntity } from '../entity-ids.js'
import { DynamicTopo } from '../scheduler/dynamic-topo.js'

describe('DynamicTopo', () => {
  it('handles empty affected slices and epoch rollover', () => {
    const store = new CellStore()
    store.allocate(1, 0, 0)
    const topo = new DynamicTopo()
    Object.defineProperty(topo, 'affectedEpoch', {
      configurable: true,
      value: 0xffff_fffe,
      writable: true,
    })

    const result = topo.repair(
      Uint32Array.of(0),
      {
        collectFormulaDependents: () => new Uint32Array(),
        forEachFormulaDependencyCell: () => {},
      },
      store,
      () => false,
    )

    expect(result.repaired).toBe(true)
    expect(result.orderedFormulaCount).toBe(0)
  })

  it('grows repair buffers for large dependent chains', () => {
    const store = new CellStore()
    for (let index = 0; index <= 140; index += 1) {
      store.allocate(1, 0, index)
    }

    const graph = new Map<number, Uint32Array>()
    const dependencies = new Map<number, number[]>()
    for (let index = 1; index < 140; index += 1) {
      graph.set(makeCellEntity(index), index === 139 ? Uint32Array.of(999) : Uint32Array.of(index + 1, 999))
      if (index > 1) {
        dependencies.set(index, [index - 1])
      }
    }

    const topo = new DynamicTopo()
    const result = topo.repair(
      Uint32Array.of(1),
      {
        collectFormulaDependents: (entityId) => graph.get(entityId) ?? new Uint32Array(),
        forEachFormulaDependencyCell: (cellIndex, fn) => {
          ;(dependencies.get(cellIndex) ?? []).forEach((dependencyCellIndex) => fn(dependencyCellIndex))
        },
      },
      store,
      (cellIndex) => cellIndex >= 1 && cellIndex < 140,
    )

    expect(result.repaired).toBe(true)
    expect(result.orderedFormulaCount).toBe(139)
    expect(result.orderedFormulaCellIndices[0]).toBe(1)
    expect(result.orderedFormulaCellIndices[138]).toBe(139)
    expect(store.topoRanks[1]).toBeLessThan(store.topoRanks[139])
  })

  it('repairs topo ranks for the affected formula slice while preserving external predecessors', () => {
    const store = new CellStore()
    for (let index = 0; index < 4; index += 1) {
      store.allocate(1, 0, index)
    }

    store.topoRanks[1] = 10
    store.topoRanks[2] = 1
    store.topoRanks[3] = 0

    const graph = new Map<number, Uint32Array>([[makeCellEntity(2), Uint32Array.of(3)]])
    const dependencies = new Map<number, number[]>([
      [2, [1]],
      [3, [2]],
    ])

    const topo = new DynamicTopo()
    const result = topo.repair(
      Uint32Array.of(2),
      {
        collectFormulaDependents: (entityId) => graph.get(entityId) ?? new Uint32Array(),
        forEachFormulaDependencyCell: (cellIndex, fn) => {
          ;(dependencies.get(cellIndex) ?? []).forEach((dependencyCellIndex) => fn(dependencyCellIndex))
        },
      },
      store,
      (cellIndex) => cellIndex >= 1 && cellIndex <= 3,
    )

    expect(result.repaired).toBe(true)
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([2, 3])
    expect(store.topoRanks[1]).toBeLessThan(store.topoRanks[2])
    expect(store.topoRanks[2]).toBeLessThan(store.topoRanks[3])
  })

  it('fails repair when the affected slice contains a cycle', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
    }

    const graph = new Map<number, Uint32Array>([[makeCellEntity(1), Uint32Array.of(2)]])
    const dependencies = new Map<number, number[]>([
      [1, [2]],
      [2, [1]],
    ])

    const topo = new DynamicTopo()
    const result = topo.repair(
      Uint32Array.of(1),
      {
        collectFormulaDependents: (entityId) => graph.get(entityId) ?? new Uint32Array(),
        forEachFormulaDependencyCell: (cellIndex, fn) => {
          ;(dependencies.get(cellIndex) ?? []).forEach((dependencyCellIndex) => fn(dependencyCellIndex))
        },
      },
      store,
      (cellIndex) => cellIndex === 1 || cellIndex === 2,
    )

    expect(result.repaired).toBe(false)
  })
})
