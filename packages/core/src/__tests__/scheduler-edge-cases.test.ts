import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import {
  entityPayload,
  isCellEntity,
  isExactLookupColumnEntity,
  isRangeEntity,
  isSortedLookupColumnEntity,
  makeCellEntity,
  makeExactLookupColumnEntity,
  makeRangeEntity,
  makeSortedLookupColumnEntity,
} from '../entity-ids.js'
import { RecalcScheduler } from '../scheduler.js'

describe('RecalcScheduler edge cases', () => {
  it('returns already-ordered dirty roots without resorting when ranks stay monotonic', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index + 1
    }

    const scheduler = new RecalcScheduler()
    const result = scheduler.collectDirty(Uint32Array.of(0, 1, 2), { getDependents: () => new Uint32Array() }, store, () => true, 0)

    expect(result.orderedFormulaCount).toBe(3)
    expect(result.orderedFormulaCellIndices).toBeDefined()
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, 3))).toEqual([0, 1, 2])
  })

  it('dedupes exact and sorted lookup columns independently across repeated traversals', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index
    }

    const exactEntity = makeExactLookupColumnEntity(7, 3)
    const sortedEntity = makeSortedLookupColumnEntity(7, 3)
    const graph = new Map<number, Uint32Array>([
      [makeCellEntity(0), Uint32Array.of(exactEntity, exactEntity, sortedEntity, sortedEntity)],
      [exactEntity, Uint32Array.of(makeCellEntity(1))],
      [sortedEntity, Uint32Array.of(makeCellEntity(2))],
    ])

    const getDependents = (entityId: number): Uint32Array => graph.get(entityId) ?? new Uint32Array()
    const scheduler = new RecalcScheduler()

    const first = scheduler.collectDirty(Uint32Array.of(0), { getDependents }, store, (cellIndex) => cellIndex > 0, 0)
    const second = scheduler.collectDirty(Uint32Array.of(0), { getDependents }, store, (cellIndex) => cellIndex > 0, 0)

    expect(Array.from(first.orderedFormulaCellIndices.subarray(0, first.orderedFormulaCount))).toEqual([1, 2])
    expect(Array.from(second.orderedFormulaCellIndices.subarray(0, second.orderedFormulaCount))).toEqual([1, 2])
  })

  it('encodes and classifies cell, range, and lookup entities consistently', () => {
    const cell = makeCellEntity(12)
    const range = makeRangeEntity(19)
    const exact = makeExactLookupColumnEntity(2, 5)
    const sorted = makeSortedLookupColumnEntity(2, 5)

    expect(isCellEntity(cell)).toBe(true)
    expect(isRangeEntity(cell)).toBe(false)
    expect(entityPayload(cell)).toBe(12)

    expect(isRangeEntity(range)).toBe(true)
    expect(isExactLookupColumnEntity(range)).toBe(false)
    expect(entityPayload(range)).toBe(19)

    expect(isExactLookupColumnEntity(exact)).toBe(true)
    expect(isSortedLookupColumnEntity(exact)).toBe(false)
    expect(entityPayload(exact)).toBe(entityPayload(sorted))

    expect(isSortedLookupColumnEntity(sorted)).toBe(true)
    expect(isExactLookupColumnEntity(sorted)).toBe(false)
  })

  it('resets lookup epochs at overflow and skips non-formula dependents', () => {
    const store = new CellStore()
    for (let index = 0; index < 3; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index
    }

    const scheduler = new RecalcScheduler()
    Reflect.set(scheduler, 'exactLookupEpoch', 0xffff_fffe)
    Reflect.set(scheduler, 'sortedLookupEpoch', 0xffff_fffe)
    const visitedExactLookupColumns = readMapField(scheduler, 'visitedExactLookupColumns')
    const visitedSortedLookupColumns = readMapField(scheduler, 'visitedSortedLookupColumns')
    visitedExactLookupColumns.set(10, 99)
    visitedSortedLookupColumns.set(11, 99)

    const graph = new Map<number, Uint32Array>([[makeCellEntity(0), Uint32Array.of(makeCellEntity(1), makeCellEntity(2))]])
    const result = scheduler.collectDirty(
      Uint32Array.of(0),
      { getDependents: (entityId: number) => graph.get(entityId) ?? new Uint32Array() },
      store,
      (cellIndex: number) => cellIndex === 2,
      0,
    )

    expect(readNumericField(scheduler, 'exactLookupEpoch')).toBe(1)
    expect(readNumericField(scheduler, 'sortedLookupEpoch')).toBe(1)
    expect(visitedExactLookupColumns.size).toBe(0)
    expect(visitedSortedLookupColumns.size).toBe(0)
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([2])
  })

  it('grows the entity queue when traversal adds more non-cell entities than the store size', () => {
    const store = new CellStore()
    store.allocate(1, 0, 0)

    const graph = new Map<number, Uint32Array>([
      [makeCellEntity(0), Uint32Array.from(Array.from({ length: 200 }, (_, index) => makeRangeEntity(index)))],
    ])

    const scheduler = new RecalcScheduler()
    const result = scheduler.collectDirty(
      Uint32Array.of(0),
      { getDependents: (entityId) => graph.get(entityId) ?? new Uint32Array() },
      store,
      () => false,
      256,
    )

    expect(result.orderedFormulaCount).toBe(0)
    expect(result.rangeNodeVisits).toBe(200)
  })

  it('grows rank buffers when many dirty formulas require counting sort', () => {
    const store = new CellStore()
    for (let index = 0; index < 130; index += 1) {
      store.allocate(1, 0, index)
      store.topoRanks[index] = index * 2
    }

    const scheduler = new RecalcScheduler()
    const changedRoots = Uint32Array.from(Array.from({ length: 130 }, (_, index) => 129 - index))
    const result = scheduler.collectDirty(changedRoots, { getDependents: () => new Uint32Array() }, store, () => true, 0)

    expect(result.orderedFormulaCount).toBe(130)
    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, 5))).toEqual([0, 1, 2, 3, 4])
    expect(Array.from(result.orderedFormulaCellIndices.subarray(125, 130))).toEqual([125, 126, 127, 128, 129])
  })
})

function readNumericField(target: object, key: string): number {
  const value = Reflect.get(target, key)
  if (typeof value !== 'number') {
    throw new Error(`Expected numeric field ${key}`)
  }
  return value
}

function readMapField(target: object, key: string): Map<unknown, unknown> {
  const value = Reflect.get(target, key)
  if (!(value instanceof Map)) {
    throw new Error(`Expected map field ${key}`)
  }
  return value
}
