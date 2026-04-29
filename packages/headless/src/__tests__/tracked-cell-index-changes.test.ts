import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  type ExpandedComparativeBenchmarkWorkload,
} from '../../../benchmarks/src/expanded-competitive-workloads.js'

import {
  detachTrackedIndexChanges,
  forceMaterializeTrackedIndexChanges,
  hasDeferredTrackedIndexChanges,
  materializeTrackedIndexChangeSourcesWithMetadata,
  materializeTrackedIndexChanges,
} from '../tracked-cell-index-changes.js'

const expectedExpandedBenchmarkWorkloads: ExpandedComparativeBenchmarkWorkload[] = [
  'build-from-sheets',
  'build-dense-literals',
  'build-mixed-content',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-many-sheets',
  'rebuild-and-recalculate',
  'rebuild-config-toggle',
  'rebuild-runtime-from-snapshot',
  'single-edit-recalc',
  'single-edit-chain',
  'single-edit-fanout',
  'partial-recompute-mixed-frontier',
  'single-formula-edit-recalc',
  'batch-edit-recalc',
  'batch-edit-single-column',
  'batch-edit-multi-column',
  'batch-edit-single-column-with-undo',
  'batch-suspended-single-column',
  'batch-suspended-multi-column',
  'structural-insert-rows',
  'structural-delete-rows',
  'structural-move-rows',
  'structural-insert-columns',
  'structural-delete-columns',
  'structural-move-columns',
  'range-read',
  'range-read-dense',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-sliding-window',
  'conditional-aggregation-reused-ranges',
  'conditional-aggregation-criteria-cell-edit',
  'lookup-no-column-index',
  'lookup-with-column-index',
  'lookup-with-column-index-after-column-write',
  'lookup-with-column-index-after-batch-write',
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-after-column-write',
  'lookup-text-exact',
  'dynamic-array-filter',
]

describe('materializeTrackedIndexChanges', () => {
  it('merges explicit and recalculated slices while materializing same-sheet changes', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-merge' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'B2', 4)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const a2 = engine.workbook.getCellIndex('Sheet1', 'A2')
    const b2 = engine.workbook.getCellIndex('Sheet1', 'B2')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(a2).toBeDefined()
    expect(b2).toBeDefined()

    const changes = materializeTrackedIndexChanges(engine, Uint32Array.of(a1, a2, b1, b2), {
      explicitChangedCount: 2,
    })

    expect(changes.map((change) => change.a1)).toEqual(['A1', 'B1', 'A2', 'B2'])
  })

  it('keeps lazy same-sheet changes array-compatible and writable', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-lazy-array' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()

    const changes = materializeTrackedIndexChanges(engine, Uint32Array.of(a1, b1), { lazy: true })

    expect(Array.isArray(changes)).toBe(true)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(Object.getOwnPropertyDescriptor(changes, '0')?.writable).toBe(true)
    const markedChanges = Object.assign(changes, { customMarker: 'kept' })
    expect(Object.keys(markedChanges)).toEqual(['0', '1', 'customMarker'])

    const firstChange = changes[0]
    expect(firstChange).toMatchObject({ a1: 'A1' })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
    expect(changes[0]).toBe(firstChange)
    expect(markedChanges.customMarker).toBe('kept')
  })

  it('lazy materialization returns ordered values from borrowed typed indices', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-lazy-values' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'B2', 4)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const a2 = engine.workbook.getCellIndex('Sheet1', 'A2')
    const b2 = engine.workbook.getCellIndex('Sheet1', 'B2')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(a2).toBeDefined()
    expect(b2).toBeDefined()

    const changedCellIndices = Uint32Array.of(a1, a2, b1, b2)
    const copyTrackedCellIndices = vi.spyOn(Uint32Array, 'from')
    let changes: ReturnType<typeof materializeTrackedIndexChanges>
    try {
      changes = materializeTrackedIndexChanges(engine, changedCellIndices, { explicitChangedCount: 2, lazy: true })
      expect(copyTrackedCellIndices).not.toHaveBeenCalled()
    } finally {
      copyTrackedCellIndices.mockRestore()
    }

    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(changes.map((change) => [change.a1, change.newValue])).toEqual([
      ['A1', { tag: ValueTag.Number, value: 1 }],
      ['B1', { tag: ValueTag.Number, value: 2 }],
      ['A2', { tag: ValueTag.Number, value: 3 }],
      ['B2', { tag: ValueTag.Number, value: 4 }],
    ])
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
  })

  it('reads a high lazy same-sheet index without materializing earlier public changes', () => {
    const rowCount = 300
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-lazy-random-access' })
    engine.createSheet('Sheet1')

    const changedCellIndices = new Uint32Array(rowCount)
    for (let row = 0; row < rowCount; row += 1) {
      const a1 = `A${row + 1}`
      engine.setCellValue('Sheet1', a1, row + 1)
      const cellIndex = engine.workbook.getCellIndex('Sheet1', a1)
      expect(cellIndex).toBeDefined()
      changedCellIndices[row] = cellIndex!
    }

    const changes = materializeTrackedIndexChanges(engine, changedCellIndices, { lazy: true })

    expect(changes).toHaveLength(rowCount)
    expect(changes[rowCount - 1]).toMatchObject({
      a1: `A${rowCount}`,
      newValue: { tag: ValueTag.Number, value: rowCount },
    })

    engine.setCellValue('Sheet1', 'A1', 999)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 999 },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
  })

  it('detaches lazy same-sheet changes from later engine writes and position moves', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-detach' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 'kept')
    engine.setCellValue('Sheet1', 'C1', 3)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const c1 = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(c1).toBeDefined()

    const changedCellIndices = Uint32Array.of(a1, b1)
    const changes = materializeTrackedIndexChanges(engine, changedCellIndices, { lazy: true })

    expect(detachTrackedIndexChanges(changes)).toBe(true)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    changedCellIndices[0] = c1!
    changedCellIndices[1] = c1!

    engine.setCellValue('Sheet1', 'A1', 99)
    engine.setCellValue('Sheet1', 'B1', 'changed')
    engine.insertRows('Sheet1', 0, 1)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      address: { row: 0, col: 0 },
      newValue: { tag: ValueTag.Number, value: 1 },
    })
    expect(changes[1]).toMatchObject({
      a1: 'B1',
      address: { row: 0, col: 1 },
      newValue: { tag: ValueTag.String, value: 'kept' },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(true)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
  })

  it('fast-paths sorted disjoint tracked event sources without generic ordering', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-sources-fast' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellValue('Sheet1', 'C1', 3)
    engine.setCellValue('Sheet1', 'D1', 4)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const c1 = engine.workbook.getCellIndex('Sheet1', 'C1')
    const d1 = engine.workbook.getCellIndex('Sheet1', 'D1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(c1).toBeDefined()
    expect(d1).toBeDefined()

    const orderChanges = vi.spyOn(Array.prototype, 'toSorted')
    let result: ReturnType<typeof materializeTrackedIndexChangeSourcesWithMetadata>
    try {
      result = materializeTrackedIndexChangeSourcesWithMetadata(engine, [
        {
          changedCellIndices: Uint32Array.of(a1, b1),
          changedCellIndicesSortedDisjoint: true,
          firstChangedCellIndex: a1,
          lastChangedCellIndex: b1,
        },
        {
          changedCellIndices: Uint32Array.of(c1, d1),
          changedCellIndicesSortedDisjoint: true,
          firstChangedCellIndex: c1,
          lastChangedCellIndex: d1,
        },
      ])
      expect(orderChanges).not.toHaveBeenCalled()
    } finally {
      orderChanges.mockRestore()
    }

    expect(result).not.toBeNull()
    expect(result?.usedSortedDisjointFastPath).toBe(true)
    expect(result?.changes.map((change) => change.a1)).toEqual(['A1', 'B1', 'C1', 'D1'])
  })

  it('returns a detached lazy public array for large same-sheet source changes', () => {
    const rowCount = 300
    const split = 150
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-sources-lazy' })
    engine.createSheet('Sheet1')

    const changedCellIndices = new Uint32Array(rowCount)
    for (let row = 0; row < rowCount; row += 1) {
      const a1 = `A${row + 1}`
      engine.setCellValue('Sheet1', a1, row + 1)
      const cellIndex = engine.workbook.getCellIndex('Sheet1', a1)
      expect(cellIndex).toBeDefined()
      changedCellIndices[row] = cellIndex!
    }

    const result = materializeTrackedIndexChangeSourcesWithMetadata(engine, [
      {
        changedCellIndices: changedCellIndices.subarray(0, split),
        changedCellIndicesSortedDisjoint: true,
        firstChangedCellIndex: changedCellIndices[0],
        lastChangedCellIndex: changedCellIndices[split - 1],
      },
      {
        changedCellIndices: changedCellIndices.subarray(split),
        changedCellIndicesSortedDisjoint: true,
        firstChangedCellIndex: changedCellIndices[split],
        lastChangedCellIndex: changedCellIndices[rowCount - 1],
      },
    ])

    expect(result).not.toBeNull()
    expect(result?.usedSortedDisjointFastPath).toBe(true)
    expect(result?.changes).toHaveLength(rowCount)
    expect(result?.changes && hasDeferredTrackedIndexChanges(result.changes)).toBe(true)
    expect(result?.changes[rowCount - 1]).toMatchObject({
      a1: `A${rowCount}`,
      newValue: { tag: ValueTag.Number, value: rowCount },
    })

    engine.setCellValue('Sheet1', 'A1', 999)

    expect(result?.changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 1 },
    })
    expect(result?.changes.map((change) => change.a1).slice(0, 3)).toEqual(['A1', 'A2', 'A3'])
    expect(result?.changes.map((change) => change.a1).slice(-3)).toEqual(['A298', 'A299', 'A300'])
    expect(result?.changes && forceMaterializeTrackedIndexChanges(result.changes)).toBe(true)
    expect(result?.changes && hasDeferredTrackedIndexChanges(result.changes)).toBe(false)
  })

  it('falls back to ordered latest-cell semantics for overlapping tracked event sources', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-index-sources-overlap' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 20)
    engine.setCellValue('Sheet1', 'C1', 3)

    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')
    const c1 = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(a1).toBeDefined()
    expect(b1).toBeDefined()
    expect(c1).toBeDefined()

    const result = materializeTrackedIndexChangeSourcesWithMetadata(engine, [
      {
        changedCellIndices: Uint32Array.of(a1, c1),
        changedCellIndicesSortedDisjoint: true,
        firstChangedCellIndex: a1,
        lastChangedCellIndex: c1,
      },
      {
        changedCellIndices: Uint32Array.of(b1, c1),
        changedCellIndicesSortedDisjoint: true,
        firstChangedCellIndex: b1,
        lastChangedCellIndex: c1,
      },
    ])

    expect(result).not.toBeNull()
    expect(result?.usedSortedDisjointFastPath).toBe(false)
    expect(result?.changes.map((change) => [change.a1, change.newValue])).toEqual([
      ['A1', { tag: ValueTag.Number, value: 1 }],
      ['B1', { tag: ValueTag.Number, value: 20 }],
      ['C1', { tag: ValueTag.Number, value: 3 }],
    ])
  })

  it('keeps expanded benchmark workload definitions unchanged', () => {
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expectedExpandedBenchmarkWorkloads)
  })
})
