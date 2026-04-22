import { describe, expect, it } from 'vitest'
import { createFormulaFamilyStore } from '../formula/formula-family-store.js'

describe('FormulaFamilyStore', () => {
  it('groups copied row formulas into one family run', () => {
    const store = createFormulaFamilyStore()

    for (let row = 0; row < 10_000; row += 1) {
      store.upsertFormula({
        cellIndex: row,
        sheetId: 1,
        row,
        col: 4,
        templateId: 7,
        shapeKey: 'relative-add',
      })
    }

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 1, memberCount: 10_000 })
    expect(store.listFamilies()[0]?.runs).toEqual([
      expect.objectContaining({
        axis: 'row',
        fixedIndex: 4,
        start: 0,
        end: 9_999,
        cellIndices: expect.arrayContaining([0, 9_999]),
      }),
    ])
  })

  it('keeps different template or shape keys in separate families', () => {
    const store = createFormulaFamilyStore()

    store.upsertFormula({ cellIndex: 1, sheetId: 1, row: 0, col: 2, templateId: 1, shapeKey: 'a' })
    store.upsertFormula({ cellIndex: 2, sheetId: 1, row: 1, col: 2, templateId: 2, shapeKey: 'a' })
    store.upsertFormula({ cellIndex: 3, sheetId: 1, row: 2, col: 2, templateId: 1, shapeKey: 'b' })

    expect(store.getStats()).toEqual({ familyCount: 3, runCount: 3, memberCount: 3 })
  })

  it('groups periodically copied row formulas into strided family runs', () => {
    const store = createFormulaFamilyStore()

    for (let row = 1; row < 10; row += 3) {
      store.upsertFormula({
        cellIndex: row + 100,
        sheetId: 1,
        row,
        col: 5,
        templateId: 7,
        shapeKey: 'mixed-template',
      })
    }

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 1, memberCount: 3 })
    expect(store.listFamilies()[0]?.runs).toEqual([
      expect.objectContaining({
        axis: 'row',
        fixedIndex: 5,
        start: 1,
        end: 7,
        step: 3,
        cellIndices: [101, 104, 107],
      }),
    ])
  })

  it('reshapes a sparse two-member run when the missing midpoint arrives later', () => {
    const store = createFormulaFamilyStore()

    store.upsertFormula({ cellIndex: 101, sheetId: 1, row: 1, col: 5, templateId: 7, shapeKey: 'mixed-template' })
    store.upsertFormula({ cellIndex: 107, sheetId: 1, row: 7, col: 5, templateId: 7, shapeKey: 'mixed-template' })
    store.upsertFormula({ cellIndex: 104, sheetId: 1, row: 4, col: 5, templateId: 7, shapeKey: 'mixed-template' })

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 1, memberCount: 3 })
    expect(store.listFamilies()[0]?.runs).toEqual([
      expect.objectContaining({
        axis: 'row',
        fixedIndex: 5,
        start: 1,
        end: 7,
        step: 3,
        cellIndices: [101, 104, 107],
      }),
    ])
  })

  it('keeps column-oriented singleton merges available through the run index', () => {
    const store = createFormulaFamilyStore()

    store.upsertFormula({ cellIndex: 201, sheetId: 1, row: 2, col: 1, templateId: 7, shapeKey: 'horizontal-template' })
    store.upsertFormula({ cellIndex: 204, sheetId: 1, row: 2, col: 4, templateId: 7, shapeKey: 'horizontal-template' })
    store.upsertFormula({ cellIndex: 207, sheetId: 1, row: 2, col: 7, templateId: 7, shapeKey: 'horizontal-template' })

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 1, memberCount: 3 })
    expect(store.listFamilies()[0]?.runs).toEqual([
      expect.objectContaining({
        axis: 'column',
        fixedIndex: 2,
        start: 1,
        end: 7,
        step: 3,
        cellIndices: [201, 204, 207],
      }),
    ])
  })

  it('preserves a strided run shape when a local edit removes a middle member', () => {
    const store = createFormulaFamilyStore()

    for (let row = 1; row <= 10; row += 3) {
      store.upsertFormula({
        cellIndex: row + 100,
        sheetId: 1,
        row,
        col: 5,
        templateId: 7,
        shapeKey: 'mixed-template',
      })
    }

    expect(store.unregisterFormula(107)).toBe(true)

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 2, memberCount: 3 })
    expect(store.listFamilies()[0]?.runs).toEqual([
      expect.objectContaining({
        axis: 'row',
        fixedIndex: 5,
        start: 1,
        end: 4,
        step: 3,
        cellIndices: [101, 104],
      }),
      expect.objectContaining({
        axis: 'row',
        fixedIndex: 5,
        start: 10,
        end: 10,
        step: 1,
        cellIndices: [110],
      }),
    ])
  })

  it('splits a family run when a local edit removes a middle member', () => {
    const store = createFormulaFamilyStore()
    for (let row = 0; row < 5; row += 1) {
      store.upsertFormula({ cellIndex: row + 10, sheetId: 1, row, col: 3, templateId: 1, shapeKey: 'a' })
    }

    expect(store.unregisterFormula(12)).toBe(true)

    expect(store.getStats()).toEqual({ familyCount: 1, runCount: 2, memberCount: 4 })
    expect(store.listFamilies()[0]?.runs.map((run) => [run.start, run.end, run.cellIndices])).toEqual([
      [0, 1, [10, 11]],
      [3, 4, [13, 14]],
    ])
  })

  it('invalidates only formulas intersecting a structural span', () => {
    const store = createFormulaFamilyStore()
    for (let row = 0; row < 6; row += 1) {
      store.upsertFormula({ cellIndex: row, sheetId: 1, row, col: 2, templateId: 1, shapeKey: 'a' })
    }
    store.upsertFormula({ cellIndex: 100, sheetId: 2, row: 3, col: 2, templateId: 1, shapeKey: 'a' })

    store.applyStructuralInvalidation({ sheetId: 1, axis: 'row', start: 2, end: 4 })

    expect(store.getMembership(0)).toBeDefined()
    expect(store.getMembership(2)).toBeUndefined()
    expect(store.getMembership(3)).toBeUndefined()
    expect(store.getMembership(100)).toBeDefined()
    expect(store.getStats()).toEqual({ familyCount: 2, runCount: 3, memberCount: 5 })
  })

  it('tracks sheet member counts and consumes family-level source transforms', () => {
    const store = createFormulaFamilyStore()
    store.upsertFormula({ cellIndex: 1, sheetId: 1, row: 0, col: 2, templateId: 1, shapeKey: 'a' })
    store.upsertFormula({ cellIndex: 2, sheetId: 1, row: 1, col: 2, templateId: 1, shapeKey: 'a' })

    const family = store.listFamilies()[0]
    expect(family).toBeDefined()
    const transform = {
      ownerSheetName: 'Sheet1',
      targetSheetName: 'Sheet1',
      transform: { kind: 'insert', axis: 'column', start: 1, count: 1 } as const,
      preservesValue: true,
    }
    store.setStructuralSourceTransform(family.id, transform)

    expect(store.countSheetMembers(1)).toBe(2)
    expect(store.getStructuralSourceTransform(1)).toBe(transform)
    expect(store.consumeStructuralSourceTransforms()).toEqual([{ cellIndices: [1, 2], transform }])
    expect(store.getStructuralSourceTransform(1)).toBeUndefined()
  })
})
