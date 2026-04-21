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
})
