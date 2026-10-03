import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { createWorkPaperInternals, doesWorkPaperColumnSearchValueMatch, type WorkPaperInternalsHooks } from '../work-paper-internals.js'
import type { WorkPaperCellAddress } from '../work-paper-types.js'

function testHooks(values: readonly CellValue[] = []): WorkPaperInternalsHooks {
  return {
    getCellDependents: vi.fn(() => [{ kind: 'name', name: 'DependentName' }]),
    getCellFormula: vi.fn(() => '=A1'),
    getCellPrecedents: vi.fn(() => [{ kind: 'name', name: 'PrecedentName' }]),
    getCellValue: vi.fn((address: WorkPaperCellAddress) => values[address.row] ?? { tag: ValueTag.Empty }),
    getNamedExpressionsFromFormula: vi.fn(() => ['NamedInput']),
    getRangeSerialized: vi.fn(() => [[1]]),
    getRangeValues: vi.fn(() => [[{ tag: ValueTag.Number, value: 1 }]]),
    getSheetDimensions: vi.fn(() => ({ width: 2, height: values.length })),
    getSheetId: vi.fn((name: string) => (name === 'Sheet1' ? 1 : undefined)),
    getSheetName: vi.fn((sheetId: number) => (sheetId === 1 ? 'Sheet1' : undefined)),
    getSheetNames: vi.fn(() => ['Sheet1']),
    hasCellValueOrFormula: vi.fn(() => true),
    isCellPartOfArray: vi.fn(() => false),
    normalizeFormula: vi.fn((formula: string) => formula.toUpperCase()),
    recalculate: vi.fn(() => []),
    calculateFormula: vi.fn(() => ({ tag: ValueTag.Number, value: 3 })),
    countSheets: vi.fn(() => 1),
    validateFormula: vi.fn(() => true),
  }
}

describe('work paper internals', () => {
  it('creates frozen adapters that delegate to runtime hooks', () => {
    const hooks = testHooks()
    const internals = createWorkPaperInternals(hooks)
    const address = { sheet: 1, row: 0, col: 0 }

    expect(Object.isFrozen(internals)).toBe(true)
    expect(Object.isFrozen(internals.graph)).toBe(true)
    expect(Object.isFrozen(internals.rangeMapping)).toBe(true)
    expect(Object.isFrozen(internals.arrayMapping)).toBe(true)
    expect(Object.isFrozen(internals.sheetMapping)).toBe(true)
    expect(Object.isFrozen(internals.addressMapping)).toBe(true)
    expect(Object.isFrozen(internals.dependencyGraph)).toBe(true)
    expect(Object.isFrozen(internals.evaluator)).toBe(true)
    expect(Object.isFrozen(internals.columnSearch)).toBe(true)
    expect(Object.isFrozen(internals.lazilyTransformingAstService)).toBe(true)

    expect(internals.graph.getDependents(address)).toEqual([{ kind: 'name', name: 'DependentName' }])
    expect(internals.dependencyGraph.getCellPrecedents(address)).toEqual([{ kind: 'name', name: 'PrecedentName' }])
    expect(internals.rangeMapping.getSerialized({ start: address, end: address })).toEqual([[1]])
    expect(internals.arrayMapping.getFormula(address)).toBe('=A1')
    expect(internals.sheetMapping.getSheetId('Sheet1')).toBe(1)
    expect(internals.addressMapping.has(address)).toBe(true)
    expect(internals.evaluator.calculateFormula('=1+2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(internals.lazilyTransformingAstService.normalizeFormula('=a1')).toBe('=A1')
  })

  it('searches a column with string and predicate matchers', () => {
    const hooks = testHooks([
      { tag: ValueTag.String, value: 'match' },
      { tag: ValueTag.String, value: 'skip' },
      { tag: ValueTag.Number, value: 42 },
    ])
    const internals = createWorkPaperInternals(hooks)

    expect(internals.columnSearch.find(1, 3, 'match')).toEqual([{ sheet: 1, row: 0, col: 3 }])
    expect(internals.columnSearch.find(1, 3, (value) => value.tag === ValueTag.Number)).toEqual([{ sheet: 1, row: 2, col: 3 }])
  })

  it('only treats string cells as direct string-search matches', () => {
    expect(doesWorkPaperColumnSearchValueMatch({ tag: ValueTag.String, value: 'needle' }, 'needle')).toBe(true)
    expect(doesWorkPaperColumnSearchValueMatch({ tag: ValueTag.Number, value: 7 }, '7')).toBe(false)
    expect(doesWorkPaperColumnSearchValueMatch({ tag: ValueTag.Boolean, value: true }, (value) => value.tag === ValueTag.Boolean)).toBe(
      true,
    )
  })
})
