import { describe, expect, it } from 'vitest'
import { mapStructuralAxisIndex, mapStructuralBoundary, structuralTransformForOp } from '../engine-structural-utils.js'

describe('engine structural utils', () => {
  it('builds structural transforms from workbook ops', () => {
    expect(structuralTransformForOp({ kind: 'insertRows', sheetName: 'Sheet1', start: 2, count: 3 })).toEqual({
      kind: 'insert',
      axis: 'row',
      start: 2,
      count: 3,
    })
    expect(
      structuralTransformForOp({
        kind: 'moveColumns',
        sheetName: 'Sheet1',
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toEqual({
      kind: 'move',
      axis: 'column',
      start: 4,
      count: 2,
      target: 1,
    })
  })

  it('maps deleted spans to undefined and shifts later indices', () => {
    const transform = structuralTransformForOp({
      kind: 'deleteRows',
      sheetName: 'Sheet1',
      start: 3,
      count: 2,
    })

    expect(mapStructuralAxisIndex(2, transform)).toBe(2)
    expect(mapStructuralAxisIndex(3, transform)).toBeUndefined()
    expect(mapStructuralAxisIndex(4, transform)).toBeUndefined()
    expect(mapStructuralAxisIndex(6, transform)).toBe(4)
    expect(mapStructuralBoundary(4, transform)).toBe(0)
  })

  it('maps moved spans when the target is before the source', () => {
    const transform = structuralTransformForOp({
      kind: 'moveRows',
      sheetName: 'Sheet1',
      start: 5,
      count: 2,
      target: 1,
    })

    expect(mapStructuralAxisIndex(1, transform)).toBe(3)
    expect(mapStructuralAxisIndex(4, transform)).toBe(6)
    expect(mapStructuralAxisIndex(5, transform)).toBe(1)
    expect(mapStructuralAxisIndex(6, transform)).toBe(2)
  })

  it('maps moved spans when the target is after the source', () => {
    const transform = structuralTransformForOp({
      kind: 'moveColumns',
      sheetName: 'Sheet1',
      start: 2,
      count: 2,
      target: 5,
    })

    expect(mapStructuralAxisIndex(2, transform)).toBe(5)
    expect(mapStructuralAxisIndex(3, transform)).toBe(6)
    expect(mapStructuralAxisIndex(4, transform)).toBe(2)
    expect(mapStructuralAxisIndex(6, transform)).toBe(4)
    expect(mapStructuralBoundary(3, transform)).toBe(6)
  })

  it('throws for unsupported structural op and transform variants', () => {
    const invalidOp = { kind: 'insertRows' as const, sheetName: 'Sheet1', start: 0, count: 1 }
    Reflect.set(invalidOp, 'kind', 'noop')

    const invalidTransform = { kind: 'insert' as const, axis: 'row' as const, start: 0, count: 1 }
    Reflect.set(invalidTransform, 'kind', 'noop')

    expect(() => structuralTransformForOp(invalidOp)).toThrow(
      'Unhandled structural transform case: {"kind":"noop","sheetName":"Sheet1","start":0,"count":1}',
    )
    expect(() => mapStructuralAxisIndex(0, invalidTransform)).toThrow(
      'Unhandled structural transform case: {"kind":"noop","axis":"row","start":0,"count":1}',
    )
  })
})
