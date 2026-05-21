import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@bilig/protocol'
import { DirectFormulaIndexCollection, PendingNumericCellValues } from '../engine/services/direct-formula-index-collection.js'

describe('direct formula index collection', () => {
  it('tracks pending numeric cells with sparse assigned slots', () => {
    const values = new PendingNumericCellValues()

    expect(values.has(4)).toBe(false)
    expect(values.get(4)).toBeUndefined()

    values.set(4, -0)

    expect(values.has(4)).toBe(true)
    expect(Object.is(values.get(4), -0)).toBe(true)
    expect(values.has(5)).toBe(false)
  })

  it('maintains direct formula indices, deltas, validation, and covered inputs through sparse and bulk paths', () => {
    const indices = new DirectFormulaIndexCollection()

    expect(indices.size).toBe(0)
    expect(indices.has(8)).toBe(false)
    expect(indices.hasCompleteDeltas()).toBe(false)
    expect(indices.hasCoveredDirectRangeInput(3)).toBe(false)
    expect(indices.hasCoveredDirectFormulaInput(3)).toBe(false)

    indices.add(8)
    indices.add(8)
    indices.addDelta(8, 2)
    indices.addScalarDelta(8, 3)
    indices.addCurrentResult(8, { kind: 'number', value: 5 })
    indices.markDirectRangeInputCovered(3)
    indices.markDirectRangeInputCovered(3)
    indices.markDirectFormulaInputCovered(4)
    indices.markDirectFormulaInputCovered(4)

    expect(indices.size).toBe(1)
    expect(indices.has(8)).toBe(true)
    expect(indices.getDelta(8)).toBe(5)
    expect(indices.getDelta(9)).toBeUndefined()
    expect(indices.getDeltaAt(0)).toBe(5)
    expect(indices.getCurrentResult(8)).toEqual({ kind: 'number', value: 5 })
    expect(indices.getCurrentResultAt(0)).toEqual({ kind: 'number', value: 5 })
    expect(indices.hasCoveredDirectRangeInput(3)).toBe(true)
    expect(indices.hasCoveredDirectRangeInput(4)).toBe(false)
    expect(indices.hasCoveredDirectFormulaInput(4)).toBe(true)
    expect(indices.hasCoveredDirectFormulaInput(5)).toBe(false)
    expect(indices.getConstantScalarDelta()).toBeUndefined()
    expect(indices.hasCompleteScalarDeltas()).toBe(false)

    const visited: number[] = []
    indices.forEach((cellIndex) => visited.push(cellIndex))
    expect(visited).toEqual([8])

    const indexed: [number, number][] = []
    indices.forEachIndexed((cellIndex, index) => indexed.push([cellIndex, index]))
    expect(indexed).toEqual([[8, 0]])

    indices.appendDeltas([10, 11], [1, 2], 'scalar')
    expect(indices.getDelta(10)).toBe(1)
    expect(indices.getScalarDeltaAt(1)).toBe(1)
    expect(indices.getScalarDeltaAt(2)).toBe(2)

    indices.appendDeltas(new Uint32Array([10, 12]), [4, 6])
    expect(indices.getDelta(10)).toBe(5)
    expect(indices.getDelta(12)).toBe(6)
    expect(indices.getScalarDeltaAt(1)).toBeUndefined()

    indices.addCurrentResult(12, { kind: 'error', code: ErrorCode.NA })
    expect(indices.getCurrentResult(12)).toEqual({ kind: 'error', code: ErrorCode.NA })
  })

  it('keeps constant delta and bulk append metadata coherent', () => {
    const constant = new DirectFormulaIndexCollection()
    constant.appendConstantDelta(new Uint32Array([20, 21, 22]), 7, 'scalar')

    expect(constant.size).toBe(3)
    expect(constant.getDelta(20)).toBe(7)
    expect(constant.getDeltaAt(2)).toBe(7)
    expect(constant.getScalarDeltaAt(1)).toBe(7)
    expect(constant.getConstantScalarDelta()).toBe(7)
    expect(constant.getConstantDelta()).toBe(7)
    expect(constant.hasCompleteScalarDeltas()).toBe(true)
    expect(constant.hasValidatedScalarDeltaCells()).toBe(false)

    constant.markScalarDeltaCellsValidated()
    expect(constant.hasValidatedScalarDeltaCells()).toBe(true)

    constant.addScalarDelta(23, 7)
    expect(constant.getDelta(23)).toBe(7)
    expect(constant.getConstantScalarDelta()).toBe(7)

    constant.addDelta(24, 7)
    expect(constant.getConstantDelta()).toBe(7)
    expect(constant.getConstantScalarDelta()).toBeUndefined()

    const mixedConstant = new DirectFormulaIndexCollection()
    mixedConstant.appendConstantDelta(
      Array.from({ length: 24 }, (_unused, index) => index + 200),
      3,
      'scalar',
    )
    mixedConstant.appendConstantDelta(
      Array.from({ length: 24 }, (_unused, index) => index + 300),
      3,
    )

    expect(mixedConstant.size).toBe(48)
    expect(mixedConstant.getConstantDelta()).toBe(3)
    expect(mixedConstant.getConstantScalarDelta()).toBeUndefined()
    expect(mixedConstant.getDelta(223)).toBe(3)
    expect(mixedConstant.getDelta(323)).toBe(3)

    mixedConstant.appendConstantDelta(new Uint32Array([323, 400]), 3)
    expect(mixedConstant.getConstantDelta()).toBeUndefined()
    expect(mixedConstant.getDelta(323)).toBe(6)
    expect(mixedConstant.getDelta(400)).toBe(3)

    const bulk = new DirectFormulaIndexCollection()
    bulk.appendConstantDelta(
      Array.from({ length: 18 }, (_unused, index) => index + 100),
      1,
      'scalar',
    )
    bulk.appendDeltas([110, 118, 119], [2, 3, 4], 'scalar')
    bulk.appendConstantDelta(new Uint32Array([119, 120]), 5)

    expect(bulk.has(110)).toBe(true)
    expect(bulk.getCellIndexAt(0)).toBe(100)
    expect(bulk.getDelta(110)).toBe(3)
    expect(bulk.getDelta(118)).toBe(3)
    expect(bulk.getDelta(119)).toBe(9)
    expect(bulk.getDelta(120)).toBe(5)
    expect(bulk.getScalarDeltaAt(19)).toBeUndefined()
  })
})
