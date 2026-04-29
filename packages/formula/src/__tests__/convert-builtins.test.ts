import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const empty = (): CellValue => ({ tag: ValueTag.Empty })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toMatchObject({ tag: ValueTag.Number })
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number result, received ${value.tag}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

describe('convert builtins', () => {
  it('covers scalar coercion, prefixed units, and remaining temperature branches', () => {
    const CONVERT = getBuiltin('CONVERT')!
    const EUROCONVERT = getBuiltin('EUROCONVERT')!

    expectNumber(CONVERT(bool(true), text('m'), text('cm')), 100)
    expectNumber(CONVERT(empty(), text('m'), text('cm')), 0)
    expect(CONVERT(text('bad'), text('m'), text('cm'))).toEqual(err(ErrorCode.Value))
    expect(CONVERT(num(1), num(1), text('cm'))).toEqual(err(ErrorCode.Value))

    expectNumber(CONVERT(num(1), text('km'), text('m')), 1000)
    expectNumber(CONVERT(num(1), text('Mbyte'), text('byte')), 1_000_000)
    expectNumber(CONVERT(num(1), text('Mibyte'), text('byte')), 1_048_576)
    expectNumber(CONVERT(num(1), text('cm2'), text('m2')), 0.0001)
    expectNumber(CONVERT(num(1), text('cm3'), text('m3')), 0.000001)

    expectNumber(CONVERT(num(491.67), text('Rank'), text('F')), 32)
    expectNumber(CONVERT(num(80), text('Reau'), text('C')), 100)
    expectNumber(CONVERT(num(273.15), text('K'), text('Rank')), 491.67)

    expect(EUROCONVERT(num(1), text('DEM'), text('EUR'), text('true'))).toEqual(err(ErrorCode.Value))
    expect(EUROCONVERT(num(1), text('DEM'), text('EUR'), bool(false), text('bad'))).toEqual(err(ErrorCode.Value))
    expectNumber(EUROCONVERT(num(1), text('EUR'), text('ITL')), 1936)
  })
})
