import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const valueError = err(ErrorCode.Value)

describe('radix builtins', () => {
  it('covers base, decimal, signed radix, and roman validation branches', () => {
    const BASE = getBuiltin('BASE')!
    const DECIMAL = getBuiltin('DECIMAL')!
    const BIN2DEC = getBuiltin('BIN2DEC')!
    const BIN2HEX = getBuiltin('BIN2HEX')!
    const DEC2BIN = getBuiltin('DEC2BIN')!
    const DEC2HEX = getBuiltin('DEC2HEX')!
    const HEX2DEC = getBuiltin('HEX2DEC')!
    const OCT2BIN = getBuiltin('OCT2BIN')!
    const ROMAN = getBuiltin('ROMAN')!
    const ARABIC = getBuiltin('ARABIC')!

    expect(BASE(num(31), num(16), num(4))).toEqual(text('001F'))
    expect(BASE(num(-1), num(16))).toEqual(valueError)
    expect(BASE(num(31), num(1))).toEqual(valueError)
    expect(BASE(num(31), num(37))).toEqual(valueError)
    expect(BASE(num(31), num(16), text('bad'))).toEqual(valueError)

    expect(DECIMAL(err(ErrorCode.Ref), num(2))).toEqual(err(ErrorCode.Ref))
    expect(DECIMAL(text(' 1f '), num(16))).toEqual(num(31))
    expect(DECIMAL(num(101), num(2))).toEqual(num(5))
    expect(DECIMAL(text(''), num(10))).toEqual(valueError)
    expect(DECIMAL(text('2'), num(2))).toEqual(valueError)
    expect(DECIMAL(text('10'), num(37))).toEqual(valueError)

    expect(BIN2DEC(text('1111111111'))).toEqual(num(-1))
    expect(BIN2DEC(err(ErrorCode.NA))).toEqual(valueError)
    expect(BIN2DEC(text(''))).toEqual(valueError)
    expect(BIN2DEC(text('102'))).toEqual(valueError)
    expect(BIN2DEC(text('11111111111'))).toEqual(valueError)
    expect(BIN2HEX(text('1111111111'))).toEqual(text('FFFFFFFFFF'))
    expect(BIN2HEX(text('1010'), text('bad'))).toEqual(valueError)
    expect(BIN2HEX(text('1010'), num(0))).toEqual(valueError)
    expect(DEC2BIN(num(511), num(10))).toEqual(text('0111111111'))
    expect(DEC2BIN(num(512))).toEqual(valueError)
    expect(DEC2HEX(num(-1))).toEqual(text('FFFFFFFFFF'))
    expect(HEX2DEC(text('FFFFFFFFFF'))).toEqual(num(-1))
    expect(OCT2BIN(text('7777777777'))).toEqual(text('1111111111'))

    expect(ROMAN(num(3999))).toEqual(text('MMMCMXCIX'))
    expect(ROMAN(num(0))).toEqual(valueError)
    expect(ARABIC(text('XLIV'))).toEqual(num(44))
    expect(ARABIC(text(''))).toEqual(valueError)
    expect(ARABIC(text('IIV'))).toEqual(valueError)
    expect(ARABIC(num(44))).toEqual(valueError)
  })
})
