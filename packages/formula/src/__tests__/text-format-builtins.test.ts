import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getTextBuiltin } from '../builtins/text.js'

describe('text format builtins', () => {
  it('supports TEXT for numeric, date-time, and text-section formatting', () => {
    const TEXT = getTextBuiltin('TEXT')!

    expect(TEXT(number(1234.567), text('#,##0.00'))).toEqual(text('1,234.57'))
    expect(TEXT(number(0.1234), text('0.0%'))).toEqual(text('12.3%'))
    expect(TEXT(number(45356), text('yyyy-mm-dd'))).toEqual(text('2024-03-05'))
    expect(TEXT(number(0.5), text('h:mm AM/PM'))).toEqual(text('12:00 PM'))
    expect(TEXT(text('alpha'), text('prefix @'))).toEqual(text('prefix alpha'))
  })

  it('supports VALUE, NUMBERVALUE, and VALUETOTEXT conversions', () => {
    const VALUE = getTextBuiltin('VALUE')!
    const NUMBERVALUE = getTextBuiltin('NUMBERVALUE')!
    const VALUETOTEXT = getTextBuiltin('VALUETOTEXT')!

    expect(VALUE(text(' 42 '))).toEqual(number(42))
    expect(NUMBERVALUE(text('2.500,27'), text(','), text('.'))).toEqual(number(2500.27))
    expect(NUMBERVALUE(text('9%%'))).toEqual(number(0.0009))
    expect(VALUETOTEXT(number(42))).toEqual(text('42'))
    expect(VALUETOTEXT(text('alpha'), number(1))).toEqual(text('"alpha"'))
    expect(VALUETOTEXT(err(ErrorCode.Ref))).toEqual(text('#REF!'))
  })

  it('keeps validation and error propagation for text formatting builtins', () => {
    const TEXT = getTextBuiltin('TEXT')!
    const VALUE = getTextBuiltin('VALUE')!
    const NUMBERVALUE = getTextBuiltin('NUMBERVALUE')!
    const VALUETOTEXT = getTextBuiltin('VALUETOTEXT')!

    expect(TEXT()).toEqual(valueError())
    expect(TEXT(text('alpha'), text('0.00'))).toEqual(valueError())
    expect(TEXT(err(ErrorCode.Ref), text('0.00'))).toEqual(err(ErrorCode.Ref))
    expect(VALUE()).toEqual(valueError())
    expect(VALUE(err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(NUMBERVALUE(text('1.2.3'), text('.'), text(','))).toEqual(valueError())
    expect(VALUETOTEXT(text('alpha'), number(2))).toEqual(valueError())
  })

  it('covers text format sections, decorations, and date-time tokens', () => {
    const TEXT = getTextBuiltin('TEXT')!

    expect(TEXT(number(-12), text('"pos";"neg";"zero"'))).toEqual(text('neg'))
    expect(TEXT(number(-12), text('"abc"'))).toEqual(text('-abc'))
    expect(TEXT(number(0), text('0.0;[Red]-0.0;"zero"'))).toEqual(text('zero'))
    expect(TEXT(number(7), text('_)*x[Blue]000"kg"'))).toEqual(text(' 007kg'))
    expect(TEXT(text('alpha'), text('"literal"'))).toEqual(text('literal'))
    expect(TEXT(text('alpha'), text('0.0;[Red]-0.0;"zero";prefix @ suffix'))).toEqual(text('prefix alpha suffix'))

    expect(TEXT(number(45356.52425925926), text('dddd, mmmm d, yyyy h:m:s a/p'))).toEqual(text('Tuesday, March 5, 2024 12:34:56 p'))
    expect(TEXT(number(45356), text('ddd mmm dd yy'))).toEqual(text('Tue Mar 05 24'))
    expect(TEXT(number(Number.POSITIVE_INFINITY), text('yyyy-mm-dd'))).toEqual(valueError())
  })

  it('covers VALUETOTEXT labels and NUMBERVALUE parser validation branches', () => {
    const NUMBERVALUE = getTextBuiltin('NUMBERVALUE')!
    const VALUETOTEXT = getTextBuiltin('VALUETOTEXT')!

    expect(VALUETOTEXT({ tag: ValueTag.Empty })).toEqual(text(''))
    expect(VALUETOTEXT({ tag: ValueTag.Boolean, value: false })).toEqual(text('FALSE'))
    expect(VALUETOTEXT(err(ErrorCode.Div0))).toEqual(text('#DIV/0!'))
    expect(VALUETOTEXT(err(ErrorCode.Value))).toEqual(text('#VALUE!'))
    expect(VALUETOTEXT(err(ErrorCode.Name))).toEqual(text('#NAME?'))
    expect(VALUETOTEXT(err(ErrorCode.NA))).toEqual(text('#N/A'))
    expect(VALUETOTEXT(err(ErrorCode.Cycle))).toEqual(text('#CYCLE!'))
    expect(VALUETOTEXT(err(ErrorCode.Spill))).toEqual(text('#SPILL!'))
    expect(VALUETOTEXT(err(ErrorCode.Blocked))).toEqual(text('#BLOCKED!'))

    expect(NUMBERVALUE(text('1%2'))).toEqual(valueError())
    expect(NUMBERVALUE(text('1.2'), text('.'), text('.'))).toEqual(valueError())
    expect(NUMBERVALUE(text('1.2,3'), text('.'), text(','))).toEqual(valueError())
    expect(NUMBERVALUE(text('+'))).toEqual(valueError())
    expect(NUMBERVALUE(text('bad'))).toEqual(valueError())
    expect(NUMBERVALUE(text('12,5'), text(','), text(''))).toEqual(number(12.5))
  })
})

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function valueError(): CellValue {
  return err(ErrorCode.Value)
}
