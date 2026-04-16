import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { evaluatePlan, evaluatePlanResult, lowerToPlan, parseFormula } from '../index.js'

const context = {
  sheetName: 'Sheet1',
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case 'A1':
        return number(2)
      case 'B1':
        return number(3)
      case 'A2':
        return { tag: ValueTag.Boolean, value: true }
      default:
        return empty()
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
    if (start === 'A1' && end === 'B2') {
      return [number(2), number(3), { tag: ValueTag.Boolean, value: true }, empty()]
    }
    if (start === 'A1' && end === 'D4') {
      return [
        empty(),
        empty(),
        empty(),
        empty(),
        empty(),
        number(1),
        number(2),
        empty(),
        empty(),
        number(3),
        empty(),
        empty(),
        empty(),
        empty(),
        empty(),
        empty(),
      ]
    }
    return []
  },
}

describe('js evaluator array specials', () => {
  it('evaluates spill-oriented array helpers', () => {
    expect(evaluatePlanResult(lowerToPlan(parseFormula('TEXTSPLIT("Ab|aB","|","",FALSE(),1)')), context)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [text('Ab'), text('aB')],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('EXPAND(A1:B2,3,3,0)')), context)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [
        number(2),
        number(3),
        number(0),
        { tag: ValueTag.Boolean, value: true },
        empty(),
        number(0),
        number(0),
        number(0),
        number(0),
      ],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('TEXTSPLIT("red,blue|green",",","|")')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('red'), text('blue'), text('green'), err(ErrorCode.NA)],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('TRIMRANGE(A1:D4)')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [number(1), number(2), number(3), empty()],
    })
  })

  it('evaluates lambda-based array helpers', () => {
    expect(evaluatePlanResult(lowerToPlan(parseFormula('MAKEARRAY(2,2,LAMBDA(r,c,r+c))')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [number(2), number(3), number(3), number(4)],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('MAP(A1:B2,LAMBDA(x,x+1))')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [number(3), number(4), number(2), number(1)],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('BYROW(A1:B2,LAMBDA(r,SUM(r)))')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [number(5), number(1)],
    })

    expect(evaluatePlanResult(lowerToPlan(parseFormula('SCAN(0,A1:B2,LAMBDA(a,x,a+x))')), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [number(2), number(5), number(6), number(6)],
    })
  })

  it('preserves validation errors for array special calls', () => {
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("alpha","")')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("a,b",",","",TRUE(),"x")')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("alpha",",","",SEQUENCE(2))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('TEXTSPLIT("alpha",",","",TRUE(),SEQUENCE(2))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('EXPAND(A1:B2,1,1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('EXPAND(A1:B2,"x",3)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('EXPAND(A1:B2,SEQUENCE(2),3)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('TRIMRANGE(A1:B2,SEQUENCE(2))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('MAKEARRAY("x",1,LAMBDA(r,c,r+c))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('MAKEARRAY(2,SEQUENCE(2),LAMBDA(r,c,r+c))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('MAP(LAMBDA(x,x))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('MAKEARRAY(2,2,LAMBDA(r,c,SEQUENCE(2)))')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('BYROW(A1:B2,LAMBDA(r,SEQUENCE(2)))')), context)).toEqual(err(ErrorCode.Value))
  })
})

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}
