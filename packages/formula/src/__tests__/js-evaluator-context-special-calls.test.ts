import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { evaluatePlan, lowerToPlan, parseFormula } from '../index.js'

const context = {
  sheetName: 'Sheet2',
  currentAddress: 'C4',
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case 'A1':
        return number(2)
      case 'B1':
        return number(3)
      default:
        return empty()
    }
  },
  resolveRange: (_sheetName: string, _start: string, _end: string): CellValue[] => [],
  listSheetNames: () => ['Sheet1', 'Sheet2', 'Summary'],
  resolveFormula: (sheetName: string, address: string): string | undefined =>
    sheetName === 'Sheet2' && address === 'B1' ? 'SUM(A1:A2)' : sheetName === 'Sheet2' && address === 'C1' ? 'A1*2' : undefined,
}

describe('js evaluator context special calls', () => {
  it('evaluates metadata and context helpers', () => {
    expect(evaluatePlan(lowerToPlan(parseFormula('ROW()')), context)).toEqual(number(4))
    expect(evaluatePlan(lowerToPlan(parseFormula('COLUMN(B:D)')), context)).toEqual(number(2))
    expect(evaluatePlan(lowerToPlan(parseFormula('FORMULATEXT(Sheet2!B1)')), context)).toEqual(text('=SUM(A1:A2)'))
    expect(evaluatePlan(lowerToPlan(parseFormula('FORMULA(Sheet2!C1)')), context)).toEqual(text('=A1*2'))
    expect(evaluatePlan(lowerToPlan(parseFormula('PHONETIC(42)')), context)).toEqual(text('42'))
    expect(evaluatePlan(lowerToPlan(parseFormula('CHOOSE(2,"a","b")')), context)).toEqual(text('b'))
    expect(evaluatePlan(lowerToPlan(parseFormula('LAMBDA(x,IF(ISOMITTED(x),1,0))()')), context)).toEqual(number(1))
  })

  it('evaluates sheet and cell info helpers', () => {
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET()')), context)).toEqual(number(2))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Summary")')), context)).toEqual(number(3))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEETS()')), context)).toEqual(number(3))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("address",B1)')), context)).toEqual(text('$B$1'))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("row",A1)')), context)).toEqual(number(1))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("col",B1)')), context)).toEqual(number(2))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("type",B1)')), context)).toEqual(text('v'))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("filename")')), context)).toEqual(text(''))
  })

  it('preserves validation and NA or REF branches', () => {
    expect(evaluatePlan(lowerToPlan(parseFormula('CHOOSE(1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('ROW(A1,B1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('COLUMN(A1,B1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('ISOMITTED()')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('FORMULATEXT(1)')), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluatePlan(lowerToPlan(parseFormula('FORMULATEXT(A1,B1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET("Missing")')), context)).toEqual(err(ErrorCode.NA))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET(1)')), context)).toEqual(err(ErrorCode.NA))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEET(A1,B1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEETS(1)')), context)).toEqual(err(ErrorCode.NA))
    expect(evaluatePlan(lowerToPlan(parseFormula('SHEETS(A1,B1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL("bogus",A1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL(1,A1)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL()')), context)).toEqual(err(ErrorCode.Value))
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("contents")')), {
        ...context,
        currentAddress: undefined,
      }),
    ).toEqual(err(ErrorCode.Value))
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("address")')), {
        ...context,
        currentAddress: undefined,
      }),
    ).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('PHONETIC()')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('PHONETIC(LAMBDA(x,x))')), context)).toEqual(err(ErrorCode.Value))
    expect(
      evaluatePlan(lowerToPlan(parseFormula('CELL("type")')), {
        ...context,
        currentAddress: undefined,
      }),
    ).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CHOOSE(0,1,2)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CHOOSE("x",1,2)')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('PHONETIC()')), context)).toEqual(err(ErrorCode.Value))
    expect(evaluatePlan(lowerToPlan(parseFormula('CELL(TRUE(),A1)')), context)).toEqual(err(ErrorCode.Value))
  })

  it('uses direct exact lookup without materializing the range when a handler is present', () => {
    const resolveExactVectorMatch = vi.fn(() => ({ handled: true, position: 2 }))
    const noteExactLookupDirect = vi.fn()
    const resolveRange = vi.fn(() => {
      throw new Error('resolveRange should not be called for direct exact lookup')
    })

    expect(
      evaluatePlan(lowerToPlan(parseFormula('XMATCH("pear",A1:A3,0,-1)')), {
        ...context,
        resolveRange,
        resolveExactVectorMatch,
        noteExactLookupDirect,
      }),
    ).toEqual(number(2))
    expect(resolveExactVectorMatch).toHaveBeenCalledWith({
      lookupValue: text('pear'),
      sheetName: 'Sheet2',
      start: 'A1',
      end: 'A3',
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 0,
      searchMode: -1,
    })
    expect(noteExactLookupDirect).toHaveBeenCalledTimes(1)
    expect(resolveRange).not.toHaveBeenCalled()
  })

  it('falls back to normal range evaluation when direct exact lookup is unavailable', () => {
    const resolveExactVectorMatch = vi.fn(() => ({ handled: false as const }))
    const noteExactLookupFallback = vi.fn()
    const noteRangeMaterialization = vi.fn()
    const resolveRange = vi.fn(() => [number(1), number(2), number(3)])

    expect(
      evaluatePlan(lowerToPlan(parseFormula('MATCH(2,A1:A3,0)')), {
        ...context,
        resolveRange,
        resolveExactVectorMatch,
        noteExactLookupFallback,
        noteRangeMaterialization,
      }),
    ).toEqual(number(2))
    expect(resolveExactVectorMatch).toHaveBeenCalledTimes(1)
    expect(noteExactLookupFallback).toHaveBeenCalledTimes(1)
    expect(noteRangeMaterialization).toHaveBeenCalledWith(3)
    expect(resolveRange).toHaveBeenCalledTimes(1)
  })

  it('evaluates approximate lookup opcodes through the normal builtin path', () => {
    const noteRangeMaterialization = vi.fn()
    const resolveRange = vi.fn(() => [number(1), number(3), number(5)])

    expect(
      evaluatePlan(
        [
          { opcode: 'push-number', value: 4 },
          {
            opcode: 'lookup-approximate-match',
            callee: 'MATCH',
            start: 'A1',
            end: 'A3',
            startRow: 0,
            endRow: 2,
            startCol: 0,
            endCol: 0,
            refKind: 'cells',
            matchMode: 1,
          },
          { opcode: 'return' },
        ],
        {
          ...context,
          resolveRange,
          noteRangeMaterialization,
        },
      ),
    ).toEqual(number(2))
    expect(noteRangeMaterialization).toHaveBeenCalledWith(3)
    expect(resolveRange).toHaveBeenCalledTimes(1)
  })

  it('uses direct approximate lookup without materializing the range when a handler is present', () => {
    const resolveApproximateVectorMatch = vi.fn(() => ({ handled: true, position: 2 }))
    const resolveRange = vi.fn(() => {
      throw new Error('resolveRange should not be called for direct approximate lookup')
    })

    expect(
      evaluatePlan(lowerToPlan(parseFormula('MATCH(4,A1:A3,1)')), {
        ...context,
        resolveRange,
        resolveApproximateVectorMatch,
      }),
    ).toEqual(number(2))
    expect(resolveApproximateVectorMatch).toHaveBeenCalledWith({
      lookupValue: number(4),
      sheetName: 'Sheet2',
      start: 'A1',
      end: 'A3',
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 0,
      matchMode: 1,
    })
    expect(resolveRange).not.toHaveBeenCalled()
  })

  it('falls back to XMATCH approximate evaluation when the direct handler declines', () => {
    const resolveApproximateVectorMatch = vi.fn(() => ({ handled: false as const }))
    const noteRangeMaterialization = vi.fn()
    const resolveRange = vi.fn(() => [number(9), number(7), number(5)])

    expect(
      evaluatePlan(lowerToPlan(parseFormula('XMATCH(6,A1:A3,-1)')), {
        ...context,
        resolveRange,
        resolveApproximateVectorMatch,
        noteRangeMaterialization,
      }),
    ).toEqual(number(2))
    expect(resolveApproximateVectorMatch).toHaveBeenCalledTimes(1)
    expect(noteRangeMaterialization).toHaveBeenCalledWith(3)
    expect(resolveRange).toHaveBeenCalledTimes(1)
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
