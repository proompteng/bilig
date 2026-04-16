import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { evaluateAst, evaluateAstResult, parseFormula } from '../index.js'

describe('js evaluator workbook special calls', () => {
  it('evaluates GETPIVOTDATA and INDIRECT through workbook hooks', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => (address === 'A1' ? number(2) : { tag: ValueTag.Empty }),
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] =>
        start === 'A1' && end === 'A3' ? [number(2), number(3), number(4)] : [],
      resolveName: (name: string): CellValue => (name === 'TaxRate' ? number(0.085) : err(ErrorCode.Name)),
      resolvePivotData: ({
        dataField,
        address,
        filters,
      }: {
        dataField: string
        sheetName: string
        address: string
        filters: ReadonlyArray<{ field: string; item: CellValue }>
      }) =>
        dataField === 'Sales' &&
        address === 'B2' &&
        filters.length === 1 &&
        filters[0]?.field === 'Region' &&
        filters[0].item.tag === ValueTag.String &&
        filters[0].item.value === 'East'
          ? number(15)
          : err(ErrorCode.Ref),
    }

    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,"Region","East")'), context)).toEqual(number(15))
    expect(evaluateAst(parseFormula('INDIRECT("A1")'), context)).toEqual(number(2))
    expect(evaluateAst(parseFormula('INDIRECT("TaxRate")+1'), context)).toEqual(number(1.085))
    expect(evaluateAst(parseFormula('INDIRECT("")'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAstResult(parseFormula('INDIRECT("A1:A3")'), context)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [number(2), number(3), number(4)],
    })
  })

  it('evaluates workbook-shaped group and multiple-operation helpers', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => (address === 'B5' ? number(0) : { tag: ValueTag.Empty }),
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A5') {
          return [text('Region'), text('East'), text('West'), text('East'), text('West')]
        }
        if (start === 'B1' && end === 'B5') {
          return [text('Product'), text('Widget'), text('Widget'), text('Gizmo'), text('Gizmo')]
        }
        if (start === 'C1' && end === 'C5') {
          return [text('Sales'), number(10), number(7), number(5), number(4)]
        }
        return []
      },
      resolveMultipleOperations: ({
        formulaSheetName,
        formulaAddress,
        rowCellAddress,
        rowReplacementAddress,
        columnCellAddress,
        columnReplacementAddress,
      }: {
        formulaSheetName: string
        formulaAddress: string
        rowCellSheetName: string
        rowCellAddress: string
        rowReplacementSheetName: string
        rowReplacementAddress: string
        columnCellSheetName?: string
        columnCellAddress?: string
        columnReplacementSheetName?: string
        columnReplacementAddress?: string
      }) =>
        formulaSheetName === 'Sheet1' &&
        formulaAddress === 'B5' &&
        rowCellAddress === 'B3' &&
        rowReplacementAddress === 'C4' &&
        columnCellAddress === 'B2' &&
        columnReplacementAddress === 'D2'
          ? number(5)
          : err(ErrorCode.Ref),
    }

    expect(evaluateAstResult(parseFormula('GROUPBY(A1:A5,C1:C5,SUM,3,1)'), context)).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [text('Region'), text('Sales'), text('East'), number(15), text('West'), number(11), text('Total'), number(26)],
    })

    expect(evaluateAstResult(parseFormula('PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)'), context)).toEqual({
      kind: 'array',
      rows: 4,
      cols: 4,
      values: [
        text('Region'),
        text('Widget'),
        text('Gizmo'),
        text('Total'),
        text('East'),
        number(10),
        number(5),
        number(15),
        text('West'),
        number(7),
        number(4),
        number(11),
        text('Total'),
        number(17),
        number(9),
        number(26),
      ],
    })

    expect(evaluateAst(parseFormula('MULTIPLE.OPERATIONS(B5,B3,C4,B2,D2)'), context)).toEqual(number(5))
  })

  it('keeps workbook-special validation failures stable', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales")'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales",1)'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('GETPIVOTDATA(A1:A2,B2)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('GETPIVOTDATA(A1,B2,"Region","East")'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,A1:A2,"East")'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,A1,A2)'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,A1,"East")'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('GROUPBY(A1:A2,B1:B2)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('GROUPBY(1,B1:B2,SUM)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('GROUPBY(A1:A2,LAMBDA(x,x),SUM)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('PIVOTBY(A1:A2,B1:B2,C1:C2)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('PIVOTBY(1,B1:B2,C1:C2,SUM)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('PIVOTBY(A1:A2,B1:B2,LAMBDA(x,x),SUM)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('MULTIPLE.OPERATIONS(B5,B3)'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('MULTIPLE.OPERATIONS(1,B3,C4)'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('MULTIPLE.OPERATIONS(B5,B3,C4,1,E4)'), context)).toEqual(err(ErrorCode.Ref))
    expect(evaluateAst(parseFormula('INDIRECT("R1C1",FALSE())'), context)).toEqual(err(ErrorCode.Value))
    expect(evaluateAst(parseFormula('INDIRECT("A1","x")'), context)).toEqual(err(ErrorCode.Value))
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
