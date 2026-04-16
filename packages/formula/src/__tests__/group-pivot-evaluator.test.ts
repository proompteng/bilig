import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { evaluateGroupBy, evaluatePivotBy, type MatrixValue } from '../group-pivot-evaluator.js'

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function boolean(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function matrix(rows: number, cols: number, values: CellValue[]): MatrixValue {
  return { rows, cols, values }
}

function sumAggregate(subset: readonly CellValue[]): CellValue {
  return number(subset.reduce((total, value) => total + (value.tag === ValueTag.Number ? value.value : 0), 0))
}

function sumCells(values: readonly CellValue[] | undefined): number {
  return (values ?? []).reduce((total, value) => total + (value.tag === ValueTag.Number ? value.value : 0), 0)
}

describe('group-pivot evaluator', () => {
  it('rejects malformed GROUPBY shapes and options', () => {
    const rows = matrix(2, 1, [text('Region'), text('East')])
    const values = matrix(1, 1, [text('Sales')])

    expect(evaluateGroupBy(rows, values, { aggregate: sumAggregate })).toEqual(error(ErrorCode.Value))
    expect(
      evaluateGroupBy(matrix(2, 1, [text('Region'), text('East')]), matrix(2, 1, [text('Sales'), number(10)]), {
        aggregate: sumAggregate,
        fieldRelationship: 2,
      }),
    ).toEqual(error(ErrorCode.Value))
    expect(
      evaluateGroupBy(matrix(2, 1, [text('Region'), text('East')]), matrix(2, 1, [text('Sales'), number(10)]), {
        aggregate: sumAggregate,
        fieldHeadersMode: 4,
      }),
    ).toEqual(error(ErrorCode.Value))
    expect(
      evaluateGroupBy(matrix(2, 1, [text('Region'), text('East')]), matrix(2, 1, [text('Sales'), number(10)]), {
        aggregate: sumAggregate,
        filterArray: matrix(2, 2, [boolean(true), boolean(true), boolean(true), boolean(true)]),
      }),
    ).toEqual(error(ErrorCode.Value))
  })

  it('supports GROUPBY filtering, descending aggregate sort, and totals-first output', () => {
    const result = evaluateGroupBy(
      matrix(5, 2, [
        text('Region'),
        text('Product'),
        text('East'),
        text('Widget'),
        text('West'),
        text('Widget'),
        text('East'),
        text('Gizmo'),
        text('West'),
        text('Gizmo'),
      ]),
      matrix(5, 2, [text('Sales'), text('Units'), number(10), number(2), number(7), number(1), number(5), number(3), number(4), number(2)]),
      {
        aggregate: sumAggregate,
        fieldHeadersMode: 3,
        totalDepth: -1,
        sortOrder: [-3],
        filterArray: matrix(1, 5, [text('Keep'), boolean(true), boolean(false), boolean(true), boolean(true)]),
      },
    )

    expect(result).toEqual({
      kind: 'array',
      rows: 5,
      cols: 4,
      values: [
        text('Region'),
        text('Product'),
        text('Sales'),
        text('Units'),
        text('Total'),
        empty(),
        number(19),
        number(7),
        text('East'),
        text('Widget'),
        number(10),
        number(2),
        text('East'),
        text('Gizmo'),
        number(5),
        number(3),
        text('West'),
        text('Gizmo'),
        number(4),
        number(2),
      ],
    })
  })

  it('returns a 1x1 empty spill when GROUPBY suppresses headers, totals, and all rows', () => {
    const result = evaluateGroupBy(
      matrix(3, 1, [text('Region'), text('East'), text('West')]),
      matrix(3, 1, [text('Sales'), number(10), number(7)]),
      {
        aggregate: sumAggregate,
        fieldHeadersMode: 1,
        totalDepth: 0,
        filterArray: matrix(2, 1, [boolean(false), boolean(false)]),
      },
    )

    expect(result).toEqual({ kind: 'array', rows: 1, cols: 1, values: [empty()] })
  })

  it('sorts GROUPBY mixed-type keys through the text fallback and appends totals last', () => {
    const result = evaluateGroupBy(
      matrix(3, 1, [error(ErrorCode.Ref), boolean(false), boolean(true)]),
      matrix(3, 1, [number(1), number(1), number(1)]),
      {
        aggregate: sumAggregate,
        fieldHeadersMode: 0,
        totalDepth: 1,
        sortOrder: [1],
      },
    )

    expect(result).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [error(ErrorCode.Ref), number(1), boolean(false), number(1), boolean(true), number(1), text('Total'), number(3)],
    })
  })

  it('keeps GROUPBY insertion order stable when aggregate sort keys tie', () => {
    const result = evaluateGroupBy(matrix(2, 1, [text('East'), text('West')]), matrix(2, 1, [number(1), number(1)]), {
      aggregate: sumAggregate,
      fieldHeadersMode: 0,
      totalDepth: 0,
      sortOrder: [-2],
    })

    expect(result).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('East'), number(1), text('West'), number(1)],
    })
  })

  it('rejects malformed PIVOTBY shapes and options', () => {
    const rowFields = matrix(2, 1, [text('Region'), text('East')])
    const colFields = matrix(3, 1, [text('Product'), text('Widget'), text('Gizmo')])
    const values = matrix(2, 1, [text('Sales'), number(10)])

    expect(evaluatePivotBy(rowFields, colFields, values, { aggregate: sumAggregate })).toEqual(error(ErrorCode.Value))
    expect(
      evaluatePivotBy(
        matrix(2, 1, [text('Region'), text('East')]),
        matrix(2, 1, [text('Product'), text('Widget')]),
        matrix(2, 1, [text('Sales'), number(10)]),
        {
          aggregate: sumAggregate,
          filterArray: matrix(2, 2, [boolean(true), boolean(true), boolean(true), boolean(true)]),
        },
      ),
    ).toEqual(error(ErrorCode.Value))
    expect(
      evaluatePivotBy(
        matrix(2, 1, [text('Region'), text('East')]),
        matrix(2, 1, [text('Product'), text('Widget')]),
        matrix(2, 1, [text('Sales'), number(10)]),
        { aggregate: sumAggregate, fieldHeadersMode: 9 },
      ),
    ).toEqual(error(ErrorCode.Value))
  })

  it('supports PIVOTBY totals-first output with blank row-field labels', () => {
    const result = evaluatePivotBy(
      matrix(3, 1, [text('Region'), text('East'), text('West')]),
      matrix(3, 1, [text('Product'), text('Widget'), text('Gizmo')]),
      matrix(3, 1, [text('Sales'), number(10), number(5)]),
      {
        aggregate: sumAggregate,
        fieldHeadersMode: 1,
        rowTotalDepth: -1,
        colTotalDepth: -1,
      },
    )

    expect(result).toEqual({
      kind: 'array',
      rows: 4,
      cols: 4,
      values: [
        empty(),
        text('Total'),
        text('Widget'),
        text('Gizmo'),
        text('Total'),
        number(15),
        number(10),
        number(5),
        text('East'),
        number(10),
        number(10),
        number(0),
        text('West'),
        number(5),
        number(0),
        number(5),
      ],
    })
  })

  it('applies PIVOTBY filters with totals appended after the data matrix', () => {
    const result = evaluatePivotBy(
      matrix(5, 1, [text('Region'), text('East'), text('West'), text('East'), text('West')]),
      matrix(5, 1, [text('Product'), text('Widget'), text('Widget'), text('Gizmo'), text('Gizmo')]),
      matrix(5, 1, [text('Sales'), number(10), number(7), number(5), number(4)]),
      {
        aggregate: sumAggregate,
        fieldHeadersMode: 3,
        rowTotalDepth: 1,
        colTotalDepth: 1,
        filterArray: matrix(5, 1, [text('Keep'), boolean(true), boolean(false), boolean(true), boolean(true)]),
      },
    )

    expect(result).toEqual({
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
        number(0),
        number(4),
        number(4),
        text('Total'),
        number(10),
        number(9),
        number(19),
      ],
    })
  })

  it('switches PIVOTBY relative totals between row, grand, and column scopes', () => {
    const rowFields = matrix(5, 1, [text('Region'), text('East'), text('West'), text('East'), text('West')])
    const colFields = matrix(5, 1, [text('Product'), text('Widget'), text('Widget'), text('Gizmo'), text('Gizmo')])
    const values = matrix(5, 1, [text('Sales'), number(10), number(7), number(5), number(4)])
    const aggregate = (_subset: readonly CellValue[], totalSet?: readonly CellValue[]) => number(sumCells(totalSet))

    const byColumn = evaluatePivotBy(rowFields, colFields, values, {
      aggregate,
      fieldHeadersMode: 3,
      rowTotalDepth: 0,
      colTotalDepth: 0,
    })
    const byRow = evaluatePivotBy(rowFields, colFields, values, {
      aggregate,
      fieldHeadersMode: 3,
      rowTotalDepth: 0,
      colTotalDepth: 0,
      relativeTo: 1,
    })
    const byGrand = evaluatePivotBy(rowFields, colFields, values, {
      aggregate,
      fieldHeadersMode: 3,
      rowTotalDepth: 0,
      colTotalDepth: 0,
      relativeTo: 2,
    })

    expect(byColumn).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [text('Region'), text('Widget'), text('Gizmo'), text('East'), number(17), number(9), text('West'), number(17), number(9)],
    })
    expect(byRow).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [text('Region'), text('Widget'), text('Gizmo'), text('East'), number(15), number(15), text('West'), number(11), number(11)],
    })
    expect(byGrand).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [text('Region'), text('Widget'), text('Gizmo'), text('East'), number(26), number(26), text('West'), number(26), number(26)],
    })
  })
})
