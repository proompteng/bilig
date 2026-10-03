import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { WorkPaperNamedExpressionNameIsInvalidError, WorkPaperNoRelativeAddressesAllowedError } from '../work-paper-errors.js'
import {
  cloneNamedExpressionValue,
  compareWorkPaperNamedExpressionChanges,
  captureWorkPaperNamedExpressionValueSnapshot,
  computeWorkPaperNamedExpressionChanges,
  createInternalNamedExpressionRecord,
  evaluateWorkPaperNamedExpression,
  listWorkPaperNamedExpressions,
  serializeWorkPaperNamedExpressions,
  trySimpleWorkPaperNamedExpressionDefinedNameSnapshot,
  validateWorkPaperNamedExpression,
  workPaperCellSnapshotToRawContent,
  workPaperNamedExpressionToDefinedNameSnapshot,
} from '../work-paper-named-expression-helpers.js'
import type { WorkPaperChange } from '../work-paper-types.js'

function namedExpressionChange(name: string, scope?: number): WorkPaperChange {
  return {
    kind: 'named-expression',
    name,
    oldValue: undefined,
    newValue: { tag: ValueTag.Number, value: 1 },
    ...(scope === undefined ? {} : { scope }),
  }
}

describe('work paper named expression helpers', () => {
  it('clones scalar and matrix named expression values', () => {
    const scalar = { tag: ValueTag.String, value: 'Rate', stringId: 3 } as const
    const matrix = [[scalar]]

    expect(cloneNamedExpressionValue(scalar)).toEqual(scalar)
    expect(cloneNamedExpressionValue(scalar)).not.toBe(scalar)
    const clonedMatrix = cloneNamedExpressionValue(matrix)
    expect(clonedMatrix).toEqual(matrix)
    expect(clonedMatrix).not.toBe(matrix)
    expect(Array.isArray(clonedMatrix)).toBe(true)
    if (!Array.isArray(clonedMatrix)) {
      throw new Error('Expected cloned named expression value to stay a matrix')
    }
    expect(clonedMatrix[0]).not.toBe(matrix[0])
  })

  it('orders named expression changes by scope then name', () => {
    const changes = [namedExpressionChange('Zulu', 2), namedExpressionChange('Alpha'), namedExpressionChange('Beta', 1)]

    expect(
      changes.toSorted(compareWorkPaperNamedExpressionChanges).map((change) => `${change.scope ?? 'workbook'}:${change.name}`),
    ).toEqual(['workbook:Alpha', '1:Beta', '2:Zulu'])
  })

  it('captures and computes named expression value snapshots', () => {
    const expressions = [
      createInternalNamedExpressionRecord({ name: 'Zulu', expression: '=2', scope: 2 }),
      createInternalNamedExpressionRecord({ name: 'Rate', expression: '=1' }),
    ]
    const before = captureWorkPaperNamedExpressionValueSnapshot(expressions, (expression) => ({
      tag: ValueTag.Number,
      value: expression.publicName === 'Rate' ? 1 : 2,
    }))
    const after = captureWorkPaperNamedExpressionValueSnapshot(expressions, (expression) => ({
      tag: ValueTag.Number,
      value: expression.publicName === 'Rate' ? 3 : 2,
    }))
    const clonedRate = after.get('workbook:RATE')

    expect(clonedRate).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      computeWorkPaperNamedExpressionChanges({
        beforeNames: before,
        afterNames: after,
        expressionsByKey: new Map(
          expressions.map((expression) => [`${expression.scope ?? 'workbook'}:${expression.normalizedName}`, expression]),
        ),
      }),
    ).toEqual([
      {
        kind: 'named-expression',
        name: 'Rate',
        scope: undefined,
        newValue: { tag: ValueTag.Number, value: 3 },
      },
    ])
  })

  it('leaves non-named expression changes equivalent', () => {
    const cellChange: WorkPaperChange = {
      kind: 'cell',
      address: { sheet: 1, row: 0, col: 0 },
      sheetName: 'Sheet1',
      a1: 'A1',
      newValue: { tag: ValueTag.Empty },
    }

    expect(compareWorkPaperNamedExpressionChanges(cellChange, namedExpressionChange('Rate'))).toBe(0)
  })

  it('validates named expression names and rejects relative formula references', () => {
    expect(() =>
      validateWorkPaperNamedExpression({
        expressionName: '1bad',
        expression: 1,
        requireScope: () => {},
        messageOf: (_error, fallback) => fallback,
      }),
    ).toThrow(WorkPaperNamedExpressionNameIsInvalidError)
    expect(() =>
      validateWorkPaperNamedExpression({
        expressionName: 'Rate',
        expression: '=A1+1',
        requireScope: () => {},
        messageOf: (_error, fallback) => fallback,
      }),
    ).toThrow(WorkPaperNoRelativeAddressesAllowedError)
    expect(() =>
      validateWorkPaperNamedExpression({
        expressionName: 'Rate',
        expression: '=SUM($A$1:$A$3)',
        scope: 7,
        requireScope: (scope) => {
          expect(scope).toBe(7)
        },
        messageOf: (_error, fallback) => fallback,
      }),
    ).not.toThrow()
  })

  it('creates internal named-expression records and defined-name snapshots', () => {
    const record = createInternalNamedExpressionRecord({
      name: ' Rate ',
      expression: '=1+2',
      scope: 3,
      options: { hidden: true },
    })

    expect(record).toEqual({
      publicName: 'Rate',
      normalizedName: 'RATE',
      internalName: '__BILIG_WORKPAPER_SCOPE_3_RATE',
      scope: 3,
      expression: '=1+2',
      options: { hidden: true },
    })
    expect(
      workPaperNamedExpressionToDefinedNameSnapshot({
        expression: 7,
        defaultScopeId: 1,
        rewriteFormulaForStorage: (formula) => formula,
      }),
    ).toBe(7)
    expect(
      workPaperNamedExpressionToDefinedNameSnapshot({
        expression: '=2',
        defaultScopeId: 1,
        rewriteFormulaForStorage: (formula) => `stored:${formula}`,
      }),
    ).toEqual({ kind: 'formula', formula: '=2' })
    expect(
      workPaperNamedExpressionToDefinedNameSnapshot({
        expression: '=SUM($A$1:$A$3)',
        defaultScopeId: 1,
        rewriteFormulaForStorage: (formula, ownerSheetId) => `stored:${ownerSheetId}:${formula}`,
      }),
    ).toEqual({ kind: 'formula', formula: '=stored:1:SUM($A$1:$A$3)' })
  })

  it('builds simple scalar defined-name snapshots without scope rewriting', () => {
    expect(trySimpleWorkPaperNamedExpressionDefinedNameSnapshot('=3')).toEqual({ kind: 'formula', formula: '=3' })
    expect(trySimpleWorkPaperNamedExpressionDefinedNameSnapshot(' plain ')).toBe(' plain ')
    expect(trySimpleWorkPaperNamedExpressionDefinedNameSnapshot('=SUM(A1:A3)')).toBeUndefined()
  })

  it('lists and serializes public named expressions in stable order', () => {
    const expressions = [
      createInternalNamedExpressionRecord({ name: 'Tax', expression: '=2', scope: 2 }),
      createInternalNamedExpressionRecord({ name: 'Rate', expression: '=1', scope: 1, options: { hidden: true } }),
      createInternalNamedExpressionRecord({ name: 'Global', expression: 3 }),
    ]

    expect(listWorkPaperNamedExpressions(expressions, 1)).toEqual(['Rate'])
    expect(listWorkPaperNamedExpressions(expressions)).toEqual(['Global'])
    expect(serializeWorkPaperNamedExpressions(expressions)).toEqual([
      { name: 'Global', expression: 3, scope: undefined, options: undefined },
      { name: 'Rate', expression: '=1', scope: 1, options: { hidden: true } },
      { name: 'Tax', expression: '=2', scope: 2, options: undefined },
    ])
  })

  it('evaluates simple named expressions and serializes cell snapshots to raw content', () => {
    const record = createInternalNamedExpressionRecord({ name: 'Flag', expression: '=TRUE' })
    expect(evaluateWorkPaperNamedExpression(record, () => ({ tag: ValueTag.Number, value: 99 }))).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(
      evaluateWorkPaperNamedExpression(createInternalNamedExpressionRecord({ name: 'Text', expression: 'plain' }), () => ({
        tag: ValueTag.Number,
        value: 99,
      })),
    ).toEqual({ tag: ValueTag.String, value: 'plain', stringId: 0 })

    const formulaCell: CellSnapshot = {
      sheetName: 'Sheet1',
      address: 'A1',
      formula: 'stored_formula',
      value: { tag: ValueTag.Number, value: 3 },
      flags: 0,
      version: 1,
    }
    expect(
      workPaperCellSnapshotToRawContent({
        cell: formulaCell,
        ownerSheetId: 1,
        restorePublicFormula: (formula, ownerSheetId) => `${formula}:${ownerSheetId}`,
      }),
    ).toBe('=stored_formula:1')
    expect(
      workPaperCellSnapshotToRawContent({
        cell: { ...formulaCell, formula: undefined, input: '=A1', value: { tag: ValueTag.String, value: 'shown', stringId: 1 } },
        ownerSheetId: 1,
        restorePublicFormula: (formula) => formula,
      }),
    ).toBe('=A1')
  })
})
