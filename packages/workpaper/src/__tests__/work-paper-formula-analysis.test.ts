import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import {
  calculateWorkPaperFormula,
  compileWorkPaperScalarFormula,
  getWorkPaperNamedExpressionsFromFormula,
  normalizeWorkPaperFormula,
  validateWorkPaperFormula,
  type WorkPaperFormulaAnalysisHooks,
} from '../work-paper-formula-analysis.js'
import { WorkPaperNotAFormulaError, WorkPaperParseError } from '../work-paper-errors.js'

const hooks: WorkPaperFormulaAnalysisHooks = {
  messageOf: (error, fallback) => (error instanceof Error && error.message ? error.message : fallback),
}

const rejectScratchWorkbookCreation = () => {
  throw new Error('scratch workbook should not be built')
}

describe('work paper formula analysis', () => {
  it('normalizes valid formula text with a leading equals sign', () => {
    expect(normalizeWorkPaperFormula('=sum(a1, 2)', hooks)).toBe('=SUM(A1,2)')
  })

  it('requires formula text for normalization and named-expression inspection', () => {
    expect(() => normalizeWorkPaperFormula('SUM(A1)', hooks)).toThrow(WorkPaperNotAFormulaError)
    expect(() => getWorkPaperNamedExpressionsFromFormula('Alpha+Beta', hooks)).toThrow(WorkPaperNotAFormulaError)
  })

  it('wraps parser failures with the public parse error', () => {
    expect(() => normalizeWorkPaperFormula('=SUM(', hooks)).toThrow(WorkPaperParseError)
    expect(() => getWorkPaperNamedExpressionsFromFormula('=SUM(', hooks)).toThrow(WorkPaperParseError)
  })

  it('collects unique named-expression references in sorted order', () => {
    expect(getWorkPaperNamedExpressionsFromFormula('=Beta+Alpha+Beta', hooks)).toEqual(['Alpha', 'Beta'])
  })

  it('evaluates pure scalar formulas without building a scratch workbook', () => {
    expect(
      calculateWorkPaperFormula({
        createWorkbook: rejectScratchWorkbookCreation,
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=SUM(1,2,3)',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(
      calculateWorkPaperFormula({
        createWorkbook: rejectScratchWorkbookCreation,
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=IF(TRUE,"yes","no")',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.String, value: 'yes', stringId: 0 })
    expect(
      calculateWorkPaperFormula({
        createWorkbook: rejectScratchWorkbookCreation,
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=CONCATENATE("baz","-","bar")',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.String, value: 'baz-bar', stringId: 0 })
    expect(
      calculateWorkPaperFormula({
        createWorkbook: rejectScratchWorkbookCreation,
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=MIN(100,120,220)+MAX(100,120,220)',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 320 })
  })

  it('resolves simple named scalars in the pure scalar fast path', () => {
    expect(
      calculateWorkPaperFormula({
        createWorkbook: () => {
          throw new Error('scratch workbook should not be built')
        },
        config: {},
        serializedSheets: {},
        namedExpressions: [
          { name: 'Rate', expression: 0.06 },
          { name: 'Periods', expression: 12 },
          { name: 'Principal', expression: 100000 },
        ],
        formula: '=PMT(Rate/Periods,Periods,Principal)',
        messageOf: hooks.messageOf,
      }),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(-8606.642970708252, 8) })
  })

  it('compiles scalar formulas with cell variable bindings', () => {
    const compiled = compileWorkPaperScalarFormula({
      config: {},
      namedExpressions: [],
      formula: '=IF(A1>0,A1+B1*2,"no")',
      messageOf: hooks.messageOf,
    })

    expect(compiled.variables).toEqual(['A1', 'B1'])
    expect(compiled.evaluate({ A1: 101, B1: 20 })).toEqual({ tag: ValueTag.Number, value: 141 })
    expect(compiled.evaluate({ A1: -1, B1: 20 })).toEqual({ tag: ValueTag.String, value: 'no', stringId: 0 })
  })

  it('rejects workbook-dependent formulas in scalar compile mode', () => {
    expect(() =>
      compileWorkPaperScalarFormula({
        config: {},
        namedExpressions: [],
        formula: '=SUM(A1:A2)',
        messageOf: hooks.messageOf,
      }),
    ).toThrow(WorkPaperParseError)
  })

  it('delegates workbook-dependent formula calculation to the scratch workbook helper', () => {
    const applied: string[] = []

    expect(
      calculateWorkPaperFormula({
        createWorkbook: () => ({
          engine: {
            createSheet: (sheetName) => {
              applied.push(`sheet:${sheetName}`)
            },
            getSpillRanges: () => [],
          },
          registerNamedExpression: () => {},
          requireSheetId: () => 1,
          replaceSheetContent: () => {},
          clearHistoryStacks: () => {},
          applyRawContent: (_address, content) => {
            applied.push(`content:${content}`)
          },
          getRangeValues: () => [],
          getCellValue: () => ({ tag: ValueTag.Number, value: 42 }),
          dispose: () => {},
        }),
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=SUM(A1,2)',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(applied).toContain('content:=SUM(A1,2)')
  })

  it('keeps custom-function formulas on the scratch workbook path', () => {
    const applied: string[] = []

    expect(
      calculateWorkPaperFormula({
        createWorkbook: () => ({
          engine: {
            createSheet: (sheetName) => {
              applied.push(`sheet:${sheetName}`)
            },
            getSpillRanges: () => [],
          },
          registerNamedExpression: () => {},
          requireSheetId: () => 1,
          replaceSheetContent: () => {},
          clearHistoryStacks: () => {},
          applyRawContent: (_address, content) => {
            applied.push(`content:${content}`)
          },
          getRangeValues: () => [],
          getCellValue: () => ({ tag: ValueTag.Number, value: 6 }),
          dispose: () => {},
        }),
        config: {
          functionPlugins: [
            {
              id: 'custom',
              implementedFunctions: { DOUBLE: { method: 'DOUBLE' } },
              functions: {
                DOUBLE: (value) => value,
              },
            },
          ],
        },
        serializedSheets: {},
        namedExpressions: [],
        formula: '=DOUBLE(3)',
        messageOf: hooks.messageOf,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(applied).toContain('content:=DOUBLE(3)')
  })

  it('disposes the scratch workbook when formula calculation fails', () => {
    let disposed = false

    expect(() =>
      calculateWorkPaperFormula({
        createWorkbook: () => ({
          engine: {
            createSheet: () => {},
            getSpillRanges: () => [],
          },
          registerNamedExpression: () => {},
          requireSheetId: () => 1,
          replaceSheetContent: () => {},
          clearHistoryStacks: () => {},
          applyRawContent: () => {
            throw new Error('scratch failed')
          },
          getRangeValues: () => [],
          getCellValue: () => ({ tag: ValueTag.Number, value: 42 }),
          dispose: () => {
            disposed = true
          },
        }),
        config: {},
        serializedSheets: {},
        namedExpressions: [],
        formula: '=SUM(A1,2)',
        messageOf: hooks.messageOf,
      }),
    ).toThrow(WorkPaperParseError)
    expect(disposed).toBe(true)
  })

  it('validates formulas without throwing', () => {
    expect(validateWorkPaperFormula('=SUM(1,2)')).toBe(true)
    expect(validateWorkPaperFormula('SUM(1,2)')).toBe(false)
    expect(validateWorkPaperFormula('=SUM(')).toBe(false)
  })
})
