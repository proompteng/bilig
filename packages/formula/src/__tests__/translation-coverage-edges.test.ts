import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@bilig/protocol'
import type { FormulaNode } from '../ast.js'
import { compileFormula, type CompiledFormula, type ParsedCellReferenceInfo, type ParsedDependencyReference } from '../compiler.js'
import type { JsPlanInstruction } from '../js-evaluator.js'
import {
  buildRelativeFormulaTemplateKeyFromAst,
  rewriteAddressForStructuralTransform,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  serializeFormula,
  translateCompiledFormula,
  translateCompiledFormulaWithoutAst,
  type StructuralAxisTransform,
} from '../translation.js'

function makeCompiledFormula(ast: FormulaNode, overrides: Partial<CompiledFormula> = {}): CompiledFormula {
  const base = compileFormula('1')
  return {
    ...base,
    source: serializeFormula(ast),
    ast,
    optimizedAst: ast,
    deps: [],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    symbolicRefs: [],
    symbolicRanges: [],
    symbolicStrings: [],
    jsPlan: [{ opcode: 'return' }],
    ...overrides,
  }
}

type ExactMatchInstruction = Extract<JsPlanInstruction, { opcode: 'lookup-exact-match' }>
type ApproximateMatchInstruction = Extract<JsPlanInstruction, { opcode: 'lookup-approximate-match' }>

function makeExactMatchInstruction(overrides: Partial<ExactMatchInstruction> = {}): ExactMatchInstruction {
  return {
    opcode: 'lookup-exact-match',
    callee: 'MATCH',
    start: 'A1',
    end: 'A1',
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
    refKind: 'cells',
    searchMode: 1,
    ...overrides,
  }
}

function makeApproximateMatchInstruction(overrides: Partial<ApproximateMatchInstruction> = {}): ApproximateMatchInstruction {
  return {
    opcode: 'lookup-approximate-match',
    callee: 'MATCH',
    start: 'A1',
    end: 'A1',
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
    refKind: 'cells',
    matchMode: 1,
    ...overrides,
  }
}

function withInvalidLookupRefKind<T extends ExactMatchInstruction | ApproximateMatchInstruction>(
  instruction: T,
  refKind: 'rows' | 'cols',
): T {
  Reflect.set(instruction, 'refKind', refKind)
  return instruction
}

function withInvalidTransformKind(transform: StructuralAxisTransform, kind: string): StructuralAxisTransform {
  Reflect.set(transform, 'kind', kind)
  return transform
}

describe('translation coverage edges', () => {
  it('builds template keys for literal, invalid, and invoke-heavy AST shapes', () => {
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'NumberLiteral', value: 7 }, 0, 0)).toBe('n:7')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'BooleanLiteral', value: false }, 0, 0)).toBe('b:0')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'StringLiteral', value: 'text' }, 0, 0)).toBe('s:"text"')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'ErrorLiteral', code: ErrorCode.Ref }, 0, 0)).toBe(`e:${ErrorCode.Ref}`)
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'NameRef', name: 'Rate' }, 0, 0)).toBe('name:Rate')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' }, 0, 0)).toBe(
      'table:Sales[Amount]',
    )
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'CellRef', ref: 'A0' } as FormulaNode, 0, 0)).toBe('cell:.:invalid:A0')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'SpillRef', ref: 'B2' }, 0, 0)).toBe('spill:.:rc1:rr1')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'ColumnRef', ref: '?' } as FormulaNode, 0, 0)).toBe('col:.:invalid:?')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'RowRef', ref: '3' }, 0, 0)).toBe('row:.:r2')
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'RangeRef', refKind: 'rows', start: '2', end: '4' }, 0, 0)).toBe(
      'range:rows:.:r1:r3',
    )
    expect(buildRelativeFormulaTemplateKeyFromAst({ kind: 'RangeRef', refKind: 'cols', start: 'B', end: 'D' }, 0, 0)).toBe(
      'range:cols:.:r1:r3',
    )
    expect(
      buildRelativeFormulaTemplateKeyFromAst({ kind: 'UnaryExpr', operator: '-', argument: { kind: 'NameRef', name: 'Rate' } }, 0, 0),
    ).toBe('unary:-:name:Rate')
    expect(
      buildRelativeFormulaTemplateKeyFromAst(
        {
          kind: 'InvokeExpr',
          callee: { kind: 'NameRef', name: 'Fn' },
          args: [
            { kind: 'NumberLiteral', value: 1 },
            { kind: 'BooleanLiteral', value: false },
          ],
        },
        0,
        0,
      ),
    ).toBe('invoke:name:Fn:n:1|b:0')
  })

  it('translates manual compiled formulas through metadata fallback and nested lambda plan branches', () => {
    const compiled = makeCompiledFormula(
      {
        kind: 'BinaryExpr',
        operator: '+',
        left: { kind: 'ColumnRef', sheetName: 'Sheet1', ref: 'C' },
        right: { kind: 'RowRef', sheetName: 'Sheet1', ref: '5' },
      },
      {
        deps: ['A1', "'Other Sheet'!A1:A3", '2:4', 'B:D'],
        symbolicRefs: ['A1', "'Other Sheet'!A1"],
        symbolicRanges: ["'Other Sheet'!A1:A3", '2:4', 'B:D'],
        parsedDeps: undefined,
        parsedSymbolicRefs: undefined,
        parsedSymbolicRanges: undefined,
        jsPlan: [
          { opcode: 'push-cell', address: 'A1' },
          {
            opcode: 'push-range',
            sheetName: 'Other Sheet',
            start: 'A1',
            end: 'A3',
            refKind: 'cells',
          },
          withInvalidLookupRefKind(
            makeApproximateMatchInstruction({
              start: '2',
              end: '4',
              startRow: 1,
              endRow: 3,
              startCol: 0,
              endCol: 0,
            }),
            'rows',
          ),
          {
            opcode: 'push-lambda',
            params: ['x'],
            body: [{ opcode: 'push-cell', address: 'A1' }, { opcode: 'return' }],
          },
          { opcode: 'return' },
        ],
      },
    )

    const translated = translateCompiledFormula(compiled, 2, 1)

    expect(translated.source).toBe('Sheet1!D+Sheet1!7')
    expect(translated.compiled.deps).toEqual(['B3', "'Other Sheet'!B3:B5", '4:6', 'C:E'])
    expect(translated.compiled.symbolicRefs).toEqual(['B3', "'Other Sheet'!B3"])
    expect(translated.compiled.symbolicRanges).toEqual(["'Other Sheet'!B3:B5", '4:6', 'C:E'])
    expect(translated.compiled.jsPlan).toEqual([
      { opcode: 'push-cell', address: 'B3' },
      {
        opcode: 'push-range',
        sheetName: 'Other Sheet',
        start: 'B3',
        end: 'B5',
        refKind: 'cells',
      },
      {
        opcode: 'lookup-approximate-match',
        callee: 'MATCH',
        start: '2',
        end: '4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'rows',
        matchMode: 1,
      },
      {
        opcode: 'push-lambda',
        params: ['x'],
        body: [{ opcode: 'push-cell', address: 'B3' }, { opcode: 'return' }],
      },
      { opcode: 'return' },
    ])
  })

  it('rewrites manual compiled formulas through axis error, raw fallback, and parsed fallback branches', () => {
    const rowCompiled = makeCompiledFormula({ kind: 'RowRef', sheetName: 'Sheet1', ref: '5' })
    const colCompiled = makeCompiledFormula({ kind: 'ColumnRef', sheetName: 'Sheet1', ref: 'C' })
    const rowInserted = rewriteCompiledFormulaForStructuralTransform(rowCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })
    const colInserted = rewriteCompiledFormulaForStructuralTransform(colCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'column',
      start: 0,
      count: 1,
    })
    const rowDeleted = rewriteCompiledFormulaForStructuralTransform(rowCompiled, 'Sheet1', 'Sheet1', {
      kind: 'delete',
      axis: 'row',
      start: 4,
      count: 1,
    })
    const malformedCellRange = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({
        kind: 'RangeRef',
        refKind: 'cells',
        start: 'A0',
        end: 'B1',
      } as FormulaNode),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      },
    )
    const malformedRowRange = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({
        kind: 'RangeRef',
        refKind: 'rows',
        start: '0',
        end: '2',
      } as FormulaNode),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      },
    )

    const parsedCellReference = {
      kind: 'cell' as const,
      address: 'A1',
      row: 0,
      col: 0,
      rowAbsolute: false,
      colAbsolute: false,
    }
    const parsedRowRange = {
      kind: 'range' as const,
      address: '1:1',
      refKind: 'rows' as const,
      startAddress: '1',
      endAddress: '1',
      startRow: 0,
      endRow: 0,
      startCol: 0,
      endCol: 0,
      startRowAbsolute: false,
      endRowAbsolute: false,
    }
    const rawFallbackCompiled = makeCompiledFormula(
      { kind: 'CellRef', sheetName: 'Sheet1', ref: 'A1' },
      {
        optimizedAst: { kind: 'NameRef', name: 'Stable' },
        deps: ['Sheet1!A1', 'Sheet1!A1:A1', '1:1'],
        symbolicRefs: ['Sheet1!A1'],
        symbolicRanges: ['Sheet1!A1:A1', '1:1'],
        parsedDeps: [parsedCellReference, parsedRowRange],
        parsedSymbolicRefs: [parsedCellReference],
        parsedSymbolicRanges: [parsedRowRange],
        jsPlan: [
          withInvalidLookupRefKind(
            makeExactMatchInstruction({
              start: '1',
              end: '1',
              startRow: 0,
              endRow: 0,
              startCol: 0,
              endCol: 0,
            }),
            'rows',
          ),
          { opcode: 'return' },
        ],
      },
    )
    const rawFallback = rewriteCompiledFormulaForStructuralTransform(rawFallbackCompiled, 'Sheet1', 'Sheet1', {
      kind: 'delete',
      axis: 'row',
      start: 0,
      count: 1,
    })
    const rowLookupCompiled = makeCompiledFormula(
      { kind: 'NumberLiteral', value: 1 },
      {
        jsPlan: [
          withInvalidLookupRefKind(
            makeExactMatchInstruction({
              start: '2',
              end: '4',
              startRow: 1,
              endRow: 3,
              startCol: 0,
              endCol: 0,
            }),
            'rows',
          ),
          { opcode: 'return' },
        ],
      },
    )
    const rowLookupRewritten = rewriteCompiledFormulaForStructuralTransform(rowLookupCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rowInserted.source).toBe('Sheet1!6')
    expect(colInserted.source).toBe('Sheet1!D')
    expect(rowDeleted.source).toBe('#REF!')
    expect(malformedCellRange.source).toBe('#REF!')
    expect(malformedRowRange.source).toBe('#REF!')
    expect(rawFallback.reusedProgram).toBe(true)
    expect(rawFallback.source).toBe('#REF!')
    expect(rawFallback.compiled.deps).toEqual(rawFallbackCompiled.deps)
    expect(rawFallback.compiled.symbolicRefs).toEqual(rawFallbackCompiled.symbolicRefs)
    expect(rawFallback.compiled.symbolicRanges).toEqual(rawFallbackCompiled.symbolicRanges)
    expect(rawFallback.compiled.parsedSymbolicRefs).toEqual(rawFallbackCompiled.parsedSymbolicRefs)
    expect(rawFallback.compiled.parsedSymbolicRanges).toEqual(rawFallbackCompiled.parsedSymbolicRanges)
    expect(rawFallback.compiled.jsPlan).toEqual(rawFallbackCompiled.jsPlan)
    expect(rowLookupRewritten.compiled.jsPlan).toEqual(rowLookupCompiled.jsPlan)
  })

  it('handles stable row refs, sparse parsed metadata, and direct lookup plan translation', () => {
    const rowUnchanged = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({ kind: 'RowRef', sheetName: 'Sheet1', ref: '5' }),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'column',
        start: 0,
        count: 1,
      },
    )
    const booleanStable = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({ kind: 'BooleanLiteral', value: false }),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      },
    )
    const stringStable = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({ kind: 'StringLiteral', value: 'keep' }),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      },
    )
    const unaryStable = rewriteCompiledFormulaForStructuralTransform(
      makeCompiledFormula({
        kind: 'UnaryExpr',
        operator: '-',
        argument: { kind: 'NumberLiteral', value: 1 },
      }),
      'Sheet1',
      'Sheet1',
      {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      },
    )
    const sparseCellRefs: NonNullable<CompiledFormula['parsedSymbolicRefs']> = [
      {
        address: 'A1',
        row: 0,
        col: 0,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ]
    sparseCellRefs.length = 2
    const sparseRangeRefs: NonNullable<CompiledFormula['parsedSymbolicRanges']> = [
      {
        address: '2:4',
        kind: 'range',
        refKind: 'rows',
        startAddress: '2',
        endAddress: '4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        startRowAbsolute: false,
        endRowAbsolute: false,
      },
    ]
    sparseRangeRefs.length = 2
    const sparseTranslated = translateCompiledFormulaWithoutAst(
      makeCompiledFormula(
        { kind: 'NumberLiteral', value: 1 },
        {
          symbolicRefs: ['Sheet1!A1'],
          symbolicRanges: ['2:4'],
          parsedSymbolicRefs: sparseCellRefs,
          parsedSymbolicRanges: sparseRangeRefs,
          jsPlan: [
            { opcode: 'push-cell', address: 'A1' },
            {
              opcode: 'call',
              callee: 'WRAP',
              argc: 2,
              argRefs: [
                { kind: 'cell', address: 'A1' },
                { kind: 'range', start: '2', end: '4', refKind: 'rows' },
              ],
            },
            { opcode: 'return' },
          ],
        },
      ),
      1,
      0,
    )
    const translatedLookup = translateCompiledFormula(
      makeCompiledFormula(
        { kind: 'NumberLiteral', value: 1 },
        {
          jsPlan: [
            {
              opcode: 'lookup-exact-match',
              callee: 'MATCH',
              start: 'A1',
              end: 'A3',
              startRow: 0,
              endRow: 2,
              startCol: 0,
              endCol: 0,
              refKind: 'cells',
              searchMode: 1,
            },
            {
              opcode: 'call',
              callee: 'WRAP',
              argc: 1,
              argRefs: [{ kind: 'range' }],
            },
            { opcode: 'return' },
          ],
        },
      ),
      1,
      0,
    )

    expect(rowUnchanged.source).toBe('Sheet1!5')
    expect(booleanStable.reusedProgram).toBe(true)
    expect(stringStable.reusedProgram).toBe(true)
    expect(unaryStable.reusedProgram).toBe(true)
    expect(
      rewriteAddressForStructuralTransform('A1', {
        kind: 'move',
        axis: 'row',
        start: 5,
        count: 1,
        target: 1,
      }),
    ).toBe('A1')
    expect(sparseTranslated.compiled.symbolicRefs[0]).toBe('A2')
    expect(sparseTranslated.compiled.symbolicRanges[0]).toBe('3:5')
    expect(sparseTranslated.compiled.jsPlan).toEqual([
      { opcode: 'push-cell', address: 'A2' },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 2,
        argRefs: [
          { kind: 'cell', address: 'A2' },
          { kind: 'range', start: '3', end: '5', refKind: 'rows' },
        ],
      },
      { opcode: 'return' },
    ])
    expect(translatedLookup.compiled.jsPlan).toEqual([
      {
        opcode: 'lookup-exact-match',
        callee: 'MATCH',
        start: 'A2',
        end: 'A4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'cells',
        searchMode: 1,
      },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 1,
        argRefs: [{ kind: 'range' }],
      },
      { opcode: 'return' },
    ])
  })

  it('falls back when AST-free translation cannot rewrite manual lookup and arg-ref metadata', () => {
    const compiled = makeCompiledFormula(
      { kind: 'NumberLiteral', value: 1 },
      {
        symbolicRanges: ['2:4'],
        parsedSymbolicRanges: [
          {
            address: '2:4',
            kind: 'range',
            refKind: 'rows',
            startAddress: '2',
            endAddress: '4',
            startRow: 1,
            endRow: 3,
            startCol: 0,
            endCol: 0,
            startRowAbsolute: false,
            endRowAbsolute: false,
          },
        ],
        jsPlan: [
          withInvalidLookupRefKind(
            makeExactMatchInstruction({
              start: '2',
              end: '4',
              startRow: 1,
              endRow: 3,
              startCol: 0,
              endCol: 0,
            }),
            'rows',
          ),
          {
            opcode: 'call',
            callee: 'WRAP',
            argc: 2,
            argRefs: [{ kind: 'cell' }, { kind: 'range' }],
          },
          { opcode: 'return' },
        ],
      },
    )

    const translated = translateCompiledFormulaWithoutAst(compiled, 2, 0)

    expect(translated.compiled.symbolicRanges).toEqual(['4:6'])
    expect(translated.compiled.jsPlan).toEqual([
      {
        opcode: 'lookup-exact-match',
        callee: 'MATCH',
        start: '2',
        end: '4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'rows',
        searchMode: 1,
      },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 2,
        argRefs: [{ kind: 'cell' }, { kind: 'range' }],
      },
      { opcode: 'return' },
    ])
  })

  it('preserves invalid sparse parsed refs when AST-free translation cannot normalize them', () => {
    const invalidDependency: ParsedDependencyReference = {
      kind: 'cell',
      address: 'NotACell',
      explicitSheet: true,
    }
    const invalidCellRef: ParsedCellReferenceInfo = {
      address: 'NotACell',
      explicitSheet: true,
    }
    const translated = translateCompiledFormulaWithoutAst(
      makeCompiledFormula(
        { kind: 'NumberLiteral', value: 1 },
        {
          deps: ['NotACell'],
          symbolicRefs: ['NotACell'],
          parsedDeps: [invalidDependency],
          parsedSymbolicRefs: [invalidCellRef],
        },
      ),
      3,
      4,
    )

    expect(translated.compiled.deps).toEqual(['NotACell'])
    expect(translated.compiled.symbolicRefs).toEqual(['NotACell'])
  })

  it('throws for invalid manual references and unsupported structural transforms', () => {
    const unsupportedTransform = withInvalidTransformKind({ kind: 'insert', axis: 'row', start: 0, count: 1 }, 'noop')

    expect(() => translateCompiledFormula(makeCompiledFormula({ kind: 'CellRef', ref: 'A0' } as FormulaNode), 0, 0)).toThrow(
      "Invalid cell reference 'A0'",
    )
    expect(() => translateCompiledFormula(makeCompiledFormula({ kind: 'ColumnRef', ref: '1' } as FormulaNode), 0, 0)).toThrow(
      "Invalid column reference '1'",
    )
    expect(() => translateCompiledFormula(makeCompiledFormula({ kind: 'RowRef', ref: 'A' } as FormulaNode), 0, 0)).toThrow(
      "Invalid row reference 'A'",
    )
    expect(() => rewriteAddressForStructuralTransform('A1', unsupportedTransform)).toThrow('Unexpected value: [object Object]')
    expect(() => rewriteRangeForStructuralTransform('A1', 'A2', unsupportedTransform)).toThrow('Unexpected value: [object Object]')
  })
})
