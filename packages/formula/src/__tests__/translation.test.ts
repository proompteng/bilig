import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@bilig/protocol'
import type { FormulaNode } from '../ast.js'
import { compileFormula, type CompiledFormula } from '../compiler.js'
import type { JsPlanInstruction } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'
import {
  buildRelativeFormulaTemplateKey,
  buildRelativeFormulaTemplateKeyFromAst,
  renameFormulaSheetReferences,
  translateCompiledFormula,
  rewriteAddressForStructuralTransform,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  serializeFormula,
  translateFormulaReferences,
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

describe('translateFormulaReferences', () => {
  it('shifts relative cell references', () => {
    expect(translateFormulaReferences('A1+B2', 2, 3)).toBe('D3+E4')
  })

  it('preserves absolute anchors while shifting relative axes', () => {
    expect(translateFormulaReferences('$A1+B$2+$C$3', 4, 5)).toBe('$A5+G$2+$C$3')
  })

  it('shifts ranges, row refs, and column refs', () => {
    expect(translateFormulaReferences('SUM(A1:B2)+SUM(C:C)+SUM(3:3)', 1, 2)).toBe('SUM(C2:D3)+SUM(E:E)+SUM(4:4)')
  })

  it('shifts sheet-qualified references without dropping the sheet name', () => {
    expect(translateFormulaReferences("'My Sheet'!A1+Sheet2!B$3", 2, 1)).toBe("'My Sheet'!B3+Sheet2!C$3")
  })

  it('preserves mixed anchors across mixed cell, column, and row ranges', () => {
    expect(translateFormulaReferences('SUM($A1:B$2,$C:$D,$5:6)', 2, 3)).toBe('SUM($A3:E$2,$C:$D,$5:8)')
  })

  it('keeps quoted sheet prefixes and nested precedence intact for mixed references', () => {
    expect(translateFormulaReferences("('My Sheet'!$A1+Sheet2!B$2)*SUM('My Sheet'!$C:$D,Sheet2!3:$4)", 4, 2)).toBe(
      "('My Sheet'!$A5+Sheet2!D$2)*SUM('My Sheet'!$C:$D,Sheet2!7:$4)",
    )
  })

  it('translates spill refs, unary expressions, and invoke expressions through the public API', () => {
    expect(translateFormulaReferences('-A1+A1#', 2, 1)).toBe('-B3+B3#')
    expect(translateFormulaReferences('LAMBDA(x,x+1)(A1)', 1, 2)).toBe('LAMBDA(x,x+1)(C2)')
  })

  it('translates standalone row and column references plus axis ranges', () => {
    expect(translateFormulaReferences('SUM(A:A)+SUM(2:2)', 3, 2)).toBe('SUM(C:C)+SUM(5:5)')
    expect(translateFormulaReferences('A:A+2:4', 1, 1)).toBe('B:B+3:5')
  })

  it('throws when a relative axis would move outside worksheet bounds even if the other axis is absolute', () => {
    expect(() => translateFormulaReferences('$A1+B$1', -1, -2)).toThrow('Translated reference moved outside worksheet bounds: $A1')
  })

  it('throws when standalone row and column refs move outside bounds', () => {
    expect(() => translateFormulaReferences('A1', -1, 0)).toThrow('Translated reference moved outside worksheet bounds: A1')
    expect(() => translateFormulaReferences('A:A', 0, -1)).toThrow('Translated reference moved outside worksheet bounds: A')
    expect(() => translateFormulaReferences('1:1', -1, 0)).toThrow('Translated reference moved outside worksheet bounds: 1')
  })

  it('builds one template key for repeated relative formula shapes', () => {
    expect(buildRelativeFormulaTemplateKey('A1+B1', 0, 2)).toBe(buildRelativeFormulaTemplateKey('A2+B2', 1, 2))
    expect(buildRelativeFormulaTemplateKey('SUM(A1:A1)', 0, 4)).not.toBe(buildRelativeFormulaTemplateKey('SUM(A1:A2)', 1, 4))
    const ast = parseFormula('A3+B$1')
    expect(buildRelativeFormulaTemplateKeyFromAst(ast, 2, 2)).toBe(buildRelativeFormulaTemplateKey('A3+B$1', 2, 2))
  })

  it('rewrites row references for structural inserts and deletes', () => {
    expect(
      rewriteFormulaForStructuralTransform('SUM(A1:A2)', 'Sheet1', 'Sheet1', {
        kind: 'insert',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBe('SUM(A1:A3)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(A1:A3)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBe('SUM(A1:A2)')
  })

  it('rewrites column references for structural moves', () => {
    expect(
      rewriteFormulaForStructuralTransform('A1', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'column',
        start: 0,
        count: 1,
        target: 2,
      }),
    ).toBe('C1')
    expect(
      rewriteFormulaForStructuralTransform('C1', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'column',
        start: 2,
        count: 1,
        target: 0,
      }),
    ).toBe('A1')
  })

  it('collapses deleted references to surviving ranges or #REF', () => {
    expect(
      rewriteFormulaForStructuralTransform('SUM(A1:B1)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'column',
        start: 0,
        count: 1,
      }),
    ).toBe('SUM(A1:A1)')
    expect(
      rewriteFormulaForStructuralTransform('B2', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'column',
        start: 1,
        count: 1,
      }),
    ).toBe('#REF!')
  })

  it('skips structural rewrites for formulas outside the target sheet', () => {
    expect(
      rewriteFormulaForStructuralTransform('A1+Sheet2!B2', 'Sheet1', 'OtherSheet', {
        kind: 'insert',
        axis: 'row',
        start: 1,
        count: 2,
      }),
    ).toBe('A1+Sheet2!B2')
  })

  it('keeps row refs on column transforms and column refs on row transforms', () => {
    expect(
      rewriteFormulaForStructuralTransform('3:5', 'Sheet1', 'Sheet1', {
        kind: 'insert',
        axis: 'column',
        start: 1,
        count: 2,
      }),
    ).toBe('3:5')
    expect(
      rewriteFormulaForStructuralTransform('B:D', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBe('B:D')
  })

  it('rewrites unary expressions during structural transforms', () => {
    expect(
      rewriteFormulaForStructuralTransform('-A1', 'Sheet1', 'Sheet1', {
        kind: 'insert',
        axis: 'column',
        start: 0,
        count: 1,
      }),
    ).toBe('-B1')
  })

  it('rewrites moved cell ranges in both directions', () => {
    expect(
      rewriteRangeForStructuralTransform('D1', 'E1', {
        kind: 'move',
        axis: 'column',
        start: 3,
        count: 1,
        target: 1,
      }),
    ).toEqual({ startAddress: 'B1', endAddress: 'E1' })
    expect(
      rewriteRangeForStructuralTransform('B1', 'D1', {
        kind: 'move',
        axis: 'column',
        start: 1,
        count: 1,
        target: 3,
      }),
    ).toEqual({ startAddress: 'B1', endAddress: 'D1' })
  })

  it('rewrites single-cell addresses and throws for invalid address inputs', () => {
    expect(
      rewriteAddressForStructuralTransform('B2', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBeUndefined()
    expect(
      rewriteAddressForStructuralTransform('A4', {
        kind: 'move',
        axis: 'row',
        start: 1,
        count: 1,
        target: 3,
      }),
    ).toBe('A3')
    expect(() =>
      rewriteAddressForStructuralTransform('bad', {
        kind: 'move',
        axis: 'column',
        start: 1,
        count: 1,
        target: 2,
      }),
    ).toThrow("Invalid cell reference 'bad'")
  })

  it('rewrites ranges across structural inserts, deletes, and throws on bad references', () => {
    expect(
      rewriteRangeForStructuralTransform('A1', 'A4', {
        kind: 'insert',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toEqual({ startAddress: 'A1', endAddress: 'A5' })
    expect(
      rewriteRangeForStructuralTransform('A1', 'A4', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toEqual({ startAddress: 'A1', endAddress: 'A3' })
    expect(() =>
      rewriteRangeForStructuralTransform('A1', 'bad', {
        kind: 'move',
        axis: 'column',
        start: 0,
        count: 1,
        target: 1,
      }),
    ).toThrow('Invalid range reference')
    expect(
      rewriteRangeForStructuralTransform('A1', 'A1', {
        kind: 'delete',
        axis: 'row',
        start: 0,
        count: 1,
      }),
    ).toBeUndefined()
  })

  it('serializes literals, structured refs, spill refs, invokes, and precedence-sensitive binaries', () => {
    expect(serializeFormula({ kind: 'BooleanLiteral', value: false })).toBe('FALSE')
    expect(serializeFormula({ kind: 'StringLiteral', value: 'a"b' })).toBe('"a""b"')
    expect(serializeFormula({ kind: 'ErrorLiteral', code: ErrorCode.Spill })).toBe('#SPILL!')
    expect(
      serializeFormula({
        kind: 'StructuredRef',
        tableName: 'Sales',
        columnName: 'Amount',
      }),
    ).toBe('Sales[Amount]')
    expect(
      serializeFormula({
        kind: 'SpillRef',
        sheetName: 'My Sheet',
        ref: 'A1',
      }),
    ).toBe("'My Sheet'!A1#")
    expect(
      serializeFormula({
        kind: 'InvokeExpr',
        callee: { kind: 'NameRef', name: 'fn' },
        args: [{ kind: 'NumberLiteral', value: 1 }],
      }),
    ).toBe('(fn)(1)')
    expect(
      serializeFormula({
        kind: 'BinaryExpr',
        operator: '^',
        left: {
          kind: 'BinaryExpr',
          operator: '^',
          left: { kind: 'NumberLiteral', value: 2 },
          right: { kind: 'NumberLiteral', value: 3 },
        },
        right: { kind: 'NumberLiteral', value: 4 },
      }),
    ).toBe('(2^3)^4')
    expect(
      serializeFormula({
        kind: 'BinaryExpr',
        operator: '-',
        left: { kind: 'NumberLiteral', value: 1 },
        right: {
          kind: 'BinaryExpr',
          operator: '-',
          left: { kind: 'NumberLiteral', value: 2 },
          right: { kind: 'NumberLiteral', value: 3 },
        },
      }),
    ).toBe('1-(2-3)')
  })

  it('renames sheet references inside nested unary, call, and invoke expressions', () => {
    expect(
      renameFormulaSheetReferences("SUM(-'Old Sheet'!A1,MAP('Old Sheet'!B:B,LAMBDA(x,x)))+'Old Sheet'!C1#", 'Old Sheet', 'New Sheet'),
    ).toBe("SUM(-'New Sheet'!A1,MAP('New Sheet'!B:B,LAMBDA(x,x)))+'New Sheet'!C1#")
    expect(renameFormulaSheetReferences("LAMBDA(x,x)('Old Sheet'!A1)", 'Old Sheet', 'New Sheet')).toBe("LAMBDA(x,x)('New Sheet'!A1)")
  })

  it('rewrites spill refs and axis ranges while collapsing deleted axis refs to #REF', () => {
    expect(
      rewriteFormulaForStructuralTransform("'My Sheet'!A1#+SUM(2:2)", 'My Sheet', 'My Sheet', {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      }),
    ).toBe("'My Sheet'!A2#+SUM(3:3)")
    expect(
      rewriteFormulaForStructuralTransform('SUM(B:B)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'column',
        start: 1,
        count: 1,
      }),
    ).toBe('SUM(#REF!)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(2:2)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBe('SUM(#REF!)')
  })

  it('rewrites invoke expressions while leaving unaffected axis refs on other sheets intact', () => {
    expect(
      rewriteFormulaForStructuralTransform('LAMBDA(x,x+A1)(B2)', 'Sheet1', 'Sheet1', {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      }),
    ).toBe('LAMBDA(x,x+A2)(B3)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(Sheet2!C:C)+SUM(Sheet2!$4:$4)', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'column',
        start: 0,
        count: 1,
        target: 2,
      }),
    ).toBe('SUM(Sheet2!C:C)+SUM(Sheet2!$4:$4)')
  })

  it('renames explicit sheet references while preserving quoting', () => {
    expect(renameFormulaSheetReferences("SUM(Data!A1:B2)+'Old Sheet'!C3+Current!D4", 'Old Sheet', 'Q1')).toBe(
      'SUM(Data!A1:B2)+Q1!C3+Current!D4',
    )
    expect(renameFormulaSheetReferences("'It''s here'!A1+Sheet2!B2", "It's here", 'New Sheet')).toBe("'New Sheet'!A1+Sheet2!B2")
  })

  it('throws when translated row or column refs move outside worksheet bounds', () => {
    expect(() => translateFormulaReferences('SUM(A:A)', 0, -1)).toThrow('Translated reference moved outside worksheet bounds: A')
    expect(() => translateFormulaReferences('SUM(1:1)', -1, 0)).toThrow('Translated reference moved outside worksheet bounds: 1')
  })

  it('serializes sheet-qualified refs, unary expressions, and call expressions', () => {
    expect(serializeFormula({ kind: 'CellRef', sheetName: 'Sheet 2', ref: '$B3' })).toBe("'Sheet 2'!$B3")
    expect(serializeFormula({ kind: 'ColumnRef', sheetName: 'Sheet2', ref: 'C' })).toBe('Sheet2!C')
    expect(serializeFormula({ kind: 'RowRef', sheetName: 'Sheet2', ref: '$4' })).toBe('Sheet2!$4')
    expect(
      serializeFormula({
        kind: 'RangeRef',
        sheetName: 'Sheet2',
        refKind: 'cells',
        start: 'A1',
        end: 'B2',
      }),
    ).toBe('Sheet2!A1:B2')
    expect(
      serializeFormula({
        kind: 'UnaryExpr',
        operator: '-',
        argument: {
          kind: 'BinaryExpr',
          operator: '+',
          left: { kind: 'NumberLiteral', value: 1 },
          right: { kind: 'NumberLiteral', value: 2 },
        },
      }),
    ).toBe('-(1+2)')
    expect(
      serializeFormula({
        kind: 'CallExpr',
        callee: 'SUM',
        args: [
          { kind: 'CellRef', sheetName: null, ref: 'A1' },
          { kind: 'RangeRef', sheetName: null, refKind: 'cols', start: 'B', end: 'C' },
        ],
      }),
    ).toBe('SUM(A1,B:C)')
  })

  it('rewrites move transforms across unaffected, shifted, and invalid intervals', () => {
    expect(
      rewriteFormulaForStructuralTransform('SUM(A:C)', 'Sheet1', 'Sheet1', {
        kind: 'insert',
        axis: 'column',
        start: 5,
        count: 2,
      }),
    ).toBe('SUM(A:C)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(C:E)', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'column',
        start: 1,
        count: 2,
        target: 4,
      }),
    ).toBe('SUM(B:F)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(5:7)', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'row',
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe('SUM(2:7)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(5:6)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 4,
        count: 2,
      }),
    ).toBe('SUM(#REF!)')
  })

  it('rewrites point references and intervals across move and delete edge segments', () => {
    expect(
      rewriteFormulaForStructuralTransform('3:3', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'row',
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe('5:5')
    expect(
      rewriteFormulaForStructuralTransform('10:10', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'row',
        start: 4,
        count: 2,
        target: 1,
      }),
    ).toBe('10:10')
    expect(
      rewriteFormulaForStructuralTransform('SUM(5:7)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 1,
        count: 1,
      }),
    ).toBe('SUM(4:6)')
    expect(
      rewriteFormulaForStructuralTransform('SUM(5:7)', 'Sheet1', 'Sheet1', {
        kind: 'delete',
        axis: 'row',
        start: 10,
        count: 1,
      }),
    ).toBe('SUM(5:7)')
  })

  it('covers remaining structural transformation edge cases for intervals and sheet quoting', () => {
    expect(
      rewriteFormulaForStructuralTransform('SUM(A1:C3)', 'Sheet1', 'Sheet1', {
        kind: 'move',
        axis: 'row',
        start: 0,
        count: 1,
        target: 5,
      }),
    ).toBe('SUM(A1:C6)')

    expect(
      rewriteFormulaForStructuralTransform("'Sheet With Spaces'!A1", 'Sheet1', 'Sheet With Spaces', {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      }),
    ).toBe("'Sheet With Spaces'!A2")

    expect(
      rewriteFormulaForStructuralTransform("'It''s a sheet'!A1", 'Sheet1', "It's a sheet", {
        kind: 'insert',
        axis: 'row',
        start: 0,
        count: 1,
      }),
    ).toBe("'It''s a sheet'!A2")
  })

  it('serializes fallback errors, nested invoke callees, and quoted sheet names', () => {
    const unknownErrorNode: { kind: 'ErrorLiteral'; code: ErrorCode } = {
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    }
    Reflect.set(unknownErrorNode, 'code', 999)
    expect(serializeFormula(unknownErrorNode)).toBe('#ERROR!')
    expect(
      serializeFormula({
        kind: 'InvokeExpr',
        callee: {
          kind: 'CallExpr',
          callee: 'LAMBDA',
          args: [
            { kind: 'NameRef', name: 'x' },
            { kind: 'NameRef', name: 'x' },
          ],
        },
        args: [{ kind: 'NumberLiteral', value: 4 }],
      }),
    ).toBe('LAMBDA(x,x)(4)')
    expect(
      serializeFormula({
        kind: 'CellRef',
        sheetName: 'Sales.Q1',
        ref: '$C$5',
      }),
    ).toBe('Sales.Q1!$C$5')
  })

  it('reuses compiled programs for shape-stable structural rewrites', () => {
    const compiled = compileFormula('SUM(A1:A10)')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 5,
      count: 1,
    })

    expect(rewritten.source).toBe('SUM(A1:A11)')
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.program).toBe(compiled.program)
    expect(rewritten.compiled.constants).toBe(compiled.constants)
    expect(rewritten.compiled.deps).toEqual(['A1:A11'])
    expect(rewritten.compiled.symbolicRanges).toEqual(['A1:A11'])
    expect(rewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'A1', end: 'A11', refKind: 'cells' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: 'A1', end: 'A11', refKind: 'cells' }],
      },
      { opcode: 'return' },
    ])
  })

  it('recompiles structural rewrites when the formula shape changes', () => {
    const compiled = compileFormula('A6')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'delete',
      axis: 'row',
      start: 5,
      count: 1,
    })

    expect(rewritten.source).toBe('#REF!')
    expect(rewritten.reusedProgram).toBe(false)
    expect(rewritten.compiled.ast.kind).toBe('ErrorLiteral')
    expect(rewritten.compiled.optimizedAst.kind).toBe('ErrorLiteral')
  })

  it('rewrites compiled cell references in place for owner-sheet structural moves', () => {
    const compiled = compileFormula('A6')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 2,
      count: 1,
    })

    expect(rewritten.source).toBe('A7')
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps).toEqual(['A7'])
    expect(rewritten.compiled.symbolicRefs).toEqual(['A7'])
    expect(rewritten.compiled.parsedDeps).toEqual([
      {
        kind: 'cell',
        address: 'A7',
        row: 6,
        col: 0,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(rewritten.compiled.parsedSymbolicRefs).toEqual([
      {
        address: 'A7',
        row: 6,
        col: 0,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(rewritten.compiled.jsPlan).toEqual([{ opcode: 'push-cell', address: 'A7' }, { opcode: 'return' }])
  })

  it('rewrites compiled row and column range plans without recompiling', () => {
    const rowCompiled = compileFormula('SUM(5:7)')
    const rowRewritten = rewriteCompiledFormulaForStructuralTransform(rowCompiled, 'Sheet1', 'Sheet1', {
      kind: 'move',
      axis: 'row',
      start: 4,
      count: 2,
      target: 1,
    })
    expect(rowRewritten.source).toBe('SUM(2:7)')
    expect(rowRewritten.reusedProgram).toBe(true)
    expect(rowRewritten.compiled.deps).toEqual(['2:7'])
    expect(rowRewritten.compiled.symbolicRanges).toEqual(['2:7'])
    expect(rowRewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: '2', end: '7', refKind: 'rows' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: '2', end: '7', refKind: 'rows' }],
      },
      { opcode: 'return' },
    ])

    const colCompiled = compileFormula('SUM(C:E)')
    const colRewritten = rewriteCompiledFormulaForStructuralTransform(colCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'column',
      start: 1,
      count: 1,
    })
    expect(colRewritten.source).toBe('SUM(D:F)')
    expect(colRewritten.reusedProgram).toBe(true)
    expect(colRewritten.compiled.deps).toEqual(['D:F'])
    expect(colRewritten.compiled.symbolicRanges).toEqual(['D:F'])
    expect(colRewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'D', end: 'F', refKind: 'cols' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: 'D', end: 'F', refKind: 'cols' }],
      },
      { opcode: 'return' },
    ])
  })

  it('preserves explicit sheet qualification when rewriting compiled range references', () => {
    const compiled = compileFormula("SUM('Other Sheet'!A1:A3)")
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Other Sheet', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.source).toBe("SUM('Other Sheet'!A2:A4)")
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps).toHaveLength(1)
    expect(rewritten.compiled.deps[0]).toBe("'Other Sheet'!A2:A4")
    expect(rewritten.compiled.symbolicRanges).toHaveLength(1)
    expect(rewritten.compiled.symbolicRanges[0]).toBe("'Other Sheet'!A2:A4")
    expect(rewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', sheetName: 'Other Sheet', start: 'A2', end: 'A4', refKind: 'cells' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', sheetName: 'Other Sheet', start: 'A2', end: 'A4', refKind: 'cells' }],
      },
      { opcode: 'return' },
    ])
  })

  it('rewrites compiled row ranges for downward moves without recompiling', () => {
    const compiled = compileFormula('SUM(2:3)')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'move',
      axis: 'row',
      start: 1,
      count: 1,
      target: 4,
    })

    expect(rewritten.source).toBe('SUM(2:5)')
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps).toEqual(['2:5'])
    expect(rewritten.compiled.symbolicRanges).toEqual(['2:5'])
    expect(rewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: '2', end: '5', refKind: 'rows' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: '2', end: '5', refKind: 'rows' }],
      },
      { opcode: 'return' },
    ])
  })

  it('leaves compiled references unchanged when the structural target sheet does not match', () => {
    const compiled = compileFormula("SUM('Other Sheet'!A1:A3)")
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.source).toBe("SUM('Other Sheet'!A1:A3)")
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps[0]).toBe("'Other Sheet'!A1:A3")
    expect(rewritten.compiled.symbolicRanges[0]).toBe("'Other Sheet'!A1:A3")
    expect(rewritten.compiled.jsPlan).toEqual(compiled.jsPlan)
  })

  it('leaves compiled row and column ranges unchanged when the transform axis does not apply', () => {
    const rowCompiled = compileFormula('SUM(5:7)')
    const rowRewritten = rewriteCompiledFormulaForStructuralTransform(rowCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'column',
      start: 0,
      count: 1,
    })
    expect(rowRewritten.source).toBe('SUM(5:7)')
    expect(rowRewritten.reusedProgram).toBe(true)
    expect(rowRewritten.compiled.jsPlan).toEqual(rowCompiled.jsPlan)

    const colCompiled = compileFormula('SUM(C:E)')
    const colRewritten = rewriteCompiledFormulaForStructuralTransform(colCompiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })
    expect(colRewritten.source).toBe('SUM(C:E)')
    expect(colRewritten.reusedProgram).toBe(true)
    expect(colRewritten.compiled.jsPlan).toEqual(colCompiled.jsPlan)
  })

  it('rewrites compiled lookup plans across structural inserts', () => {
    const compiled = compileFormula('MATCH(A6,A1:A10,0)')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.source).toBe('MATCH(A7,A2:A11,0)')
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps).toEqual(['A7', 'A2:A11'])
    expect(rewritten.compiled.symbolicRefs).toEqual(['A7'])
    expect(rewritten.compiled.symbolicRanges).toEqual(['A2:A11'])
    expect(rewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-cell', address: 'A7' },
      {
        opcode: 'lookup-exact-match',
        callee: 'MATCH',
        start: 'A2',
        end: 'A11',
        startRow: 1,
        endRow: 10,
        startCol: 0,
        endCol: 0,
        refKind: 'cells',
        searchMode: 1,
      },
      { opcode: 'return' },
    ])
  })

  it('rewrites compiled call argument references for mixed cell and range operands', () => {
    const compiled = compileFormula('COUNTIF(A1:A3,B1)')
    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.source).toBe('COUNTIF(A2:A4,B2)')
    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.compiled.deps).toEqual(['A2:A4', 'B2'])
    expect(rewritten.compiled.symbolicRefs).toEqual(['B2'])
    expect(rewritten.compiled.symbolicRanges).toEqual(['A2:A4'])
    expect(rewritten.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'A2', end: 'A4', refKind: 'cells' },
      { opcode: 'push-cell', address: 'B2' },
      {
        opcode: 'call',
        callee: 'COUNTIF',
        argc: 2,
        argRefs: [
          { kind: 'range', start: 'A2', end: 'A4', refKind: 'cells' },
          { kind: 'cell', address: 'B2' },
        ],
      },
      { opcode: 'return' },
    ])
  })

  it('rewrites manual compiled invoke formulas and lambda plan bodies without recompiling', () => {
    const ast: FormulaNode = {
      kind: 'InvokeExpr',
      callee: { kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' },
      args: [
        { kind: 'ColumnRef', ref: 'C', sheetName: 'Sheet1' },
        { kind: 'RowRef', ref: '5', sheetName: 'Sheet1' },
      ],
    }
    const jsPlan: JsPlanInstruction[] = [
      {
        opcode: 'push-lambda',
        params: ['x'],
        body: [{ opcode: 'push-cell', sheetName: 'Sheet1', address: 'A1' }, { opcode: 'return' }],
      },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 4,
        argRefs: [{ kind: 'range' }, { kind: 'cell', sheetName: 'Other Sheet', address: 'B2' }, { kind: 'row' }, { kind: 'col' }],
      },
      { opcode: 'return' },
    ]
    const compiled = makeCompiledFormula(ast, { jsPlan })

    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.reusedProgram).toBe(true)
    expect(rewritten.source).toBe('(Sales[Amount])(Sheet1!C,Sheet1!6)')
    expect(rewritten.compiled.ast).toEqual({
      kind: 'InvokeExpr',
      callee: { kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' },
      args: [
        { kind: 'ColumnRef', ref: 'C', sheetName: 'Sheet1' },
        { kind: 'RowRef', ref: '6', sheetName: 'Sheet1' },
      ],
    })
    expect(rewritten.compiled.jsPlan).toEqual([
      {
        opcode: 'push-lambda',
        params: ['x'],
        body: [{ opcode: 'push-cell', sheetName: 'Sheet1', address: 'A2' }, { opcode: 'return' }],
      },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 4,
        argRefs: [{ kind: 'range' }, { kind: 'cell', sheetName: 'Other Sheet', address: 'B2' }, { kind: 'row' }, { kind: 'col' }],
      },
      { opcode: 'return' },
    ])
  })

  it('collapses malformed manual compiled refs to #REF! when structural rewrites cannot preserve them', () => {
    const compiled = makeCompiledFormula(
      {
        kind: 'CallExpr',
        callee: 'SUM',
        args: [
          { kind: 'CellRef', ref: 'A0', sheetName: 'Sheet1' },
          { kind: 'RowRef', ref: '0', sheetName: 'Sheet1' },
        ],
      },
      {
        jsPlan: [{ opcode: 'push-cell', sheetName: 'Sheet1', address: 'A0' }, { opcode: 'return' }],
      },
    )

    const rewritten = rewriteCompiledFormulaForStructuralTransform(compiled, 'Sheet1', 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(rewritten.reusedProgram).toBe(false)
    expect(rewritten.source).toBe('SUM(#REF!,#REF!)')
    expect(rewritten.compiled.ast).toEqual({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [
        { kind: 'ErrorLiteral', code: ErrorCode.Ref },
        { kind: 'ErrorLiteral', code: ErrorCode.Ref },
      ],
    })
  })

  it('translates compiled formulas by row and column offsets without recompiling', () => {
    const compiled = compileFormula('SUM(A1:B1)+C1')
    const translated = translateCompiledFormula(compiled, 2, 1)

    expect(translated.source).toBe('SUM(B3:C3)+D3')
    expect(translated.compiled.deps).toEqual(['B3:C3', 'D3'])
    expect(translated.compiled.symbolicRefs).toEqual(['D3'])
    expect(translated.compiled.symbolicRanges).toEqual(['B3:C3'])
    expect(translated.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'B3', end: 'C3', refKind: 'cells' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: 'B3', end: 'C3', refKind: 'cells' }],
      },
      { opcode: 'push-cell', address: 'D3' },
      { opcode: 'binary', operator: '+' },
      { opcode: 'return' },
    ])
  })
})
