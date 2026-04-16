import { describe, expect, it } from 'vitest'
import { FormulaMode } from '@bilig/protocol'
import { compileFormula, type CompiledFormula } from '../compiler.js'
import { canTranslateCompiledFormulaWithoutAst, translateCompiledFormulaWithoutAst } from '../translation.js'

describe('translateCompiledFormulaWithoutAst', () => {
  it('reports when a compiled formula can use the AST-free translation fast path', () => {
    const simple = compileFormula('A1+B1')
    const jsOnly = compileFormula('FORMULATEXT(A1)')
    const directAggregate = compileFormula('SUM(A1:A3)')
    const rangeWithoutAggregate = compileFormula('COUNTIF(A1:A3,B1)')
    const lookup = compileFormula('MATCH(A6,A1:A10,0)')

    expect(canTranslateCompiledFormulaWithoutAst(simple)).toBe(true)
    expect(canTranslateCompiledFormulaWithoutAst(jsOnly)).toBe(true)
    expect(canTranslateCompiledFormulaWithoutAst(directAggregate)).toBe(false)
    expect(canTranslateCompiledFormulaWithoutAst(rangeWithoutAggregate)).toBe(false)
    expect(canTranslateCompiledFormulaWithoutAst(lookup)).toBe(false)
  })

  it('translates parsed refs and deps while reusing the wasm js plan for cell-only formulas', () => {
    const compiled = compileFormula('A1+B1')

    const translated = translateCompiledFormulaWithoutAst(compiled, 2, 1, 'B3+C3')

    expect(translated.source).toBe('B3+C3')
    expect(translated.compiled.ast).toBe(compiled.ast)
    expect(translated.compiled.optimizedAst).toBe(compiled.optimizedAst)
    expect(translated.compiled.astMatchesSource).toBe(false)
    expect(translated.compiled.deps).toEqual(['B3', 'C3'])
    expect(translated.compiled.symbolicRefs).toEqual(['B3', 'C3'])
    expect(translated.compiled.parsedDeps).toEqual([
      {
        kind: 'cell',
        address: 'B3',
        row: 2,
        col: 1,
        rowAbsolute: false,
        colAbsolute: false,
      },
      {
        kind: 'cell',
        address: 'C3',
        row: 2,
        col: 2,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(translated.compiled.parsedSymbolicRefs).toEqual([
      {
        address: 'B3',
        row: 2,
        col: 1,
        rowAbsolute: false,
        colAbsolute: false,
      },
      {
        address: 'C3',
        row: 2,
        col: 2,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(translated.compiled.jsPlan).toBe(compiled.jsPlan)
  })

  it('rewrites js plan operands for js-only formulas without relying on the AST', () => {
    const compiled = compileFormula('CELL("row",A1)')

    const translated = translateCompiledFormulaWithoutAst(compiled, 1, 2, 'CELL("row",C2)')

    expect(translated.source).toBe('CELL("row",C2)')
    expect(translated.compiled.astMatchesSource).toBe(false)
    expect(translated.compiled.deps).toEqual(['C2'])
    expect(translated.compiled.parsedDeps).toEqual([
      {
        kind: 'cell',
        address: 'C2',
        row: 1,
        col: 2,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(translated.compiled.jsPlan).toEqual([
      { opcode: 'push-string', value: 'row' },
      { opcode: 'push-cell', address: 'C2' },
      {
        opcode: 'call',
        callee: 'CELL',
        argc: 2,
        argRefs: [undefined, { kind: 'cell', address: 'C2' }],
      },
      { opcode: 'return' },
    ])
  })

  it('rewrites range metadata and js plan operands for direct aggregate formulas', () => {
    const compiled = compileFormula('SUM(A1:A3)')

    const translated = translateCompiledFormulaWithoutAst(compiled, 2, 0, 'SUM(A3:A5)')

    expect(compiled.mode).toBe(FormulaMode.WasmFastPath)
    expect(translated.compiled.mode).toBe(FormulaMode.WasmFastPath)
    expect(translated.source).toBe('SUM(A3:A5)')
    expect(translated.compiled.deps).toEqual(['A3:A5'])
    expect(translated.compiled.symbolicRanges).toEqual(['A3:A5'])
    expect(translated.compiled.parsedDeps).toEqual([
      {
        address: 'A3:A5',
        kind: 'range',
        refKind: 'cells',
        startAddress: 'A3',
        endAddress: 'A5',
        startRow: 2,
        endRow: 4,
        startCol: 0,
        endCol: 0,
        startRowAbsolute: false,
        endRowAbsolute: false,
        startColAbsolute: false,
        endColAbsolute: false,
      },
    ])
    expect(translated.compiled.parsedSymbolicRanges).toEqual([
      {
        address: 'A3:A5',
        kind: 'range',
        refKind: 'cells',
        startAddress: 'A3',
        endAddress: 'A5',
        startRow: 2,
        endRow: 4,
        startCol: 0,
        endCol: 0,
        startRowAbsolute: false,
        endRowAbsolute: false,
        startColAbsolute: false,
        endColAbsolute: false,
      },
    ])
    expect(translated.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'A3', end: 'A5', refKind: 'cells' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: 'A3', end: 'A5', refKind: 'cells' }],
      },
      { opcode: 'return' },
    ])
  })

  it('translates axis ranges and call argument refs when parsed metadata is present', () => {
    const rows = translateCompiledFormulaWithoutAst(compileFormula('SUM(5:7)'), 2, 0, 'SUM(7:9)')
    expect(rows.compiled.deps).toEqual(['7:9'])
    expect(rows.compiled.symbolicRanges).toEqual(['7:9'])
    expect(rows.compiled.parsedDeps).toEqual([
      {
        address: '7:9',
        kind: 'range',
        refKind: 'rows',
        startAddress: '7',
        endAddress: '9',
        startRow: 6,
        endRow: 8,
        startCol: 0,
        endCol: 0,
        startRowAbsolute: false,
        endRowAbsolute: false,
      },
    ])
    expect(rows.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: '7', end: '9', refKind: 'rows' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: '7', end: '9', refKind: 'rows' }],
      },
      { opcode: 'return' },
    ])

    const cols = translateCompiledFormulaWithoutAst(compileFormula('SUM(C:E)'), 0, 2, 'SUM(E:G)')
    expect(cols.compiled.deps).toEqual(['E:G'])
    expect(cols.compiled.symbolicRanges).toEqual(['E:G'])
    expect(cols.compiled.parsedDeps).toEqual([
      {
        address: 'E:G',
        kind: 'range',
        refKind: 'cols',
        startAddress: 'E',
        endAddress: 'G',
        startRow: 0,
        endRow: 0,
        startCol: 4,
        endCol: 6,
        startColAbsolute: false,
        endColAbsolute: false,
      },
    ])
    expect(cols.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'E', end: 'G', refKind: 'cols' },
      {
        opcode: 'call',
        callee: 'SUM',
        argc: 1,
        argRefs: [{ kind: 'range', start: 'E', end: 'G', refKind: 'cols' }],
      },
      { opcode: 'return' },
    ])

    const countIf = translateCompiledFormulaWithoutAst(compileFormula('COUNTIF(A1:A3,B1)'), 1, 2, 'COUNTIF(C2:C4,D2)')
    expect(countIf.compiled.deps).toEqual(['C2:C4', 'D2'])
    expect(countIf.compiled.symbolicRanges).toEqual(['C2:C4'])
    expect(countIf.compiled.symbolicRefs).toEqual(['D2'])
    expect(countIf.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'C2', end: 'C4', refKind: 'cells' },
      { opcode: 'push-cell', address: 'D2' },
      {
        opcode: 'call',
        callee: 'COUNTIF',
        argc: 2,
        argRefs: [
          { kind: 'range', start: 'C2', end: 'C4', refKind: 'cells' },
          { kind: 'cell', address: 'D2' },
        ],
      },
      { opcode: 'return' },
    ])
  })

  it('falls back to generic plan translation when parsed metadata is unavailable', () => {
    const compiled = compileFormula('COUNTIF(A1:A3,B1)')
    const withoutParsedMetadata: CompiledFormula = {
      ...compiled,
      parsedDeps: undefined,
      parsedSymbolicRefs: undefined,
      parsedSymbolicRanges: undefined,
    }

    const translated = translateCompiledFormulaWithoutAst(withoutParsedMetadata, 2, 1, 'COUNTIF(B3:B5,C3)')

    expect(translated.compiled.deps).toEqual(['B3:B5', 'C3'])
    expect(translated.compiled.symbolicRanges).toEqual(['B3:B5'])
    expect(translated.compiled.symbolicRefs).toEqual(['C3'])
    expect(translated.compiled.parsedDeps).toBeUndefined()
    expect(translated.compiled.parsedSymbolicRefs).toBeUndefined()
    expect(translated.compiled.parsedSymbolicRanges).toBeUndefined()
    expect(translated.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'B3', end: 'B5', refKind: 'cells' },
      { opcode: 'push-cell', address: 'C3' },
      {
        opcode: 'call',
        callee: 'COUNTIF',
        argc: 2,
        argRefs: [
          { kind: 'range', start: 'B3', end: 'B5', refKind: 'cells' },
          { kind: 'cell', address: 'C3' },
        ],
      },
      { opcode: 'return' },
    ])
  })

  it('translates lookup instructions and nested lambda bodies without an AST', () => {
    const lookup = translateCompiledFormulaWithoutAst(compileFormula('MATCH(A6,A1:A10,0)'), 1, 0, 'MATCH(A7,A2:A11,0)')
    expect(lookup.compiled.deps).toEqual(['A7', 'A2:A11'])
    expect(lookup.compiled.jsPlan).toEqual([
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

    const manualCompiled: CompiledFormula = {
      ...compileFormula('A1'),
      source: 'WRAP(A1:B2,A1,2,C,LAMBDA(x,A1))',
      deps: ['A1:B2', 'A1'],
      parsedDeps: [
        {
          address: 'A1:B2',
          kind: 'range',
          refKind: 'cells',
          startAddress: 'A1',
          endAddress: 'B2',
          startRow: 0,
          endRow: 1,
          startCol: 0,
          endCol: 1,
          startRowAbsolute: false,
          endRowAbsolute: false,
          startColAbsolute: false,
          endColAbsolute: false,
        },
        {
          kind: 'cell',
          address: 'A1',
          row: 0,
          col: 0,
          rowAbsolute: false,
          colAbsolute: false,
        },
      ],
      symbolicRefs: ['A1'],
      parsedSymbolicRefs: [
        {
          address: 'A1',
          row: 0,
          col: 0,
          rowAbsolute: false,
          colAbsolute: false,
        },
      ],
      symbolicRanges: ['A1:B2'],
      parsedSymbolicRanges: [
        {
          address: 'A1:B2',
          kind: 'range',
          refKind: 'cells',
          startAddress: 'A1',
          endAddress: 'B2',
          startRow: 0,
          endRow: 1,
          startCol: 0,
          endCol: 1,
          startRowAbsolute: false,
          endRowAbsolute: false,
          startColAbsolute: false,
          endColAbsolute: false,
        },
      ],
      jsPlan: [
        { opcode: 'push-range', start: 'A1', end: 'B2', refKind: 'cells' },
        { opcode: 'push-cell', address: 'A1' },
        {
          opcode: 'call',
          callee: 'WRAP',
          argc: 4,
          argRefs: [
            { kind: 'range', start: 'A1', end: 'B2', refKind: 'cells' },
            { kind: 'cell', address: 'A1' },
            { kind: 'row', address: '2' },
            { kind: 'col', address: 'C' },
          ],
        },
        {
          opcode: 'push-lambda',
          params: ['x'],
          body: [{ opcode: 'push-cell', address: 'A1' }, { opcode: 'return' }],
        },
        { opcode: 'return' },
      ],
    }

    const translatedManual = translateCompiledFormulaWithoutAst(manualCompiled, 1, 2, 'WRAP(C2:D3,C2,3,E,LAMBDA(x,C2))')
    expect(translatedManual.compiled.deps).toEqual(['C2:D3', 'C2'])
    expect(translatedManual.compiled.symbolicRefs).toEqual(['C2'])
    expect(translatedManual.compiled.symbolicRanges).toEqual(['C2:D3'])
    expect(translatedManual.compiled.jsPlan).toEqual([
      { opcode: 'push-range', start: 'C2', end: 'D3', refKind: 'cells' },
      { opcode: 'push-cell', address: 'C2' },
      {
        opcode: 'call',
        callee: 'WRAP',
        argc: 4,
        argRefs: [
          { kind: 'range', start: 'C2', end: 'D3', refKind: 'cells' },
          { kind: 'cell', address: 'C2' },
          { kind: 'row', address: '3' },
          { kind: 'col', address: 'E' },
        ],
      },
      {
        opcode: 'push-lambda',
        params: ['x'],
        body: [{ opcode: 'push-cell', address: 'C2' }, { opcode: 'return' }],
      },
      { opcode: 'return' },
    ])
  })
})
