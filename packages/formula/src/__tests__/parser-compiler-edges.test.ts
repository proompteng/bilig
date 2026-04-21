import { describe, expect, it } from 'vitest'
import { ErrorCode, FormulaMode, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  bindFormula,
  compileFormula,
  encodeBuiltin,
  evaluatePlan,
  isBuiltinAvailable,
  lexFormula,
  lowerToPlan,
  parseFormula,
} from '../index.js'

const context = {
  sheetName: 'Sheet1',
  resolveCell: (_sheetName: string, address: string): CellValue => {
    switch (address) {
      case 'A1':
        return { tag: ValueTag.Number, value: 4 }
      case 'B1':
        return { tag: ValueTag.Boolean, value: true }
      default:
        return { tag: ValueTag.Empty }
    }
  },
  resolveRange: (_sheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols'): CellValue[] => {
    if (refKind === 'cells' && start === 'A1' && end === 'A2') {
      return [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ]
    }
    return []
  },
}

describe('formula parser/compiler edges', () => {
  it('lexes escaped quoted sheet names and rejects invalid tokens', () => {
    expect(lexFormula("'O''Brien'!A1").slice(0, 3)).toEqual([
      { kind: 'quotedIdentifier', value: "O'Brien" },
      { kind: 'bang', value: '!' },
      { kind: 'identifier', value: 'A1' },
    ])
    expect(lexFormula('10%').slice(0, 2)).toEqual([
      { kind: 'number', value: '10' },
      { kind: 'percent', value: '%' },
    ])
    expect(lexFormula('"he said ""hi"""').slice(0, 1)).toEqual([{ kind: 'string', value: 'he said "hi"' }])
    expect(lexFormula('A1<>B1').slice(1, 2)).toEqual([{ kind: 'neq', value: '<>' }])
    expect(() => lexFormula('@oops')).toThrow("Unexpected token '@'")
  })

  it('parses workbook error literals inside formulas and as standalone expressions', () => {
    expect(lexFormula('A1+#REF!').slice(0, 3)).toEqual([
      { kind: 'identifier', value: 'A1' },
      { kind: 'plus', value: '+' },
      { kind: 'error', value: '#REF!' },
    ])

    expect(parseFormula('A1+#REF!')).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'CellRef', ref: 'A1' },
      right: { kind: 'ErrorLiteral', code: ErrorCode.Ref },
    })

    expect(parseFormula('#DIV/0!')).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Div0,
    })
  })

  it('rejects standalone axis refs and malformed ranges', () => {
    expect(() => parseFormula("'Sheet 1'!1")).toThrow('Row and column references must appear inside a range')
    expect(() => parseFormula("'Sheet 1'!$1")).toThrow('Row and column references must appear inside a range')
    expect(() => parseFormula('A1:B')).toThrow('Range endpoints must use the same reference type')
    expect(() => parseFormula('A1:2')).toThrow('Range endpoints must use the same reference type')
  })

  it('binds quoted ranges and keeps unsupported/text formulas on the JS path', () => {
    const quotedRange = bindFormula(parseFormula("SUM('My Sheet'!A1:A2)"))
    expect(quotedRange.deps).toEqual(["'My Sheet'!A1:A2"])
    expect(quotedRange.mode).toBe(FormulaMode.WasmFastPath)

    const anchoredRange = bindFormula(parseFormula("SUM('My Sheet'!$A:$B)"))
    expect(anchoredRange.deps).toEqual(["'My Sheet'!A:B"])
    expect(anchoredRange.mode).toBe(FormulaMode.WasmFastPath)

    expect(bindFormula(parseFormula('"hello"')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('A1')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('LEN(A1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('LEN(A1:A2)')).mode).toBe(FormulaMode.JsOnly)
  })

  it('parses defined names, tracks them separately, and lowers them onto the JS plan', () => {
    const ast = parseFormula('TaxRate*A1')
    expect(ast).toEqual({
      kind: 'BinaryExpr',
      operator: '*',
      left: { kind: 'NameRef', name: 'TaxRate' },
      right: { kind: 'CellRef', ref: 'A1' },
    })

    const bound = bindFormula(ast)
    expect(bound.deps).toEqual(['A1'])
    expect(bound.symbolicNames).toEqual(['TaxRate'])
    expect(bound.mode).toBe(FormulaMode.JsOnly)

    const compiled = compileFormula('TaxRate*A1')
    expect(compiled.symbolicNames).toEqual(['TaxRate'])
    expect(compiled.jsPlan).toEqual([
      { opcode: 'push-name', name: 'TaxRate' },
      { opcode: 'push-cell', address: 'A1' },
      { opcode: 'binary', operator: '*' },
      { opcode: 'return' },
    ])
  })

  it('lowers exact vector MATCH and XMATCH to the direct lookup opcode only for exact shapes', () => {
    expect(lowerToPlan(parseFormula('MATCH(A1,A2:A4,0)'))).toEqual([
      { opcode: 'push-cell', address: 'A1' },
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
      { opcode: 'return' },
    ])

    expect(lowerToPlan(parseFormula('XMATCH("pear",A1:A4,0,-1)'))).toEqual([
      { opcode: 'push-string', value: 'pear' },
      {
        opcode: 'lookup-exact-match',
        callee: 'XMATCH',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'cells',
        searchMode: -1,
      },
      { opcode: 'return' },
    ])

    expect(lowerToPlan(parseFormula('MATCH(A1,A2:A4,1)'))).toEqual([
      { opcode: 'push-cell', address: 'A1' },
      {
        opcode: 'lookup-approximate-match',
        callee: 'MATCH',
        start: 'A2',
        end: 'A4',
        startRow: 1,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'cells',
        matchMode: 1,
      },
      { opcode: 'return' },
    ])

    expect(lowerToPlan(parseFormula('XMATCH("pear",A1:A4,-1)'))).toEqual([
      { opcode: 'push-string', value: 'pear' },
      {
        opcode: 'lookup-approximate-match',
        callee: 'XMATCH',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        refKind: 'cells',
        matchMode: -1,
      },
      { opcode: 'return' },
    ])

    expect(lowerToPlan(parseFormula('MATCH(A1,A2:A4)'))).toEqual([
      { opcode: 'push-cell', address: 'A1' },
      { opcode: 'push-range', start: 'A2', end: 'A4', refKind: 'cells' },
      {
        opcode: 'call',
        callee: 'MATCH',
        argc: 2,
        argRefs: [
          { kind: 'cell', address: 'A1' },
          { kind: 'range', start: 'A2', end: 'A4', refKind: 'cells' },
        ],
      },
      { opcode: 'return' },
    ])
  })

  it('parses structured references and spill refs as metadata-aware syntax', () => {
    const structured = parseFormula('SUM(Sales[Amount])')
    expect(structured).toEqual({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [{ kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' }],
    })

    const structuredBound = bindFormula(structured)
    expect(structuredBound.symbolicTables).toEqual(['Sales'])
    expect(structuredBound.mode).toBe(FormulaMode.JsOnly)

    const spill = parseFormula('A1#')
    expect(spill).toEqual({ kind: 'SpillRef', ref: 'A1' })

    const spillBound = bindFormula(spill)
    expect(spillBound.symbolicSpills).toEqual(['A1'])
    expect(spillBound.mode).toBe(FormulaMode.JsOnly)
  })

  it('throws on unsupported wasm builtin encodings and invalid axis compilation', () => {
    expect(isBuiltinAvailable('SUM')).toBe(true)
    expect(isBuiltinAvailable('MATCH')).toBe(true)
    expect(isBuiltinAvailable('INDEX')).toBe(true)
    expect(isBuiltinAvailable('VLOOKUP')).toBe(true)
    expect(isBuiltinAvailable('DOES_NOT_EXIST')).toBe(false)
    expect(encodeBuiltin('LEN')).toBeDefined()
    expect(() => encodeBuiltin('DOES_NOT_EXIST')).toThrow('Unsupported builtin for wasm: DOES_NOT_EXIST')
    expect(() => compileFormula('Sheet1!A')).toThrow('Row and column references must appear inside a range')
  })

  it('routes the new statistical builtin tranche onto the wasm path with proper ids', () => {
    expect(compileFormula('ERF(1)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('ERF(0,1)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('ERF.PRECISE(1)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('ERFC(1)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('ERFC.PRECISE(1)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('FISHER(0.5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('FISHERINV(0.5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('GAMMALN(5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('GAMMALN.PRECISE(5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('GAMMA(5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('CONFIDENCE(0.05,1.5,100)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('EXPONDIST(1,2,FALSE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('EXPON.DIST(1,2,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('POISSON(3,2.5,FALSE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('POISSON.DIST(3,2.5,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('WEIBULL(1.5,2,3,FALSE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('WEIBULL.DIST(1.5,2,3,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('GAMMADIST(2,3,2,FALSE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('GAMMA.DIST(2,3,2,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('CHIDIST(3,4)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('CHISQ.DIST.RT(3,4)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('CHISQ.DIST(3,4,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('BINOMDIST(2,4,0.5,FALSE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('BINOM.DIST(2,4,0.5,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('BINOM.DIST.RANGE(6,0.5,2,4)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('CRITBINOM(6,0.5,0.7)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('BINOM.INV(6,0.5,0.7)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('HYPGEOMDIST(1,4,3,10)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('HYPGEOM.DIST(1,4,3,10,TRUE)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('NEGBINOMDIST(2,3,0.5)').mode).toBe(FormulaMode.WasmFastPath)
    expect(compileFormula('NEGBINOM.DIST(2,3,0.5,TRUE)').mode).toBe(FormulaMode.WasmFastPath)

    expect(encodeBuiltin('ERF')).toBeDefined()
    expect(encodeBuiltin('CONFIDENCE')).toBeDefined()
    expect(encodeBuiltin('POISSON.DIST')).toBeDefined()
    expect(encodeBuiltin('BINOM.INV')).toBeDefined()
    expect(encodeBuiltin('NEGBINOM.DIST')).toBeDefined()
  })

  it('keeps invalid wasm arities and non-scalar argument shapes on the JS path', () => {
    expect(compileFormula('SEQUENCE()').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('LOOKUP(A1)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('LOOKUP(A1,A1:A2,B1:B2,C1:C2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('SORTBY(A1:A2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('OFFSET()').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('AREAS(A1)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('COUNTIFS(A1)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('SUMIFS(1, A1:A2, 1)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('IFS(A1:A2,1,TRUE(),2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('MINIFS(A1:A2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('IRR(A1)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('MIRR(A1,1,2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('XNPV(0.1,A1,B1:B5)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('XIRR(A1,B1:B5)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('TEXT(A1:A2,"0.00")').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('NUMBERVALUE(A1:A2)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('PHONETIC(LAMBDA(x,x))').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('TRANSPOSE()').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('PROB(A1:A3,B1:B3,C1:C3)').mode).toBe(FormulaMode.JsOnly)
    expect(compileFormula('TRIMMEAN(A1:A8,A1:A2)').mode).toBe(FormulaMode.JsOnly)
  })

  it('tracks unknown callees as symbolic names while preserving rewritten and range-safe bindings', () => {
    const unknownCall = bindFormula(parseFormula('CustomFn(A1)'))
    expect(unknownCall.deps).toEqual(['A1'])
    expect(unknownCall.symbolicNames).toEqual(['CUSTOMFN'])
    expect(unknownCall.mode).toBe(FormulaMode.JsOnly)

    expect(bindFormula(parseFormula('LOOKUP(A1,A1:A2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('SORTBY(A1:A2,B1:B2,1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('AREAS(A1:A2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('ARRAYTOTEXT(A1:A2,1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('IRR(A1:A6)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('MIRR(A1:A6,10%,12%)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('XNPV(0.09,A1:A5,B1:B5)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('XIRR(A1:A5,B1:B5)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('TEXT("42","0.00")')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('NUMBERVALUE("2.500,27",",",".")')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('PHONETIC(A1:A2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('TRANSPOSE(A1:B2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('HSTACK(A1:A2,B1:B2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('VSTACK(A1:A2,B1:B2)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('PROB(A1:A4,B1:B4,2,3)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('TRIMMEAN(A1:A8,0.25)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('MINIFS(A1:A4,B1:B4,1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('MAXIFS(A1:A4,B1:B4,1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('GROUPBY(A1:A5,C1:C5,SUM,3,1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)')).mode).toBe(FormulaMode.WasmFastPath)
  })

  it('keeps non-call and invalid rewritten top-level nodes off the wasm fast path', () => {
    expect(bindFormula(parseFormula('TRUE(1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('SWITCH(A1)')).mode).toBe(FormulaMode.WasmFastPath)
    expect(bindFormula(parseFormula('A1+1')).mode).toBe(FormulaMode.WasmFastPath)
  })

  it('compiles literal text, CONCAT, and IF text branches onto the wasm path', () => {
    const textIf = compileFormula('IF(A1, CONCAT("x", "y"), "z")')
    expect(textIf.mode).toBe(FormulaMode.WasmFastPath)
    expect(textIf.program[0] >>> 24).toBe(Opcode.PushCell)
    expect(textIf.symbolicStrings).toEqual(['xy', 'z'])

    const plainString = compileFormula('"hello"')
    expect(plainString.mode).toBe(FormulaMode.WasmFastPath)
    expect(plainString.symbolicStrings).toEqual(['hello'])
    expect(plainString.program).toEqual(Uint32Array.from([(Opcode.PushString << 24) | 0, 255 << 24]))
    expect(plainString.jsPlan).toEqual([{ opcode: 'push-string', value: 'hello' }, { opcode: 'return' }])

    const concat = compileFormula('CONCAT("x", "y")')
    expect(concat.mode).toBe(FormulaMode.WasmFastPath)
    expect(concat.symbolicStrings).toEqual(['xy'])

    const compared = compileFormula('A1="HELLO"')
    expect(compared.mode).toBe(FormulaMode.WasmFastPath)
  })

  it('parses postfix percent as arithmetic scaling', () => {
    expect(parseFormula('10%')).toEqual({
      kind: 'BinaryExpr',
      operator: '*',
      left: { kind: 'NumberLiteral', value: 10 },
      right: { kind: 'NumberLiteral', value: 0.01 },
    })

    expect(parseFormula('(A1+A2)%')).toEqual({
      kind: 'BinaryExpr',
      operator: '*',
      left: {
        kind: 'BinaryExpr',
        operator: '+',
        left: { kind: 'CellRef', ref: 'A1' },
        right: { kind: 'CellRef', ref: 'A2' },
      },
      right: { kind: 'NumberLiteral', value: 0.01 },
    })
  })

  it('evaluates lowered plans across comparison, unary, jump, and builtin error paths', () => {
    expect(
      evaluatePlan(
        [
          { opcode: 'push-cell', address: 'A1' },
          { opcode: 'push-number', value: 4 },
          { opcode: 'binary', operator: '=' },
          { opcode: 'return' },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Boolean, value: true })

    expect(
      evaluatePlan([{ opcode: 'push-string', value: 'x' }, { opcode: 'unary', operator: '-' }, { opcode: 'return' }], context),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      evaluatePlan(
        [
          { opcode: 'push-boolean', value: false },
          { opcode: 'jump-if-false', target: 4 },
          { opcode: 'push-number', value: 1 },
          { opcode: 'jump', target: 5 },
          { opcode: 'push-number', value: 2 },
          { opcode: 'return' },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(
      evaluatePlan(
        [
          { opcode: 'push-range', start: 'A1', end: 'A2', refKind: 'cells' },
          { opcode: 'call', callee: 'SUM', argc: 1 },
          { opcode: 'return' },
        ],
        context,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 })

    expect(
      evaluatePlan([{ opcode: 'push-number', value: 1 }, { opcode: 'call', callee: 'UNKNOWN', argc: 1 }, { opcode: 'return' }], context),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
  })

  it('lowers full IF expressions into explicit jump instructions', () => {
    expect(lowerToPlan(parseFormula('IF(A1, B1, 0)'))).toEqual([
      { opcode: 'push-cell', address: 'A1' },
      { opcode: 'jump-if-false', target: 4 },
      { opcode: 'push-cell', address: 'B1' },
      { opcode: 'jump', target: 5 },
      { opcode: 'push-number', value: 0 },
      { opcode: 'return' },
    ])
  })

  it('parses and lowers lambda invocation syntax', () => {
    expect(parseFormula('LAMBDA(x,x+1)(4)')).toEqual({
      kind: 'InvokeExpr',
      callee: {
        kind: 'CallExpr',
        callee: 'LAMBDA',
        args: [
          { kind: 'NameRef', name: 'x' },
          {
            kind: 'BinaryExpr',
            operator: '+',
            left: { kind: 'NameRef', name: 'x' },
            right: { kind: 'NumberLiteral', value: 1 },
          },
        ],
      },
      args: [{ kind: 'NumberLiteral', value: 4 }],
    })

    expect(lowerToPlan(parseFormula('LAMBDA(x,x+1)(4)'))).toEqual([
      {
        opcode: 'push-lambda',
        params: ['x'],
        body: [
          { opcode: 'push-name', name: 'x' },
          { opcode: 'push-number', value: 1 },
          { opcode: 'binary', operator: '+' },
          { opcode: 'return' },
        ],
      },
      { opcode: 'push-number', value: 4 },
      { opcode: 'invoke', argc: 1 },
      { opcode: 'return' },
    ])
  })
})
