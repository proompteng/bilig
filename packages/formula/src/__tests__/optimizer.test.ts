import { ErrorCode } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import type { FormulaNode } from '../ast.js'
import { optimizeFormula } from '../optimizer.js'
import { parseFormula } from '../parser.js'

describe('optimizer', () => {
  it('keeps volatile and non-foldable array calls intact', () => {
    expect(optimizeFormula(parseFormula('TODAY()'))).toEqual({
      kind: 'CallExpr',
      callee: 'TODAY',
      args: [],
    })
    expect(optimizeFormula(parseFormula('SEQUENCE(2)'))).toEqual({
      kind: 'CallExpr',
      callee: 'SEQUENCE',
      args: [{ kind: 'NumberLiteral', value: 2 }],
    })
  })

  it('folds IF, unary operators, and special call rewrites into ordinary AST', () => {
    expect(optimizeFormula(parseFormula('IF("",1,2)'))).toEqual({
      kind: 'NumberLiteral',
      value: 2,
    })
    expect(optimizeFormula(parseFormula('IF(TRUE(),1,2)'))).toEqual({
      kind: 'NumberLiteral',
      value: 1,
    })
    expect(optimizeFormula(parseFormula('+A1'))).toEqual({
      kind: 'CellRef',
      ref: 'A1',
    })
    expect(optimizeFormula(parseFormula('-2'))).toEqual({
      kind: 'NumberLiteral',
      value: -2,
    })
    expect(optimizeFormula(parseFormula('TRUE()'))).toEqual({
      kind: 'BooleanLiteral',
      value: true,
    })
    expect(optimizeFormula(parseFormula('FALSE()'))).toEqual({
      kind: 'BooleanLiteral',
      value: false,
    })
    expect(optimizeFormula(parseFormula('IFS(FALSE,1,TRUE,2)'))).toEqual({
      kind: 'NumberLiteral',
      value: 2,
    })
    expect(optimizeFormula(parseFormula('SWITCH("b","a",1,"b",2,9)'))).toEqual({
      kind: 'NumberLiteral',
      value: 2,
    })
    expect(optimizeFormula(parseFormula('XOR(TRUE(),FALSE(),TRUE())'))).toEqual({
      kind: 'BooleanLiteral',
      value: false,
    })
  })

  it('rewrites LET and direct lambda invocations while preserving shadowing', () => {
    expect(optimizeFormula(parseFormula('LET(x,1,LET(x,2,x+3)+x)'))).toEqual({
      kind: 'NumberLiteral',
      value: 6,
    })
    expect(optimizeFormula(parseFormula('LET(fn,LAMBDA(x,x+1),fn(A1))'))).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'CellRef', ref: 'A1' },
      right: { kind: 'NumberLiteral', value: 1 },
    })
    expect(optimizeFormula(parseFormula('LET(x,10,LAMBDA(x,x+1)(4)+x)'))).toEqual({
      kind: 'NumberLiteral',
      value: 15,
    })
    expect(optimizeFormula(parseFormula('LAMBDA(y,LET(x,2,x+y))(A1)'))).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'NumberLiteral', value: 2 },
      right: { kind: 'CellRef', ref: 'A1' },
    })
    expect(optimizeFormula(parseFormula('LET(x,2,LAMBDA(y,x+y)(A1))'))).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'NumberLiteral', value: 2 },
      right: { kind: 'CellRef', ref: 'A1' },
    })
    expect(optimizeFormula(parseFormula('LET(x,1,(fn)(x))'))).toEqual({
      kind: 'InvokeExpr',
      callee: { kind: 'NameRef', name: 'fn' },
      args: [{ kind: 'NumberLiteral', value: 1 }],
    })
    expect(optimizeFormula(parseFormula('LAMBDA(x,(fn)(x))(1)'))).toEqual({
      kind: 'InvokeExpr',
      callee: { kind: 'NameRef', name: 'fn' },
      args: [{ kind: 'NumberLiteral', value: 1 }],
    })
    expect(optimizeFormula(parseFormula('LET(x,-A1,x)'))).toEqual({
      kind: 'UnaryExpr',
      operator: '-',
      argument: { kind: 'CellRef', ref: 'A1' },
    })
    expect(optimizeFormula(parseFormula('LET(x,(fn)(1),x)'))).toEqual({
      kind: 'InvokeExpr',
      callee: { kind: 'NameRef', name: 'fn' },
      args: [{ kind: 'NumberLiteral', value: 1 }],
    })
    expect(optimizeFormula(parseFormula('MAP(A1:A3,LAMBDA(x,x*2))'))).toEqual({
      kind: 'BinaryExpr',
      operator: '*',
      left: { kind: 'RangeRef', refKind: 'cells', start: 'A1', end: 'A3' },
      right: { kind: 'NumberLiteral', value: 2 },
    })
    expect(optimizeFormula(parseFormula('MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))'))).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'RangeRef', refKind: 'cells', start: 'A1', end: 'A3' },
      right: { kind: 'RangeRef', refKind: 'cells', start: 'B1', end: 'B3' },
    })
  })

  it('returns value errors for invalid LET and lambda invocation rewrites', () => {
    expect(optimizeFormula(parseFormula('LET(x,1)'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })
    expect(optimizeFormula(parseFormula('LET(1,2,3)'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })
    expect(optimizeFormula(parseFormula('LAMBDA(x,x+1)(4,5)'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })
    expect(optimizeFormula(parseFormula('TRUE(1)'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })
    expect(optimizeFormula(parseFormula('MAP(A1:A3,LAMBDA(x,y,x+y))'))).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })

    const invalidLambdaParam: FormulaNode = {
      kind: 'InvokeExpr',
      callee: {
        kind: 'CallExpr',
        callee: 'LAMBDA',
        args: [
          { kind: 'CellRef', ref: 'A1' },
          { kind: 'NumberLiteral', value: 2 },
        ],
      },
      args: [{ kind: 'NumberLiteral', value: 3 }],
    }
    expect(optimizeFormula(invalidLambdaParam)).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })

    expect(
      optimizeFormula({
        kind: 'CallExpr',
        callee: 'MAP',
        args: [
          { kind: 'RangeRef', refKind: 'cells', start: 'A1', end: 'A3' },
          {
            kind: 'CallExpr',
            callee: 'LAMBDA',
            args: [
              { kind: 'CellRef', ref: 'A1' },
              { kind: 'NumberLiteral', value: 1 },
            ],
          },
        ],
      }),
    ).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })
  })

  it('preserves non-lambda invocation nodes and metadata-aware refs', () => {
    const invokeNameRef: FormulaNode = {
      kind: 'InvokeExpr',
      callee: { kind: 'NameRef', name: 'Fn' },
      args: [{ kind: 'NumberLiteral', value: 3 }],
    }
    expect(optimizeFormula(invokeNameRef)).toEqual(invokeNameRef)

    const invokeCallExpr: FormulaNode = {
      kind: 'InvokeExpr',
      callee: {
        kind: 'CallExpr',
        callee: 'SUM',
        args: [{ kind: 'NumberLiteral', value: 1 }],
      },
      args: [{ kind: 'NumberLiteral', value: 2 }],
    }
    expect(optimizeFormula(invokeCallExpr)).toEqual({
      kind: 'ErrorLiteral',
      code: ErrorCode.Value,
    })

    expect(optimizeFormula(parseFormula('TaxRate'))).toEqual({ kind: 'NameRef', name: 'TaxRate' })
    expect(optimizeFormula(parseFormula('A1#'))).toEqual({ kind: 'SpillRef', ref: 'A1' })
    expect(optimizeFormula(parseFormula('SUM(Sales[Amount])'))).toEqual({
      kind: 'CallExpr',
      callee: 'SUM',
      args: [{ kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' }],
    })

    const metadataAwareLet: FormulaNode = {
      kind: 'CallExpr',
      callee: 'LET',
      args: [
        { kind: 'NameRef', name: 'x' },
        { kind: 'NumberLiteral', value: 1 },
        {
          kind: 'CallExpr',
          callee: 'CONCAT',
          args: [
            { kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' },
            { kind: 'SpillRef', ref: 'A1' },
            { kind: 'RowRef', ref: '1' },
            { kind: 'ColumnRef', ref: 'A' },
            { kind: 'RangeRef', refKind: 'cells', start: 'B1', end: 'B2' },
            { kind: 'NameRef', name: 'x' },
          ],
        },
      ],
    }

    expect(optimizeFormula(metadataAwareLet)).toEqual({
      kind: 'CallExpr',
      callee: 'CONCAT',
      args: [
        { kind: 'StructuredRef', tableName: 'Sales', columnName: 'Amount' },
        { kind: 'SpillRef', ref: 'A1' },
        { kind: 'RowRef', ref: '1' },
        { kind: 'ColumnRef', ref: 'A' },
        { kind: 'RangeRef', refKind: 'cells', start: 'B1', end: 'B2' },
        { kind: 'NumberLiteral', value: 1 },
      ],
    })

    const invokeInsideLet: FormulaNode = {
      kind: 'CallExpr',
      callee: 'LET',
      args: [
        { kind: 'NameRef', name: 'x' },
        { kind: 'NumberLiteral', value: 2 },
        {
          kind: 'InvokeExpr',
          callee: {
            kind: 'CallExpr',
            callee: 'LAMBDA',
            args: [
              { kind: 'NameRef', name: 'y' },
              {
                kind: 'BinaryExpr',
                operator: '+',
                left: { kind: 'NameRef', name: 'x' },
                right: { kind: 'NameRef', name: 'y' },
              },
            ],
          },
          args: [{ kind: 'SpillRef', ref: 'A1' }],
        },
      ],
    }
    expect(optimizeFormula(invokeInsideLet)).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'NumberLiteral', value: 2 },
      right: { kind: 'SpillRef', ref: 'A1' },
    })
  })
})
