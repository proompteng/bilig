import { describe, expect, it } from 'vitest'
import { parseFormula } from '../index.js'
import {
  getNativeAxisAggregateCode,
  getNativeRunningFoldCode,
  isCellRangeNode,
  isCellVectorNode,
  isNativeMakearraySumLambda,
  isWasmSafeBuiltinArgs,
  isWasmSafeBuiltinArity,
} from '../binder-wasm-rules.js'
import type { FormulaNode } from '../ast.js'

describe('binder wasm rules', () => {
  it('detects native lambda patterns and cell range shapes', () => {
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,SUM(x))'))).toBe(1)
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,AVERAGE(x))'))).toBe(2)
    expect(getNativeAxisAggregateCode(parseFormula('LAMBDA(x,SUM(y))'))).toBeNull()

    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,acc+x)'))).toBe(1)
    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,x*acc)'))).toBe(2)
    expect(getNativeRunningFoldCode(parseFormula('LAMBDA(acc,x,acc-x)'))).toBeNull()

    expect(isNativeMakearraySumLambda(parseFormula('LAMBDA(r,c,r+c)'))).toBe(true)
    expect(isNativeMakearraySumLambda(parseFormula('LAMBDA(r,c,r-c)'))).toBe(false)

    expect(isCellRangeNode(parseFormula('A1:B3'))).toBe(true)
    expect(isCellRangeNode(parseFormula('A:B'))).toBe(false)
    expect(isCellVectorNode(parseFormula('A1:A4'))).toBe(true)
    expect(isCellVectorNode(parseFormula('A1:B4'))).toBe(false)
  })

  it('keeps representative arity and wasm-argument policies stable', () => {
    expect(isWasmSafeBuiltinArity('COUNTIFS', 4)).toBe(true)
    expect(isWasmSafeBuiltinArity('COUNTIFS', 3)).toBe(false)
    expect(isWasmSafeBuiltinArity('LOOKUP', 2)).toBe(true)
    expect(isWasmSafeBuiltinArity('LOOKUP', 4)).toBe(false)
    expect(isWasmSafeBuiltinArity('NA', 0)).toBe(true)
    expect(isWasmSafeBuiltinArity('NA', 1)).toBe(false)

    const deps = {
      isWasmSafe(node: FormulaNode, allowRange = false): boolean {
        if (
          node.kind === 'NumberLiteral' ||
          node.kind === 'BooleanLiteral' ||
          node.kind === 'StringLiteral' ||
          node.kind === 'ErrorLiteral'
        ) {
          return true
        }
        return node.kind === 'RangeRef' ? allowRange : false
      },
    }

    expect(isWasmSafeBuiltinArgs('COUNTIFS', parseFormula('COUNTIFS(A1:A4,1,B1:B4,2)').args, deps)).toBe(true)
    expect(isWasmSafeBuiltinArgs('LOOKUP', parseFormula('LOOKUP(1,A1:A4,B1:B4)').args, deps)).toBe(true)
    expect(isWasmSafeBuiltinArgs('LOOKUP', parseFormula('LOOKUP(A1:A2,A1:A4)').args, deps)).toBe(false)
    expect(isWasmSafeBuiltinArgs('TEXTJOIN', parseFormula('TEXTJOIN(",",TRUE,A1:A4)').args, deps)).toBe(true)
  })
})
