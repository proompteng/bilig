import { describe, expect, it } from 'vitest'
import {
  formulaMayContainFullRecalcPreservableUnavailableFormulaCall,
  formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc,
} from '../snapshot/unsupported-formula-cache.js'

describe('unsupported formula cache preservation', () => {
  it('uses a cheap marker gate for full-recalc-preservable cached formulas', () => {
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('SUM(A1:A10)')).toBe(false)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('UNKNOWNFUNC(A1)')).toBe(false)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_FV(A1,"Industry")')).toBe(true)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_fv(A1,"Industry")')).toBe(true)
    expect(formulaMayContainFullRecalcPreservableUnavailableFormulaCall('_xldudf_WISEPRICE(B1,"Shares Outstanding")')).toBe(true)
  })

  it('preserves only supported imported-cache markers during full recalculation', () => {
    const definedNames = new Set<string>()

    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('SUM(1,2)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('UNKNOWNFUNC(42)', definedNames)).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_FV(A1,"Industry")', definedNames)).toBe(true)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_fv(A1,"Industry")', definedNames)).toBe(true)
    expect(
      formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_xldudf_WISEPRICE(B1,"Shares Outstanding")', definedNames),
    ).toBe(true)
  })

  it('uses the AST walk as the authority when marker text is present', () => {
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('"_FV(A1)"', new Set())).toBe(false)
    expect(formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc('_FV(A1)', new Set(['_FV']))).toBe(false)
  })
})
