import { afterEach, describe, expect, it, vi } from 'vitest'
import { BuiltinId, ErrorCode, ValueTag } from '@bilig/protocol'
import { getBuiltin, getBuiltinId } from '../builtins.js'
import { getLookupBuiltin } from '../builtins/lookup.js'
import type { ArrayValue } from '../runtime-values.js'
import { placeholderBuiltinNames, protocolPlaceholderBuiltinNames } from '../builtins/placeholder.js'
import { clearExternalFunctionAdapters, installExternalFunctionAdapter } from '../external-function-adapter.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins', () => {
  it('supports CHOOSE, COUNTBLANK, and bitwise builtins', () => {
    const CHOOSE = getBuiltin('CHOOSE')!
    const COUNTBLANK = getBuiltin('COUNTBLANK')!
    const BITAND = getBuiltin('BITAND')!
    const BITOR = getBuiltin('BITOR')!
    const BITXOR = getBuiltin('BITXOR')!
    const BITLSHIFT = getBuiltin('BITLSHIFT')!
    const BITRSHIFT = getBuiltin('BITRSHIFT')!

    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: 'zero', stringId: 1 },
        {
          tag: ValueTag.Number,
          value: 10,
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
        {
          tag: ValueTag.Number,
          value: 20,
        },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(CHOOSE({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(COUNTBLANK({ tag: ValueTag.Empty }, { tag: ValueTag.String, value: 'x', stringId: 1 }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    expect(BITAND({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(BITOR({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(BITXOR({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(BITRSHIFT({ tag: ValueTag.Number, value: 16 }, { tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad' })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports numeric aggregates and error propagation', () => {
    const sum = getBuiltin('SUM')
    const avg = getBuiltin('AVG')
    const mod = getBuiltin('MOD')

    expect(sum?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    expect(
      avg?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.String, value: 'skip', stringId: 1 }, { tag: ValueTag.Empty }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })

    expect(sum?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(mod?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
  })

  it('supports SEQUENCE spill generation and validation', () => {
    const SEQUENCE = getBuiltin('SEQUENCE')!

    expect(
      SEQUENCE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 12 },
        { tag: ValueTag.Number, value: 14 },
        { tag: ValueTag.Number, value: 16 },
        { tag: ValueTag.Number, value: 18 },
        { tag: ValueTag.Number, value: 20 },
      ],
    })
    expect(SEQUENCE({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      SEQUENCE(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 400 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports Bessel engineering builtins', () => {
    const BESSELI = getBuiltin('BESSELI')!
    const BESSELJ = getBuiltin('BESSELJ')!
    const BESSELK = getBuiltin('BESSELK')!
    const BESSELY = getBuiltin('BESSELY')!
    const expectNumeric = (value: CellValue, expected: number) => {
      expect(value).toMatchObject({ tag: ValueTag.Number })
      if (value.tag !== ValueTag.Number) {
        throw new Error('Expected numeric builtin result')
      }
      expect(value.value).toBeCloseTo(expected, 7)
    }

    const besseli = BESSELI({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(besseli, 0.981666428)
    const besselj = BESSELJ({ tag: ValueTag.Number, value: 1.9 }, { tag: ValueTag.Number, value: 2 })
    expectNumeric(besselj, 0.329925728)
    const besselk = BESSELK({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(besselk, 0.277387804)
    const bessely = BESSELY({ tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(bessely, 0.145918138)
    expect(BESSELK({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports CONVERT and EUROCONVERT', () => {
    const CONVERT = getBuiltin('CONVERT')!
    const EUROCONVERT = getBuiltin('EUROCONVERT')!

    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.String, value: 'mi', stringId: 1 },
        { tag: ValueTag.String, value: 'km', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 9.656064,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 68 },
        { tag: ValueTag.String, value: 'F', stringId: 3 },
        { tag: ValueTag.String, value: 'C', stringId: 4 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 20,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 2.5 },
        { tag: ValueTag.String, value: 'ft', stringId: 5 },
        { tag: ValueTag.String, value: 'sec', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1.2 },
        { tag: ValueTag.String, value: 'DEM', stringId: 7 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0.61,
    })
    const triangulated = EUROCONVERT(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'FRF', stringId: 9 },
      { tag: ValueTag.String, value: 'DEM', stringId: 7 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Number, value: 3 },
    )
    expect(triangulated).toMatchObject({ tag: ValueTag.Number })
    if (triangulated.tag !== ValueTag.Number) {
      throw new Error('Expected EUROCONVERT result to be numeric')
    }
    expect(triangulated.value).toBeCloseTo(0.29728616, 12)
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'BAD', stringId: 10 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 32 },
        { tag: ValueTag.String, value: 'F', stringId: 11 },
        { tag: ValueTag.String, value: 'K', stringId: 12 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 273.15,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'F', stringId: 11 },
        { tag: ValueTag.String, value: 'm', stringId: 13 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: '??', stringId: 14 },
        { tag: ValueTag.String, value: 'm', stringId: 13 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 3.5 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    })
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'FRF', stringId: 9 },
        { tag: ValueTag.String, value: 'DEM', stringId: 7 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports boolean and string builtins and builtin ids', () => {
    expect(getBuiltin('AND')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })

    expect(getBuiltin('OR')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })

    expect(getBuiltin('NOT')?.({ tag: ValueTag.Boolean, value: false })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })

    expect(getBuiltin('LEN')?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })

    expect(
      getBuiltin('CONCAT')?.(
        { tag: ValueTag.String, value: 'alpha', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'alpha2', stringId: 0 })

    expect(
      getBuiltin('EXACT')?.({ tag: ValueTag.String, value: 'Alpha', stringId: 1 }, { tag: ValueTag.String, value: 'Alpha', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Boolean, value: true })

    expect(
      getBuiltin('EXACT')?.({ tag: ValueTag.String, value: 'Alpha', stringId: 1 }, { tag: ValueTag.String, value: 'alpha', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Boolean, value: false })

    expect(getBuiltin('LEFT')?.({ tag: ValueTag.String, value: 'alpha', stringId: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: 'alp',
      stringId: 0,
    })

    expect(
      getBuiltin('TEXTBEFORE')?.(
        { tag: ValueTag.String, value: 'alpha-beta', stringId: 1 },
        { tag: ValueTag.String, value: '-', stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'alpha', stringId: 0 })

    expect(
      getBuiltin('IFERROR')?.({ tag: ValueTag.Error, code: ErrorCode.Div0 }, { tag: ValueTag.String, value: 'fallback', stringId: 1 }),
    ).toEqual({ tag: ValueTag.String, value: 'fallback', stringId: 1 })

    expect(
      getBuiltin('DATE')?.({ tag: ValueTag.Number, value: 2026 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 15 }),
    ).toEqual({ tag: ValueTag.Number, value: 46096 })

    expect(
      getBuiltin('AVERAGE')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 6 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })

    expect(getBuiltinId('sum')).toBe(BuiltinId.Sum)
    expect(getBuiltinId('concat')).toBe(BuiltinId.Concat)
    expect(getBuiltinId('choose')).toBe(BuiltinId.Choose)
    expect(getBuiltinId('countblank')).toBe(BuiltinId.Countblank)
    expect(getBuiltinId('lenb')).toBe(BuiltinId.Lenb)
    expect(getBuiltinId('leftb')).toBe(BuiltinId.Leftb)
    expect(getBuiltinId('midb')).toBe(BuiltinId.Midb)
    expect(getBuiltinId('rightb')).toBe(BuiltinId.Rightb)
    expect(getBuiltinId('findb')).toBe(BuiltinId.Findb)
    expect(getBuiltinId('searchb')).toBe(BuiltinId.Searchb)
    expect(getBuiltinId('replaceb')).toBe(BuiltinId.Replaceb)
    expect(getBuiltinId('address')).toBe(BuiltinId.Address)
    expect(getBuiltinId('days360')).toBe(BuiltinId.Days360)
    expect(getBuiltinId('dollar')).toBe(BuiltinId.Dollar)
    expect(getBuiltinId('dollarde')).toBe(BuiltinId.Dollarde)
    expect(getBuiltinId('dollarfr')).toBe(BuiltinId.Dollarfr)
    expect(getBuiltinId('yearfrac')).toBe(BuiltinId.Yearfrac)
    expect(getBuiltinId('isoweeknum')).toBe(BuiltinId.Isoweeknum)
    expect(getBuiltinId('timevalue')).toBe(BuiltinId.Timevalue)
    expect(getBuiltinId('textbefore')).toBe(BuiltinId.Textbefore)
    expect(getBuiltinId('textafter')).toBe(BuiltinId.Textafter)
    expect(getBuiltinId('textjoin')).toBe(BuiltinId.Textjoin)
    expect(getBuiltinId('textsplit')).toBe(BuiltinId.Textsplit)
    expect(getBuiltinId('correl')).toBe(BuiltinId.Correl)
    expect(getBuiltinId('covar')).toBe(BuiltinId.Covar)
    expect(getBuiltinId('pearson')).toBe(BuiltinId.Pearson)
    expect(getBuiltinId('covariance.p')).toBe(BuiltinId.CovarianceP)
    expect(getBuiltinId('covariance.s')).toBe(BuiltinId.CovarianceS)
    expect(getBuiltinId('forecast')).toBe(BuiltinId.Forecast)
    expect(getBuiltinId('intercept')).toBe(BuiltinId.Intercept)
    expect(getBuiltinId('median')).toBe(BuiltinId.Median)
    expect(getBuiltinId('small')).toBe(BuiltinId.Small)
    expect(getBuiltinId('large')).toBe(BuiltinId.Large)
    expect(getBuiltinId('percentile')).toBe(BuiltinId.Percentile)
    expect(getBuiltinId('percentile.inc')).toBe(BuiltinId.PercentileInc)
    expect(getBuiltinId('percentile.exc')).toBe(BuiltinId.PercentileExc)
    expect(getBuiltinId('percentrank')).toBe(BuiltinId.Percentrank)
    expect(getBuiltinId('percentrank.inc')).toBe(BuiltinId.PercentrankInc)
    expect(getBuiltinId('percentrank.exc')).toBe(BuiltinId.PercentrankExc)
    expect(getBuiltinId('quartile')).toBe(BuiltinId.Quartile)
    expect(getBuiltinId('quartile.inc')).toBe(BuiltinId.QuartileInc)
    expect(getBuiltinId('quartile.exc')).toBe(BuiltinId.QuartileExc)
    expect(getBuiltinId('mode.mult')).toBe(BuiltinId.ModeMult)
    expect(getBuiltinId('frequency')).toBe(BuiltinId.Frequency)
    expect(getBuiltinId('rank')).toBe(BuiltinId.Rank)
    expect(getBuiltinId('rank.eq')).toBe(BuiltinId.RankEq)
    expect(getBuiltinId('rank.avg')).toBe(BuiltinId.RankAvg)
    expect(getBuiltinId('rsq')).toBe(BuiltinId.Rsq)
    expect(getBuiltinId('slope')).toBe(BuiltinId.Slope)
    expect(getBuiltinId('steyx')).toBe(BuiltinId.Steyx)
    expect(getBuiltinId('disc')).toBe(BuiltinId.Disc)
    expect(getBuiltinId('intrate')).toBe(BuiltinId.Intrate)
    expect(getBuiltinId('received')).toBe(BuiltinId.Received)
    expect(getBuiltinId('irr')).toBe(BuiltinId.Irr)
    expect(getBuiltinId('mirr')).toBe(BuiltinId.Mirr)
    expect(getBuiltinId('xnpv')).toBe(BuiltinId.Xnpv)
    expect(getBuiltinId('xirr')).toBe(BuiltinId.Xirr)
    expect(getBuiltinId('base')).toBe(BuiltinId.Base)
    expect(getBuiltinId('decimal')).toBe(BuiltinId.Decimal)
    expect(getBuiltinId('convert')).toBe(BuiltinId.Convert)
    expect(getBuiltinId('euroconvert')).toBe(BuiltinId.Euroconvert)
    expect(getBuiltinId('bitand')).toBe(BuiltinId.Bitand)
    expect(getBuiltinId('bitor')).toBe(BuiltinId.Bitor)
    expect(getBuiltinId('bitxor')).toBe(BuiltinId.Bitxor)
    expect(getBuiltinId('bitlshift')).toBe(BuiltinId.Bitlshift)
    expect(getBuiltinId('bitrshift')).toBe(BuiltinId.Bitrshift)
    expect(getBuiltinId('besseli')).toBe(BuiltinId.Besseli)
    expect(getBuiltinId('besselj')).toBe(BuiltinId.Besselj)
    expect(getBuiltinId('besselk')).toBe(BuiltinId.Besselk)
    expect(getBuiltinId('bessely')).toBe(BuiltinId.Bessely)
    expect(getBuiltinId('use.the.countif')).toBe(BuiltinId.Countif)
    expect(getBuiltinId('')).toBeUndefined()
    expect(getBuiltin('missing')).toBeUndefined()
  })

  it('supports the remaining scalar numeric builtins and conditional defaults', () => {
    expect(getBuiltin('MIN')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })

    expect(
      getBuiltin('MAX')?.(
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(
      getBuiltin('COUNT')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(
      getBuiltin('COUNTA')?.(
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'x', stringId: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(getBuiltin('ABS')?.({ tag: ValueTag.Number, value: -3.4 })).toEqual({
      tag: ValueTag.Number,
      value: 3.4,
    })
    expect(getBuiltin('INT')?.({ tag: ValueTag.Number, value: -3.1 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('ROUND')?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3.15,
    })
    expect(getBuiltin('ROUNDDOWN')?.({ tag: ValueTag.Number, value: -3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -3.14,
    })
    expect(getBuiltin('ROUND')?.({ tag: ValueTag.Number, value: 3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3.15,
    })
    expect(getBuiltin('FLOOR')?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('FLOOR')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('CEILING')?.({ tag: ValueTag.Number, value: 3.1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('CEILING')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })

    expect(getBuiltin('IF')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'truthy', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'truthy',
      stringId: 1,
    })

    expect(getBuiltin('IF')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Empty })
  })

  it('supports radix conversion and complex engineering builtins', () => {
    expect(getBuiltin('BIN2DEC')?.({ tag: ValueTag.String, value: '1111111111', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(getBuiltin('DEC2BIN')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: 0,
    })
    expect(getBuiltin('BIN2HEX')?.({ tag: ValueTag.String, value: '1111111111', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'FFFFFFFFFF',
      stringId: 0,
    })
    expect(getBuiltin('HEX2BIN')?.({ tag: ValueTag.String, value: 'A', stringId: 1 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: 0,
    })
    expect(getBuiltin('HEX2DEC')?.({ tag: ValueTag.String, value: 'FFFFFFFFFF', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(getBuiltin('OCT2HEX')?.({ tag: ValueTag.String, value: '17', stringId: 1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.String,
      value: '000F',
      stringId: 0,
    })

    expect(
      getBuiltin('COMPLEX')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: -4 },
        { tag: ValueTag.String, value: 'j', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: '3-4j', stringId: 0 })
    expect(getBuiltin('IMREAL')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('IMAGINARY')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(getBuiltin('IMABS')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(getBuiltin('IMARGUMENT')?.({ tag: ValueTag.String, value: 'i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 2,
    })
    expect(getBuiltin('IMCONJUGATE')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '3-4i',
      stringId: 0,
    })
    expect(
      getBuiltin('IMSUM')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 }, { tag: ValueTag.String, value: '-1+2i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '2+6i', stringId: 0 })
    expect(
      getBuiltin('IMPRODUCT')?.({ tag: ValueTag.String, value: '1+i', stringId: 1 }, { tag: ValueTag.String, value: '1-i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '2', stringId: 0 })
    expect(
      getBuiltin('IMDIV')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 }, { tag: ValueTag.String, value: '1-i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '-0.5+3.5i', stringId: 0 })
    expect(getBuiltin('IMSQRT')?.({ tag: ValueTag.String, value: '-4', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '2i',
      stringId: 0,
    })
    expect(getBuiltin('IMSIN')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMCOS')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMSECH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
  })

  it('covers the remaining complex engineering and value-classification builtins', () => {
    expect(getBuiltin('IMEXP')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMLN')?.({ tag: ValueTag.String, value: '1', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMLOG10')?.({ tag: ValueTag.String, value: '10', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMLOG2')?.({ tag: ValueTag.String, value: '8', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '3',
      stringId: 0,
    })
    expect(getBuiltin('IMPOWER')?.({ tag: ValueTag.String, value: '2', stringId: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: '8',
      stringId: 0,
    })
    expect(getBuiltin('IMTAN')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMSINH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMCOSH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMSEC')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMCSC')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMCOT')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMCSCH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMLOG10')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('IMPOWER')?.({ tag: ValueTag.String, value: '1+i', stringId: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 1999 })).toEqual({
      tag: ValueTag.String,
      value: 'MCMXCIX',
      stringId: 0,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'MCMXCIX', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 1999,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('T')?.({ tag: ValueTag.String, value: 'text', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'text',
      stringId: 1,
    })
    expect(getBuiltin('T')?.({ tag: ValueTag.Number, value: 7 })).toEqual({ tag: ValueTag.Empty })
    expect(getBuiltin('T')?.({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(getBuiltin('ISOMITTED')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(getBuiltin('ISOMITTED')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('N')?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('N')?.({ tag: ValueTag.String, value: 'text', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('N')?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    const matrix: ArrayValue = {
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ],
    }
    expect(getBuiltin('TYPE')?.({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    expect(Reflect.apply(getBuiltin('TYPE')!, undefined, [matrix])).toEqual({
      tag: ValueTag.Number,
      value: 64,
    })
    expect(getBuiltin('DELTA')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('GESTEP')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('supports expanded math and numeric utility builtins', () => {
    expect(getBuiltin('SIN')?.({ tag: ValueTag.Number, value: Math.PI / 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('COS')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.Number, value: -9 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: -3.98 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: -3.9,
    })
    expect(getBuiltin('CEILING.MATH')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: 5 })).toEqual({
      tag: ValueTag.Number,
      value: 120,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(getBuiltin('GCD')?.({ tag: ValueTag.Number, value: 18 }, { tag: ValueTag.Number, value: 24 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('LCM')?.({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.Number,
      value: 24,
    })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 9,
    })
    expect(
      getBuiltin('PRODUCT')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('SUMSQ')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 13,
    })
    expect(
      getBuiltin('SERIESSUM')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 18 })
    expect(
      getBuiltin('BASE')?.({ tag: ValueTag.Number, value: 31 }, { tag: ValueTag.Number, value: 16 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.String, value: '001F', stringId: 0 })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '1F', stringId: 1 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 31,
    })
    expect(getBuiltin('BASE')?.({ tag: ValueTag.Number, value: 15 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.String,
      value: '1111',
      stringId: 0,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '  1111  ', stringId: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 15,
    })
    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.String,
      value: 'XIV',
      stringId: 0,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'XIV', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 14,
    })

    expect(getBuiltin('MUNIT')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    })

    const randArray = getBuiltin('RANDARRAY')?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Boolean, value: true },
    )
    expect(randArray).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    if (!(randArray && 'kind' in randArray && randArray.kind === 'array')) {
      throw new Error('expected RANDARRAY to return an array')
    }
    for (const value of randArray.values) {
      expect(value.tag).toBe(ValueTag.Number)
      expect(value.value).toBeGreaterThanOrEqual(3)
      expect(value.value).toBeLessThanOrEqual(7)
    }
  })

  it('covers math builtin edge cases and aggregate variants', () => {
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('EVEN')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ODD')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('LN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('SQRT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('COT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('FACTDOUBLE')?.({ tag: ValueTag.Number, value: 6 })).toEqual({
      tag: ValueTag.Number,
      value: 48,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('GCD')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('LCM')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: -10 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('MULTINOMIAL')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('RANDBETWEEN')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BASE')?.({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '1Z', stringId: 1 }, { tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'IIII', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('RANDARRAY')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('RANDARRAY')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    expect(getBuiltin('MUNIT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('SERIESSUM')?.(
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('SQRTPI')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 9 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(getBuiltin('SUBTOTAL')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 99 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers address, formatting, and mean helper builtins', () => {
    expect(
      getBuiltin('MAXA')?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      getBuiltin('MINA')?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "O'Brien", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "'O''Brien'!$AB2", stringId: 0 })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'R2C[28]', stringId: 0 })
    expect(
      getBuiltin('ADDRESS')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 28 }, { tag: ValueTag.Number, value: 5 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: -1234.567 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.6',
      stringId: 0,
    })
    expect(
      getBuiltin('DOLLAR')?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: '$1234.6', stringId: 0 })
    expect(
      getBuiltin('FIXED')?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: '1234.6', stringId: 0 })
    expect(getBuiltin('FIXED')?.({ tag: ValueTag.Number, value: 1234.567 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.08 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 1.5,
    })
    expect(getBuiltin('DOLLARFR')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 1.08,
    })
    expect(getBuiltin('DOLLARFR')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('GEOMEAN')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(getBuiltin('GEOMEAN')?.({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 3 / 1.75 })
    expect(getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers extended trigonometric and precise rounding builtins', () => {
    expect(getBuiltin('SINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COSH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('TANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ACOSH')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ATANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ACOT')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 4,
    })
    expect(getBuiltin('ACOT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 2,
    })
    expect(getBuiltin('ACOTH')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5 * Math.log(3),
    })
    expect(getBuiltin('COTH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSCH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('SEC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SECH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('FLOOR.MATH')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: 5.1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 5.1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('supports ACCRINT, ACCRINTM, AMORDEGRC, and AMORLINC', () => {
    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstInterest = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 30 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    )
    const cost = { tag: ValueTag.Number, value: 2000 }
    const salvage = { tag: ValueTag.Number, value: 10 }
    const period = { tag: ValueTag.Number, value: 4 }
    const rate = { tag: ValueTag.Number, value: 0.1 }
    const basis = { tag: ValueTag.Number, value: 0 }

    expect(issue?.tag).toBe(ValueTag.Number)
    expect(firstInterest?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)

    const firstAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    )
    expect(firstAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(firstAccrual?.tag === ValueTag.Number ? firstAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    const omittedBasisAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
    )
    expect(omittedBasisAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(omittedBasisAccrual?.tag === ValueTag.Number ? omittedBasisAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    const fullAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    )
    const shortAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
      { tag: ValueTag.Boolean, value: false },
    )
    expect(fullAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(shortAccrual).toMatchObject({ tag: ValueTag.Number })
    const shortAccrualValue = shortAccrual?.tag === ValueTag.Number ? shortAccrual.value : Number.NaN
    const fullAccrualValue = fullAccrual?.tag === ValueTag.Number ? fullAccrual.value : Number.NaN
    expect(shortAccrualValue).toBeLessThan(fullAccrualValue)

    const maturityAccrual = getBuiltin('ACCRINTM')?.(issue, settlement, rate, undefined, basis)
    expect(maturityAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(maturityAccrual?.tag === ValueTag.Number ? maturityAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    expect(getBuiltin('AMORLINC')?.(cost, issue, settlement, salvage, period, rate, basis)).toEqual({ tag: ValueTag.Number, value: 200 })

    expect(getBuiltin('AMORDEGRC')?.(cost, issue, settlement, salvage, period, rate, basis)).toEqual({ tag: ValueTag.Number, value: 163 })

    expect(
      getBuiltin('ACCRINT')?.(issue, settlement, issue, rate, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      getBuiltin('ACCRINT')?.(
        issue,
        settlement,
        settlement,
        rate,
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers ACCRINT and ACCRINTM basis variants and invalid argument branches', () => {
    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstInterest = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const rate = { tag: ValueTag.Number, value: 0.08 }
    const par = { tag: ValueTag.Number, value: 1000 }
    const frequency = { tag: ValueTag.Number, value: 2 }

    expect(issue?.tag).toBe(ValueTag.Number)
    expect(firstInterest?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)

    for (const basis of [0, 1, 2, 3, 4]) {
      expect(
        getBuiltin('ACCRINT')?.(issue, firstInterest, settlement, rate, par, frequency, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number })
      expect(
        getBuiltin('ACCRINTM')?.(issue, settlement, rate, par, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number })
    }

    expect(
      getBuiltin('ACCRINT')?.(issue, firstInterest, settlement, rate, par, frequency, {
        tag: ValueTag.Number,
        value: 5,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('ACCRINTM')?.(issue, settlement, rate, par, { tag: ValueTag.Number, value: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers AMORLINC and AMORDEGRC branch-heavy scenarios', () => {
    const cost = { tag: ValueTag.Number, value: 1000 }
    const datePurchased = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstPeriod = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const basis = { tag: ValueTag.Number, value: 0 }

    expect(datePurchased?.tag).toBe(ValueTag.Number)
    expect(firstPeriod?.tag).toBe(ValueTag.Number)

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 75 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.2 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 240 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.3 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 247 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.5 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 250 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1.2 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 100 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 200 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers combinatorics, product, quotient, and financial validation edge branches', () => {
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      getBuiltin('GCD')?.({ tag: ValueTag.Number, value: 54 }, { tag: ValueTag.Number, value: 24.9 }, { tag: ValueTag.Number, value: 6 }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(
      getBuiltin('LCM')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3.8 }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(
      getBuiltin('MULTINOMIAL')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(getBuiltin('PRODUCT')?.()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(getBuiltin('PRODUCT')?.({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    )
    expect(issue?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)
    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.String, value: 'bad', stringId: 2 },
        issue,
        settlement,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('AMORLINC')?.(
        { tag: ValueTag.Number, value: 1000 },
        issue,
        settlement,
        { tag: ValueTag.String, value: 'bad', stringId: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers bitwise, integer, and rounding validation branches', () => {
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.String, value: 'bad', stringId: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITLSHIFT')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITLSHIFT')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 33 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(getBuiltin('BITRSHIFT')?.({ tag: ValueTag.String, value: 'bad', stringId: 4 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('INT')?.({ tag: ValueTag.String, value: 'bad', stringId: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 6 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ROUNDDOWN')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 7 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 8 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: -12.34 })).toEqual({
      tag: ValueTag.Number,
      value: -12,
    })
  })

  it('covers ceiling, floor, parity, factorial, and combinatoric branches', () => {
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(
      getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.String, value: 'bad', stringId: 9 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('BITAND')?.({ tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITAND')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.String, value: 'bad', stringId: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 11 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('EVEN')?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('ODD')?.({ tag: ValueTag.Number, value: -2 })).toEqual({
      tag: ValueTag.Number,
      value: -3,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('FACTDOUBLE')?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('GCD')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('LCM')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers logarithmic, hyperbolic, and sign-related math edge branches', () => {
    expect(getBuiltin('HARMEAN')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('LOG10')?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('LOG10')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 100 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(
      getBuiltin('ACOT')?.({
        tag: ValueTag.Number,
        value: 0,
      }),
    ).toEqual({ tag: ValueTag.Number, value: Math.PI / 2 })
    expect(getBuiltin('ACOTH')?.({ tag: ValueTag.Number, value: 0.5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('COT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('COTH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSCH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('SECH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.Number, value: -42 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.String, value: 'bad', stringId: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers direct trig, exponential, and positive rounding builtin paths', () => {
    expect(getBuiltin('SIN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COS')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('TAN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASIN')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 2, 12),
    })
    expect(getBuiltin('ACOS')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ATAN')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    })
    expect(getBuiltin('ATAN2')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    })
    expect(getBuiltin('DEGREES')?.({ tag: ValueTag.Number, value: Math.PI })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(180, 12),
    })
    expect(getBuiltin('RADIANS')?.({ tag: ValueTag.Number, value: 180 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    })
    expect(getBuiltin('EXP')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.E, 12),
    })
    expect(getBuiltin('LN')?.({ tag: ValueTag.Number, value: Math.E })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(getBuiltin('SQRT')?.({ tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('PI')?.()).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    })
    expect(getBuiltin('SINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COSH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('TANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ACOSH')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ATANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })

    expect(getBuiltin('FLOOR.MATH')?.({ tag: ValueTag.Number, value: 5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('CEILING.MATH')?.({ tag: ValueTag.Number, value: 5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITAND')?.({ tag: ValueTag.String, value: 'bad', stringId: 13 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.String, value: 'bad', stringId: 14 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers AVERAGEA and SUBTOTAL aggregate dispatch branches', () => {
    expect(
      getBuiltin('AVERAGEA')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: 'skip', stringId: 15 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0.75 })

    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      getBuiltin('SUBTOTAL')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'skip', stringId: 17 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(Math.sqrt(2), 12) })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 11 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
  })

  it('covers aggregate aliases and formatting validation branches', () => {
    expect(getBuiltin('AVERAGEA')?.({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('AVERAGE')?.()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(getBuiltin('AVG')?.()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(getBuiltin('AVERAGE')?.({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('AVG')?.({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(getBuiltin('MAXA')?.({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('MINA')?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      getBuiltin('ADDRESS')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 5 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1.5 })).toEqual({
      tag: ValueTag.String,
      value: '$10.0',
      stringId: 0,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.6 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports ADDRESS and DOLLAR formatting edge cases', () => {
    const ADDRESS = getBuiltin('ADDRESS')!
    expect(ADDRESS({ tag: ValueTag.Number, value: 12 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: '$C$12',
      stringId: 0,
    })
    expect(ADDRESS({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.String,
      value: 'B$7',
      stringId: 0,
    })
    expect(
      ADDRESS(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: 'R4C5',
      stringId: 0,
    })

    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: -1234.5 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.5',
      stringId: 0,
    })
    expect(
      getBuiltin('DOLLAR')?.(
        { tag: ValueTag.Number, value: 1234.56 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: '$1235',
      stringId: 0,
    })
  })

  it('covers the new type, statistical, distribution, and combinatoric builtins', () => {
    const T = getBuiltin('T')!
    const N = getBuiltin('N')!
    const TYPE = getBuiltin('TYPE')!
    const DELTA = getBuiltin('DELTA')!
    const GESTEP = getBuiltin('GESTEP')!
    const GAUSS = getBuiltin('GAUSS')!
    const PHI = getBuiltin('PHI')!
    const STANDARDIZE = getBuiltin('STANDARDIZE')!
    const MODE = getBuiltin('MODE')!
    const MODE_SNGL = getBuiltin('MODE.SNGL')!
    const STDEV = getBuiltin('STDEV')!
    const STDEV_S = getBuiltin('STDEV.S')!
    const STDEVP = getBuiltin('STDEVP')!
    const STDEV_P = getBuiltin('STDEV.P')!
    const STDEVA = getBuiltin('STDEVA')!
    const STDEVPA = getBuiltin('STDEVPA')!
    const VAR = getBuiltin('VAR')!
    const VAR_S = getBuiltin('VAR.S')!
    const VARP = getBuiltin('VARP')!
    const VAR_P = getBuiltin('VAR.P')!
    const VARA = getBuiltin('VARA')!
    const VARPA = getBuiltin('VARPA')!
    const SKEW = getBuiltin('SKEW')!
    const SKEW_P = getBuiltin('SKEW.P')!
    const KURT = getBuiltin('KURT')!
    const NORMDIST = getBuiltin('NORMDIST')!
    const NORM_DIST = getBuiltin('NORM.DIST')!
    const NORMINV = getBuiltin('NORMINV')!
    const NORM_INV = getBuiltin('NORM.INV')!
    const NORMSDIST = getBuiltin('NORMSDIST')!
    const NORM_S_DIST = getBuiltin('NORM.S.DIST')!
    const NORMSINV = getBuiltin('NORMSINV')!
    const NORM_S_INV = getBuiltin('NORM.S.INV')!
    const LOGINV = getBuiltin('LOGINV')!
    const LOGNORMDIST = getBuiltin('LOGNORMDIST')!
    const LOGNORM_DIST = getBuiltin('LOGNORM.DIST')!
    const LOGNORM_INV = getBuiltin('LOGNORM.INV')!
    const CONFIDENCE_NORM = getBuiltin('CONFIDENCE.NORM')!
    const PERMUT = getBuiltin('PERMUT')!
    const PERMUTATIONA = getBuiltin('PERMUTATIONA')!

    expect(T({ tag: ValueTag.String, value: 'alpha', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'alpha',
      stringId: 1,
    })
    expect(T({ tag: ValueTag.Number, value: 42 })).toEqual({ tag: ValueTag.Empty })
    expect(T({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(N({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(N({ tag: ValueTag.String, value: 'alpha', stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(N({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    expect(TYPE({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(TYPE({ tag: ValueTag.String, value: 'alpha', stringId: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(TYPE({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(TYPE({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    const arrayValue: ArrayValue = {
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Number, value: 1 }],
    }
    expect(TYPE(arrayValue)).toEqual({
      tag: ValueTag.Number,
      value: 64,
    })

    expect(DELTA({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(DELTA({ tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(DELTA({ tag: ValueTag.String, value: 'bad', stringId: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(GESTEP({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(GESTEP({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(GESTEP({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.String, value: 'bad', stringId: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(GAUSS({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
    expect(PHI({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    })
    expect(
      STANDARDIZE({ tag: ValueTag.Number, value: 42 }, { tag: ValueTag.Number, value: 40 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      STANDARDIZE({ tag: ValueTag.Number, value: 42 }, { tag: ValueTag.Number, value: 40 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(MODE({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(
      MODE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(MODE_SNGL({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(MODE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(MODE_SNGL({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      STDEV(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    })
    expect(
      STDEV_S(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      STDEVP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(1.25), 12),
    })
    expect(STDEV_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      STDEVA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 7 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      STDEVPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 8 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(2 / 3), 12),
    })
    expect(STDEV({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(STDEV_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(STDEVP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(STDEV_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(STDEVA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(STDEVPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      VAR(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    })
    expect(VAR_S({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      VARP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.25, 12),
    })
    expect(VAR_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      VARA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 9 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      VARPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    })
    expect(VAR({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(VAR_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(VARP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(VAR_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(VARA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(VARPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      SKEW(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      SKEW_P(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      KURT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    })
    expect(KURT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(SKEW({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(SKEW_P({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(KURT({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      NORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(
      NORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    })
    expect(
      NORMINV({ tag: ValueTag.Number, value: 0.8413447460685429 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(
      NORM_INV({ tag: ValueTag.Number, value: 0.8413447460685429 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(NORMSDIST({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(NORM_S_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: false })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    })
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.001 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    })
    expect(NORM_S_INV({ tag: ValueTag.Number, value: 0.999 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3.090232306167813, 8),
    })
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      NORMINV({ tag: ValueTag.String, value: 'bad', stringId: 11 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NORM_S_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NORMSINV({ tag: ValueTag.String, value: 'bad', stringId: 13 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(LOGINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      LOGINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(
      LOGNORMDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    })
    expect(
      LOGNORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    })
    expect(
      LOGNORM_INV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(
      CONFIDENCE_NORM({ tag: ValueTag.Number, value: 0.05 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    })
    expect(
      CONFIDENCE_NORM({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(LOGNORMDIST({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 })).toEqual(
      {
        tag: ValueTag.Error,
        code: ErrorCode.Value,
      },
    )

    expect(PERMUT({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 60,
    })
    expect(PERMUTATIONA({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(PERMUTATIONA({ tag: ValueTag.String, value: 'bad', stringId: 14 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(PERMUT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltinId('norm.dist')).toBe(BuiltinId.NormDist)
    expect(getBuiltinId('norm.s.inv')).toBe(BuiltinId.NormSInv)
    expect(getBuiltinId('confidence.norm')).toBe(BuiltinId.ConfidenceNorm)
    expect(getBuiltinId('confidence.t')).toBe(BuiltinId.ConfidenceT)
    expect(getBuiltinId('gamma.inv')).toBe(BuiltinId.GammaInv)
    expect(getBuiltinId('gammainv')).toBe(BuiltinId.Gammainv)
    expect(getBuiltinId('permutationa')).toBe(BuiltinId.Permutationa)
    expect(getBuiltinId('chisq.test')).toBe(BuiltinId.ChisqTest)
    expect(getBuiltinId('chitest')).toBe(BuiltinId.Chitest)
    expect(getBuiltinId('legacy.chitest')).toBe(BuiltinId.LegacyChitest)
    expect(getBuiltinId('f.test')).toBe(BuiltinId.FTest)
    expect(getBuiltinId('ftest')).toBe(BuiltinId.Ftest)
    expect(getBuiltinId('z.test')).toBe(BuiltinId.ZTest)
    expect(getBuiltinId('ztest')).toBe(BuiltinId.Ztest)
    expect(getBuiltinId('workday.intl')).toBe(BuiltinId.WorkdayIntl)
    expect(getBuiltinId('networkdays.intl')).toBe(BuiltinId.NetworkdaysIntl)
    expect(getBuiltinId('numbervalue')).toBe(BuiltinId.Numbervalue)
    expect(getBuiltinId('valuetotext')).toBe(BuiltinId.Valuetotext)
    expect(getBuiltinId('asc')).toBe(BuiltinId.Asc)
    expect(getBuiltinId('jis')).toBe(BuiltinId.Jis)
    expect(getBuiltinId('dbcs')).toBe(BuiltinId.Dbcs)
    expect(getBuiltinId('daverage')).toBe(BuiltinId.Daverage)
    expect(getBuiltinId('dcount')).toBe(BuiltinId.Dcount)
    expect(getBuiltinId('dcounta')).toBe(BuiltinId.Dcounta)
    expect(getBuiltinId('dget')).toBe(BuiltinId.Dget)
    expect(getBuiltinId('dmax')).toBe(BuiltinId.Dmax)
    expect(getBuiltinId('dmin')).toBe(BuiltinId.Dmin)
    expect(getBuiltinId('dproduct')).toBe(BuiltinId.Dproduct)
    expect(getBuiltinId('dstdev')).toBe(BuiltinId.Dstdev)
    expect(getBuiltinId('dstdevp')).toBe(BuiltinId.Dstdevp)
    expect(getBuiltinId('dsum')).toBe(BuiltinId.Dsum)
    expect(getBuiltinId('dvar')).toBe(BuiltinId.Dvar)
    expect(getBuiltinId('dvarp')).toBe(BuiltinId.Dvarp)
    expect(getBuiltinId('prob')).toBe(BuiltinId.Prob)
    expect(getBuiltinId('trimmean')).toBe(BuiltinId.Trimmean)
    expect(getBuiltinId('growth')).toBe(BuiltinId.Growth)
    expect(getBuiltinId('trend')).toBe(BuiltinId.Trend)
    expect(getBuiltinId('forecast.linear')).toBe(BuiltinId.Forecast)
  })

  it('supports the new statistical distribution builtins and aliases', () => {
    const ERF = getBuiltin('ERF')!
    const ERF_PRECISE = getBuiltin('ERF.PRECISE')!
    const ERFC = getBuiltin('ERFC')!
    const ERFC_PRECISE = getBuiltin('ERFC.PRECISE')!
    const FISHER = getBuiltin('FISHER')!
    const FISHERINV = getBuiltin('FISHERINV')!
    const GAMMALN = getBuiltin('GAMMALN')!
    const GAMMALN_PRECISE = getBuiltin('GAMMALN.PRECISE')!
    const GAMMA = getBuiltin('GAMMA')!
    const CONFIDENCE = getBuiltin('CONFIDENCE')!
    const EXPONDIST = getBuiltin('EXPONDIST')!
    const EXPON_DIST = getBuiltin('EXPON.DIST')!
    const POISSON = getBuiltin('POISSON')!
    const POISSON_DIST = getBuiltin('POISSON.DIST')!
    const WEIBULL = getBuiltin('WEIBULL')!
    const WEIBULL_DIST = getBuiltin('WEIBULL.DIST')!
    const GAMMADIST = getBuiltin('GAMMADIST')!
    const GAMMA_DIST = getBuiltin('GAMMA.DIST')!
    const GAMMA_INV = getBuiltin('GAMMA.INV')!
    const GAMMAINV = getBuiltin('GAMMAINV')!
    const CHIDIST = getBuiltin('CHIDIST')!
    const CHISQ_DIST_RT = getBuiltin('CHISQ.DIST.RT')!
    const CHISQ_DIST = getBuiltin('CHISQ.DIST')!
    const BETA_DIST = getBuiltin('BETA.DIST')!
    const BETA_INV = getBuiltin('BETA.INV')!
    const BETADIST = getBuiltin('BETADIST')!
    const BETAINV = getBuiltin('BETAINV')!
    const F_DIST = getBuiltin('F.DIST')!
    const F_DIST_RT = getBuiltin('F.DIST.RT')!
    const F_INV = getBuiltin('F.INV')!
    const F_INV_RT = getBuiltin('F.INV.RT')!
    const FDIST = getBuiltin('FDIST')!
    const FINV = getBuiltin('FINV')!
    const LEGACY_FDIST = getBuiltin('LEGACY.FDIST')!
    const LEGACY_FINV = getBuiltin('LEGACY.FINV')!
    const T_DIST = getBuiltin('T.DIST')!
    const T_DIST_RT = getBuiltin('T.DIST.RT')!
    const T_DIST_2T = getBuiltin('T.DIST.2T')!
    const TDIST = getBuiltin('TDIST')!
    const T_INV = getBuiltin('T.INV')!
    const T_INV_2T = getBuiltin('T.INV.2T')!
    const TINV = getBuiltin('TINV')!
    const T_TEST = getLookupBuiltin('T.TEST')!
    const TTEST = getLookupBuiltin('TTEST')!
    const CONFIDENCE_T = getBuiltin('CONFIDENCE.T')!
    const BINOMDIST = getBuiltin('BINOMDIST')!
    const BINOM_DIST = getBuiltin('BINOM.DIST')!
    const BINOM_DIST_RANGE = getBuiltin('BINOM.DIST.RANGE')!
    const CRITBINOM = getBuiltin('CRITBINOM')!
    const BINOM_INV = getBuiltin('BINOM.INV')!
    const HYPGEOMDIST = getBuiltin('HYPGEOMDIST')!
    const HYPGEOM_DIST = getBuiltin('HYPGEOM.DIST')!
    const NEGBINOMDIST = getBuiltin('NEGBINOMDIST')!
    const NEGBINOM_DIST = getBuiltin('NEGBINOM.DIST')!

    expect(ERF({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERF({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERF_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERFC({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    })
    expect(ERFC_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    })
    expect(ERF({ tag: ValueTag.String, value: 'bad', stringId: 15 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FISHER({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5493061443340549, 12),
    })
    expect(FISHERINV({ tag: ValueTag.Number, value: 0.5493061443340549 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(FISHERINV({ tag: ValueTag.String, value: 'bad', stringId: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(GAMMALN({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    })
    expect(GAMMALN_PRECISE({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    })
    expect(GAMMA({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(24, 10),
    })
    expect(
      CONFIDENCE({ tag: ValueTag.Number, value: 0.05 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.764892328404345, 12),
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      EXPONDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: false }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2706705664732254, 12),
    })
    expect(
      EXPON_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8646647167633873, 12),
    })
    expect(
      EXPONDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Boolean, value: false }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      POISSON({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Boolean, value: false }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.21376301724973648, 12),
    })
    expect(
      POISSON_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.7575761331330662, 12),
    })
    expect(
      POISSON({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Boolean, value: false }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2596002610238016, 12),
    })
    expect(
      WEIBULL_DIST(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.22119921692859512, 12),
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09196986029286061, 12),
    })
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08030139707139418, 12),
    })
    expect(
      GAMMA_INV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(GAMMA_INV({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMA_INV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMAINV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    })
    expect(CHIDIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    })
    expect(CHISQ_DIST_RT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    })
    expect(
      CHISQ_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.4421745996289252, 12),
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6854705810117458, 10),
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.4837646, 7),
    })
    expect(
      BETADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6854705810117458, 10),
    })
    expect(
      BETA_INV(
        { tag: ValueTag.Number, value: 0.6854705810117458 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      BETAINV(
        { tag: ValueTag.Number, value: 0.6854705810117458 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      F_DIST(
        { tag: ValueTag.Number, value: 15.2068649 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.99, 9),
    })
    expect(
      F_DIST(
        { tag: ValueTag.Number, value: 15.2068649 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0012238, 9),
    })
    expect(
      F_DIST_RT({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      FDIST({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      LEGACY_FDIST({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      F_INV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.10930991466299911, 8),
    })
    expect(
      F_INV_RT({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      FINV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      LEGACY_FINV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      T_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.75, 12),
    })
    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.25, 12),
    })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      T_DIST(
        { tag: ValueTag.String, value: 'bad', stringId: 23 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 24 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(TDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject(
      {
        tag: ValueTag.Number,
        value: expect.closeTo(0.25, 12),
      },
    )
    expect(T_INV({ tag: ValueTag.Number, value: 0.75 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(T_INV_2T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(TINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(
      T_TEST(
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 2 },
            { tag: ValueTag.Number, value: 4 },
          ],
        },
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 3 },
            { tag: ValueTag.Number, value: 3 },
          ],
        },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      TTEST(
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 2 },
            { tag: ValueTag.Number, value: 4 },
          ],
        },
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 3 },
            { tag: ValueTag.Number, value: 3 },
          ],
        },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.375, 12),
    })
    expect(
      BINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6875, 12),
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.78125, 12),
    })
    expect(
      CRITBINOM({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.7 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      CRITBINOM({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.999999999999 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      BINOM_INV({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.7 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    })
    expect(
      NEGBINOMDIST({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 0.5 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })

    expect(FISHER({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(GAMMA({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(CHIDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CHISQ_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Boolean, value: true }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(F_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      F_INV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.String, value: 'bad', stringId: 71 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      F_INV_RT(
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.String, value: 'bad', stringId: 72 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(CRITBINOM({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toEqual(
      {
        tag: ValueTag.Error,
        code: ErrorCode.Value,
      },
    )
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NEGBINOMDIST({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 1.5 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltinId('gamma.dist')).toBe(BuiltinId.GammaDist)
    expect(getBuiltinId('negbinom.dist')).toBe(BuiltinId.NegbinomDist)
    expect(getBuiltinId('binom.inv')).toBe(BuiltinId.BinomInv)
  })

  it('covers remaining Student t and binomial validation guards', () => {
    const T_DIST_RT = getBuiltin('T.DIST.RT')!
    const T_INV = getBuiltin('T.INV')!
    const T_INV_2T = getBuiltin('T.INV.2T')!
    const TDIST = getBuiltin('TDIST')!
    const BINOMDIST = getBuiltin('BINOMDIST')!
    const BINOM_DIST_RANGE = getBuiltin('BINOM.DIST.RANGE')!

    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_INV({ tag: ValueTag.String, value: 'bad', stringId: 401 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_INV_2T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(TDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(TDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers the new financial builtins and their error branches', () => {
    const EFFECT = getBuiltin('EFFECT')!
    const NOMINAL = getBuiltin('NOMINAL')!
    const PDURATION = getBuiltin('PDURATION')!
    const RRI = getBuiltin('RRI')!
    const FV = getBuiltin('FV')!
    const PV = getBuiltin('PV')!
    const PMT = getBuiltin('PMT')!
    const NPER = getBuiltin('NPER')!
    const RATE = getBuiltin('RATE')!
    const NPV = getBuiltin('NPV')!
    const IPMT = getBuiltin('IPMT')!
    const PPMT = getBuiltin('PPMT')!
    const ISPMT = getBuiltin('ISPMT')!
    const CUMIPMT = getBuiltin('CUMIPMT')!
    const CUMPRINC = getBuiltin('CUMPRINC')!
    const DATE = getBuiltin('DATE')!
    const FVSCHEDULE = getBuiltin('FVSCHEDULE')!
    const DB = getBuiltin('DB')!
    const DDB = getBuiltin('DDB')!
    const VDB = getBuiltin('VDB')!
    const SLN = getBuiltin('SLN')!
    const SYD = getBuiltin('SYD')!
    const DISC = getBuiltin('DISC')!
    const INTRATE = getBuiltin('INTRATE')!
    const RECEIVED = getBuiltin('RECEIVED')!
    const PRICEDISC = getBuiltin('PRICEDISC')!
    const YIELDDISC = getBuiltin('YIELDDISC')!
    const PRICEMAT = getBuiltin('PRICEMAT')!
    const YIELDMAT = getBuiltin('YIELDMAT')!
    const ODDFPRICE = getBuiltin('ODDFPRICE')!
    const ODDFYIELD = getBuiltin('ODDFYIELD')!
    const ODDLPRICE = getBuiltin('ODDLPRICE')!
    const ODDLYIELD = getBuiltin('ODDLYIELD')!
    const TBILLPRICE = getBuiltin('TBILLPRICE')!
    const TBILLYIELD = getBuiltin('TBILLYIELD')!
    const TBILLEQ = getBuiltin('TBILLEQ')!

    expect(EFFECT({ tag: ValueTag.Number, value: 0.12 }, { tag: ValueTag.Number, value: 12 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12682503013196977, 12),
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: 0.12682503013196977 }, { tag: ValueTag.Number, value: 12 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(
      PDURATION({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(
      RRI({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1, 12),
    })

    expect(
      FV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: -1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 12),
    })
    expect(
      FVSCHEDULE(
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 0.11 },
        { tag: ValueTag.Number, value: 0.1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1330.89, 12),
    })
    expect(
      DB(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3690, 12),
    })
    expect(
      DDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(384, 12),
    })
    expect(
      VDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(691.2, 12),
    })
    expect(
      PV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: 1420 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1000, 12),
    })
    expect(PV({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 })).toEqual({
      tag: ValueTag.Number,
      value: 1000,
    })
    expect(
      PMT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 1000 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904761, 12),
    })
    expect(PMT({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: -100,
    })
    expect(
      RATE({ tag: ValueTag.Number, value: 48 }, { tag: ValueTag.Number, value: -200 }, { tag: ValueTag.Number, value: 8000 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(RATE({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 }, { tag: ValueTag.Number, value: 1000 })).toEqual(
      {
        tag: ValueTag.Number,
        value: 0,
      },
    )
    expect(
      RATE(
        { tag: ValueTag.Number, value: 48 },
        { tag: ValueTag.Number, value: -200 },
        { tag: ValueTag.Number, value: 8000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(
      NPER(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: -576.1904761904761 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(NPER({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: -100 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(
      NPV({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(173.55371900826447, 12),
    })
    expect(
      IPMT(
        { tag: ValueTag.String, value: 'bad', stringId: 17 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -100,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.19047619047615, 12),
    })
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -50,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-11135.232130750845, 12),
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 12),
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-998.2910880208206, 12),
    })
    expect(SLN({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 1000,
    })
    expect(
      SYD(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1800,
    })
    expect(
      DISC(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(
      INTRATE(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 1030 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(
      RECEIVED(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0.12 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1030.9278350515465, 12),
    })
    expect(
      PRICEDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0525 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.79583333333333, 12),
    })
    expect(
      YIELDDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 99.795 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.05282257198685834, 12),
    })
    expect(
      PRICEMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 13,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.98449887555694, 12),
    })
    expect(
      YIELDMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 3,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 8,
          },
        ),
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 100.0123 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.060954333691538576, 12),
    })
    expect(
      ODDFPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0785 },
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(113.597717474079, 12),
    })
    expect(
      ODDFYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 84.5 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0772455415972989, 11),
    })
    expect(
      ODDLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 7,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0375 },
        { tag: ValueTag.Number, value: 0.0405 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.8782860147213, 12),
    })
    expect(
      ODDLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 20,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 12 },
          {
            tag: ValueTag.Number,
            value: 24,
          },
        ),
        { tag: ValueTag.Number, value: 0.0375 },
        { tag: ValueTag.Number, value: 99.875 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0451922356291692, 12),
    })
    expect(
      TBILLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.09 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(98.45, 12),
    })
    expect(
      TBILLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 98.45 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09141696292534264, 12),
    })
    expect(
      TBILLEQ(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0914 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09415149356594302, 12),
    })
    expect(
      DISC(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })

    expect(EFFECT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PDURATION({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(RRI({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(PMT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PV({ tag: ValueTag.String, value: 'bad', stringId: 39 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PMT({ tag: ValueTag.String, value: 'bad', stringId: 40 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1000 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(RATE({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      RATE(
        { tag: ValueTag.String, value: 'bad', stringId: 41 },
        { tag: ValueTag.Number, value: -200 },
        { tag: ValueTag.Number, value: 8000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NPER({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NPER(
        { tag: ValueTag.String, value: 'bad', stringId: 42 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NPV({ tag: ValueTag.String, value: 'bad', stringId: 11 }, { tag: ValueTag.Number, value: 100 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      FV({ tag: ValueTag.String, value: 'bad', stringId: 17 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: -100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FVSCHEDULE({ tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.String, value: 'bad', stringId: 18 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FVSCHEDULE({ tag: ValueTag.String, value: 'bad', stringId: 43 }, { tag: ValueTag.Number, value: 0.09 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DB(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DB(
        { tag: ValueTag.String, value: 'bad', stringId: 44 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DDB(
        { tag: ValueTag.String, value: 'bad', stringId: 45 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      VDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      VDB(
        { tag: ValueTag.String, value: 'bad', stringId: 46 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.String, value: 'bad', stringId: 47 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.String, value: 'bad', stringId: 48 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(SLN({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      SYD(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DISC(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      INTRATE(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1030 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      INTRATE(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 1030 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      RECEIVED(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PRICEDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        { tag: ValueTag.Number, value: 0.0525 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      YIELDDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PRICEMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 13,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      YIELDMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 3,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 8,
          },
        ),
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDFPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0785 },
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDFYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: -0.0575 },
        { tag: ValueTag.Number, value: 84.5 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 7,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0375 },
        { tag: ValueTag.Number, value: 0.0405 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 20,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 12 },
          {
            tag: ValueTag.Number,
            value: 24,
          },
        ),
        { tag: ValueTag.Number, value: -0.0375 },
        { tag: ValueTag.Number, value: 99.875 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      TBILLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.09 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      TBILLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      TBILLEQ(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      RECEIVED(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0.12 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports legacy statistical aliases', () => {
    const legacyNormsDist = getBuiltin('LEGACY.NORMSDIST')?.({ tag: ValueTag.Number, value: 0 })
    expect(legacyNormsDist).toMatchObject({ tag: ValueTag.Number })
    if (legacyNormsDist?.tag !== ValueTag.Number) {
      throw new Error('LEGACY.NORMSDIST should return a number')
    }
    expect(legacyNormsDist.value).toBeCloseTo(0.5, 8)
    expect(getBuiltin('LEGACY.NORMSINV')?.({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(getBuiltin('LEGACY.CHIDIST')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.exp(-0.5), 12),
    })
    expect(
      getBuiltin('SKEWP')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
  })

  it('covers coupon-date and periodic bond helpers with accelerated semantics', () => {
    const DATE = getBuiltin('DATE')!
    const COUPDAYBS = getBuiltin('COUPDAYBS')!
    const COUPDAYS = getBuiltin('COUPDAYS')!
    const COUPDAYSNC = getBuiltin('COUPDAYSNC')!
    const COUPNCD = getBuiltin('COUPNCD')!
    const COUPNUM = getBuiltin('COUPNUM')!
    const COUPPCD = getBuiltin('COUPPCD')!
    const PRICE = getBuiltin('PRICE')!
    const YIELD = getBuiltin('YIELD')!
    const DURATION = getBuiltin('DURATION')!
    const MDURATION = getBuiltin('MDURATION')!

    const couponSettlement = DATE(
      { tag: ValueTag.Number, value: 2007 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 25 },
    )
    const couponMaturity = DATE(
      { tag: ValueTag.Number, value: 2009 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 15 },
    )

    expect(COUPDAYBS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 70,
    })
    expect(COUPDAYS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 180,
    })
    expect(COUPDAYSNC(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 110,
    })
    expect(COUPNCD(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 39217,
    })
    expect(COUPNUM(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(COUPPCD(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 39036,
    })

    expect(
      PRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2017 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 0.065 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(94.63436162132213, 12),
    })
    expect(
      YIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2016 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 95.04287 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.065, 7),
    })
    expect(
      DURATION(
        DATE(
          { tag: ValueTag.Number, value: 2018 },
          { tag: ValueTag.Number, value: 7 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2048 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(10.919145281591925, 12),
    })
    expect(
      MDURATION(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2016 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5.735669813918838, 12),
    })

    expect(COUPDAYBS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PRICE(
        couponSettlement,
        couponMaturity,
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: -0.01 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      YIELD(
        couponSettlement,
        couponMaturity,
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      DURATION(
        couponMaturity,
        couponSettlement,
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers remaining complex and distribution error branches', () => {
    const IMPRODUCT = getBuiltin('IMPRODUCT')!
    const IMSUB = getBuiltin('IMSUB')!
    const IMTAN = getBuiltin('IMTAN')!
    const IMSEC = getBuiltin('IMSEC')!
    const IMCSC = getBuiltin('IMCSC')!
    const IMCOT = getBuiltin('IMCOT')!
    const NORMDIST = getBuiltin('NORMDIST')!
    const LOGNORM_DOT_DIST = getBuiltin('LOGNORM.DIST')!
    const LOGNORMDIST = getBuiltin('LOGNORMDIST')!
    const NPV = getBuiltin('NPV')!

    expect(IMPRODUCT()).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMSUB({ tag: ValueTag.String, value: 'bad', stringId: 89 }, { tag: ValueTag.String, value: '1+i', stringId: 94 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMTAN({ tag: ValueTag.String, value: 'bad', stringId: 90 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMSEC({ tag: ValueTag.String, value: 'bad', stringId: 92 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMCSC({ tag: ValueTag.String, value: 'bad', stringId: 93 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMCOT({ tag: ValueTag.String, value: 'bad', stringId: 91 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      LOGNORMDIST(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      LOGNORM_DOT_DIST(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NPV({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.String, value: 'bad', stringId: 95 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('registers protocol-declared placeholder builtins as blocked', () => {
    for (const name of placeholderBuiltinNames) {
      expect(getBuiltin(name)?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    }

    for (const name of protocolPlaceholderBuiltinNames) {
      expect(getBuiltinId(name.toLowerCase())).toBeDefined()
    }

    expect(getBuiltinId('sin')).toBe(BuiltinId.Sin)
    expect(getBuiltinId('weeknum')).toBe(BuiltinId.Weeknum)
    expect(getBuiltinId('rept')).toBe(BuiltinId.Rept)
    expect(getBuiltinId('filter')).toBe(BuiltinId.Filter)
    expect(getBuiltinId('let')).toBeUndefined()
    expect(getBuiltinId('textjoin')).toBe(BuiltinId.Textjoin)
  })

  it('routes provider-backed formulas through external adapters and blocks when none are installed', () => {
    const TRANSLATE = getBuiltin('TRANSLATE')!
    const HYPERLINK = getBuiltin('HYPERLINK')!
    const DDE = getBuiltin('DDE')!
    const INFO = getBuiltin('INFO')!
    const REGISTER_ID = getBuiltin('REGISTER.ID')!
    const FILTERXML = getLookupBuiltin('FILTERXML')!
    const STOCKHISTORY = getLookupBuiltin('STOCKHISTORY')!

    const hello = { tag: ValueTag.String, value: 'hello', stringId: 1 } as const
    const sourceLang = { tag: ValueTag.String, value: 'en', stringId: 2 } as const
    const targetLang = { tag: ValueTag.String, value: 'es', stringId: 3 } as const

    expect(placeholderBuiltinNames).not.toContain('TRANSLATE')
    expect(placeholderBuiltinNames).not.toContain('FILTERXML')
    expect(placeholderBuiltinNames).not.toContain('STOCKHISTORY')

    expect(TRANSLATE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(HYPERLINK(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(DDE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(INFO(hello)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    expect(REGISTER_ID(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(FILTERXML(hello, sourceLang)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    expect(STOCKHISTORY(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })

    const translateImpl = vi.fn(() => ({
      tag: ValueTag.String,
      value: 'hola',
      stringId: 0,
    }))
    const hyperlinkImpl = vi.fn(() => ({
      tag: ValueTag.String,
      value: 'linked',
      stringId: 0,
    }))
    const ddeImpl = vi.fn(() => ({ tag: ValueTag.Number, value: 42 }))
    const infoImpl = vi.fn(() => ({ tag: ValueTag.String, value: 'mac', stringId: 0 }))
    const registerIdImpl = vi.fn(() => ({ tag: ValueTag.Number, value: 17 }))
    const filterXmlImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'one', stringId: 0 },
        { tag: ValueTag.String, value: 'two', stringId: 0 },
      ],
    }))
    const stockHistoryImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 11 },
      ],
    }))

    installExternalFunctionAdapter({
      surface: 'web',
      resolveFunction(name) {
        if (name === 'TRANSLATE') {
          return { kind: 'scalar', implementation: translateImpl }
        }
        if (name === 'HYPERLINK') {
          return { kind: 'scalar', implementation: hyperlinkImpl }
        }
        if (name === 'INFO') {
          return { kind: 'scalar', implementation: infoImpl }
        }
        if (name === 'FILTERXML') {
          return { kind: 'lookup', implementation: filterXmlImpl }
        }
        if (name === 'STOCKHISTORY') {
          return { kind: 'lookup', implementation: stockHistoryImpl }
        }
        return undefined
      },
    })
    installExternalFunctionAdapter({
      surface: 'external-data',
      resolveFunction(name) {
        if (name === 'DDE') {
          return { kind: 'scalar', implementation: ddeImpl }
        }
        if (name === 'REGISTER.ID') {
          return { kind: 'scalar', implementation: registerIdImpl }
        }
        return undefined
      },
    })

    expect(TRANSLATE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.String,
      value: 'hola',
      stringId: 0,
    })
    expect(HYPERLINK(hello, sourceLang)).toEqual({
      tag: ValueTag.String,
      value: 'linked',
      stringId: 0,
    })
    expect(DDE(hello, sourceLang, targetLang)).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(INFO(hello)).toEqual({
      tag: ValueTag.String,
      value: 'mac',
      stringId: 0,
    })
    expect(REGISTER_ID(hello, sourceLang)).toEqual({ tag: ValueTag.Number, value: 17 })
    expect(FILTERXML(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'one', stringId: 0 },
        { tag: ValueTag.String, value: 'two', stringId: 0 },
      ],
    })
    expect(STOCKHISTORY(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 11 },
      ],
    })

    expect(translateImpl).toHaveBeenCalledWith(hello, sourceLang, targetLang)
    expect(filterXmlImpl).toHaveBeenCalledWith(hello, sourceLang)
    expect(stockHistoryImpl).toHaveBeenCalledWith(hello, sourceLang)
  })
})
