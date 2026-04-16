import { describe, expect, it } from 'vitest'
import { builtinJsSpecialNames, builtinWasmEnabledNames, getBuiltinCapability } from '../builtin-capabilities.js'

describe('builtin capabilities', () => {
  it('tracks native production coverage for the current promoted builtin set', () => {
    expect(builtinWasmEnabledNames.has('SUM')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUNTBLANK')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUNTIFS')).toBe(true)
    expect(builtinWasmEnabledNames.has('IFS')).toBe(true)
    expect(builtinWasmEnabledNames.has('SWITCH')).toBe(true)
    expect(builtinWasmEnabledNames.has('XOR')).toBe(true)
    expect(builtinWasmEnabledNames.has('DATEDIF')).toBe(true)
    expect(builtinWasmEnabledNames.has('DAYS360')).toBe(true)
    expect(builtinWasmEnabledNames.has('FVSCHEDULE')).toBe(true)
    expect(builtinWasmEnabledNames.has('FV')).toBe(true)
    expect(builtinWasmEnabledNames.has('PV')).toBe(true)
    expect(builtinWasmEnabledNames.has('PMT')).toBe(true)
    expect(builtinWasmEnabledNames.has('NPER')).toBe(true)
    expect(builtinWasmEnabledNames.has('NPV')).toBe(true)
    expect(builtinWasmEnabledNames.has('RATE')).toBe(true)
    expect(builtinWasmEnabledNames.has('IPMT')).toBe(true)
    expect(builtinWasmEnabledNames.has('PPMT')).toBe(true)
    expect(builtinWasmEnabledNames.has('ISPMT')).toBe(true)
    expect(builtinWasmEnabledNames.has('CUMIPMT')).toBe(true)
    expect(builtinWasmEnabledNames.has('CUMPRINC')).toBe(true)
    expect(builtinWasmEnabledNames.has('DISC')).toBe(true)
    expect(builtinWasmEnabledNames.has('INTRATE')).toBe(true)
    expect(builtinWasmEnabledNames.has('RECEIVED')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPDAYBS')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPDAYS')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPDAYSNC')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPNCD')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPNUM')).toBe(true)
    expect(builtinWasmEnabledNames.has('COUPPCD')).toBe(true)
    expect(builtinWasmEnabledNames.has('PRICEDISC')).toBe(true)
    expect(builtinWasmEnabledNames.has('YIELDDISC')).toBe(true)
    expect(builtinWasmEnabledNames.has('PRICEMAT')).toBe(true)
    expect(builtinWasmEnabledNames.has('YIELDMAT')).toBe(true)
    expect(builtinWasmEnabledNames.has('ODDLPRICE')).toBe(true)
    expect(builtinWasmEnabledNames.has('ODDLYIELD')).toBe(true)
    expect(builtinWasmEnabledNames.has('PRICE')).toBe(true)
    expect(builtinWasmEnabledNames.has('YIELD')).toBe(true)
    expect(builtinWasmEnabledNames.has('DURATION')).toBe(true)
    expect(builtinWasmEnabledNames.has('MDURATION')).toBe(true)
    expect(builtinWasmEnabledNames.has('TBILLPRICE')).toBe(true)
    expect(builtinWasmEnabledNames.has('TBILLYIELD')).toBe(true)
    expect(builtinWasmEnabledNames.has('TBILLEQ')).toBe(true)
    expect(builtinWasmEnabledNames.has('EFFECT')).toBe(true)
    expect(builtinWasmEnabledNames.has('NOMINAL')).toBe(true)
    expect(builtinWasmEnabledNames.has('PDURATION')).toBe(true)
    expect(builtinWasmEnabledNames.has('RRI')).toBe(true)
    expect(builtinWasmEnabledNames.has('IRR')).toBe(true)
    expect(builtinWasmEnabledNames.has('MIRR')).toBe(true)
    expect(builtinWasmEnabledNames.has('XNPV')).toBe(true)
    expect(builtinWasmEnabledNames.has('XIRR')).toBe(true)
    expect(builtinWasmEnabledNames.has('CORREL')).toBe(true)
    expect(builtinWasmEnabledNames.has('COVAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('PEARSON')).toBe(true)
    expect(builtinWasmEnabledNames.has('COVARIANCE.P')).toBe(true)
    expect(builtinWasmEnabledNames.has('COVARIANCE.S')).toBe(true)
    expect(builtinWasmEnabledNames.has('STANDARDIZE')).toBe(true)
    expect(builtinWasmEnabledNames.has('CONFIDENCE.NORM')).toBe(true)
    expect(builtinWasmEnabledNames.has('MODE')).toBe(true)
    expect(builtinWasmEnabledNames.has('MODE.SNGL')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEV')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEV.P')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEV.S')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEVA')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEVP')).toBe(true)
    expect(builtinWasmEnabledNames.has('STDEVPA')).toBe(true)
    expect(builtinWasmEnabledNames.has('VAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('VAR.P')).toBe(true)
    expect(builtinWasmEnabledNames.has('VAR.S')).toBe(true)
    expect(builtinWasmEnabledNames.has('VARA')).toBe(true)
    expect(builtinWasmEnabledNames.has('VARP')).toBe(true)
    expect(builtinWasmEnabledNames.has('VARPA')).toBe(true)
    expect(builtinWasmEnabledNames.has('SKEW')).toBe(true)
    expect(builtinWasmEnabledNames.has('SKEW.P')).toBe(true)
    expect(builtinWasmEnabledNames.has('KURT')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORMDIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORM.DIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORMINV')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORM.INV')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORMSDIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORM.S.DIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORMSINV')).toBe(true)
    expect(builtinWasmEnabledNames.has('NORM.S.INV')).toBe(true)
    expect(builtinWasmEnabledNames.has('LOGINV')).toBe(true)
    expect(builtinWasmEnabledNames.has('LOGNORMDIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('LOGNORM.DIST')).toBe(true)
    expect(builtinWasmEnabledNames.has('LOGNORM.INV')).toBe(true)
    expect(builtinWasmEnabledNames.has('FORECAST')).toBe(true)
    expect(builtinWasmEnabledNames.has('GROWTH')).toBe(true)
    expect(builtinWasmEnabledNames.has('INTERCEPT')).toBe(true)
    expect(builtinWasmEnabledNames.has('LINEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('LOGEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('MEDIAN')).toBe(true)
    expect(builtinWasmEnabledNames.has('MODE.MULT')).toBe(true)
    expect(builtinWasmEnabledNames.has('FREQUENCY')).toBe(true)
    expect(builtinWasmEnabledNames.has('SMALL')).toBe(true)
    expect(builtinWasmEnabledNames.has('LARGE')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTILE')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTILE.INC')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTILE.EXC')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTRANK')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTRANK.INC')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERCENTRANK.EXC')).toBe(true)
    expect(builtinWasmEnabledNames.has('PROB')).toBe(true)
    expect(builtinWasmEnabledNames.has('QUARTILE')).toBe(true)
    expect(builtinWasmEnabledNames.has('QUARTILE.INC')).toBe(true)
    expect(builtinWasmEnabledNames.has('QUARTILE.EXC')).toBe(true)
    expect(builtinWasmEnabledNames.has('RANK')).toBe(true)
    expect(builtinWasmEnabledNames.has('RANK.EQ')).toBe(true)
    expect(builtinWasmEnabledNames.has('RANK.AVG')).toBe(true)
    expect(builtinWasmEnabledNames.has('RSQ')).toBe(true)
    expect(builtinWasmEnabledNames.has('SLOPE')).toBe(true)
    expect(builtinWasmEnabledNames.has('STEYX')).toBe(true)
    expect(builtinWasmEnabledNames.has('TREND')).toBe(true)
    expect(builtinWasmEnabledNames.has('TRIMMEAN')).toBe(true)
    expect(builtinWasmEnabledNames.has('ISOWEEKNUM')).toBe(true)
    expect(builtinWasmEnabledNames.has('TIMEVALUE')).toBe(true)
    expect(builtinWasmEnabledNames.has('YEARFRAC')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHISQ.INV.RT')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHIINV')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHISQ.TEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHITEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('F.TEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('Z.TEST')).toBe(true)
    expect(builtinWasmEnabledNames.has('WORKDAY.INTL')).toBe(true)
    expect(builtinWasmEnabledNames.has('NETWORKDAYS.INTL')).toBe(true)
    expect(builtinWasmEnabledNames.has('T')).toBe(true)
    expect(builtinWasmEnabledNames.has('N')).toBe(true)
    expect(builtinWasmEnabledNames.has('TYPE')).toBe(true)
    expect(builtinWasmEnabledNames.has('DELTA')).toBe(true)
    expect(builtinWasmEnabledNames.has('GESTEP')).toBe(true)
    expect(builtinWasmEnabledNames.has('GAUSS')).toBe(true)
    expect(builtinWasmEnabledNames.has('PHI')).toBe(true)
    expect(builtinWasmEnabledNames.has('NUMBERVALUE')).toBe(true)
    expect(builtinWasmEnabledNames.has('TEXT')).toBe(true)
    expect(builtinWasmEnabledNames.has('PHONETIC')).toBe(true)
    expect(builtinWasmEnabledNames.has('VALUETOTEXT')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('CODE')).toBe(true)
    expect(builtinWasmEnabledNames.has('ASC')).toBe(true)
    expect(builtinWasmEnabledNames.has('JIS')).toBe(true)
    expect(builtinWasmEnabledNames.has('DBCS')).toBe(true)
    expect(builtinWasmEnabledNames.has('BAHTTEXT')).toBe(true)
    expect(builtinWasmEnabledNames.has('ODDFPRICE')).toBe(true)
    expect(builtinWasmEnabledNames.has('ODDFYIELD')).toBe(true)
    expect(builtinWasmEnabledNames.has('DAVERAGE')).toBe(true)
    expect(builtinWasmEnabledNames.has('DCOUNT')).toBe(true)
    expect(builtinWasmEnabledNames.has('DCOUNTA')).toBe(true)
    expect(builtinWasmEnabledNames.has('DGET')).toBe(true)
    expect(builtinWasmEnabledNames.has('DMAX')).toBe(true)
    expect(builtinWasmEnabledNames.has('DMIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('DPRODUCT')).toBe(true)
    expect(builtinWasmEnabledNames.has('DSTDEV')).toBe(true)
    expect(builtinWasmEnabledNames.has('DSTDEVP')).toBe(true)
    expect(builtinWasmEnabledNames.has('DSUM')).toBe(true)
    expect(builtinWasmEnabledNames.has('DVAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('DVARP')).toBe(true)
    expect(builtinWasmEnabledNames.has('UNICODE')).toBe(true)
    expect(builtinWasmEnabledNames.has('UNICHAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('CLEAN')).toBe(true)
    expect(builtinWasmEnabledNames.has('CHOOSE')).toBe(true)
    expect(builtinWasmEnabledNames.has('TEXTBEFORE')).toBe(true)
    expect(builtinWasmEnabledNames.has('TEXTAFTER')).toBe(true)
    expect(builtinWasmEnabledNames.has('TEXTJOIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('TEXTSPLIT')).toBe(true)
    expect(builtinWasmEnabledNames.has('SEQUENCE')).toBe(true)
    expect(builtinWasmEnabledNames.has('LENB')).toBe(true)
    expect(builtinWasmEnabledNames.has('LEFTB')).toBe(true)
    expect(builtinWasmEnabledNames.has('MIDB')).toBe(true)
    expect(builtinWasmEnabledNames.has('RIGHTB')).toBe(true)
    expect(builtinWasmEnabledNames.has('FINDB')).toBe(true)
    expect(builtinWasmEnabledNames.has('SEARCHB')).toBe(true)
    expect(builtinWasmEnabledNames.has('REPLACEB')).toBe(true)
    expect(builtinWasmEnabledNames.has('ADDRESS')).toBe(true)
    expect(builtinWasmEnabledNames.has('DOLLAR')).toBe(true)
    expect(builtinWasmEnabledNames.has('DOLLARDE')).toBe(true)
    expect(builtinWasmEnabledNames.has('DOLLARFR')).toBe(true)
    expect(builtinWasmEnabledNames.has('BASE')).toBe(true)
    expect(builtinWasmEnabledNames.has('DECIMAL')).toBe(true)
    expect(builtinWasmEnabledNames.has('BIN2DEC')).toBe(true)
    expect(builtinWasmEnabledNames.has('BIN2HEX')).toBe(true)
    expect(builtinWasmEnabledNames.has('BIN2OCT')).toBe(true)
    expect(builtinWasmEnabledNames.has('DEC2BIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('DEC2HEX')).toBe(true)
    expect(builtinWasmEnabledNames.has('DEC2OCT')).toBe(true)
    expect(builtinWasmEnabledNames.has('HEX2BIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('HEX2DEC')).toBe(true)
    expect(builtinWasmEnabledNames.has('HEX2OCT')).toBe(true)
    expect(builtinWasmEnabledNames.has('OCT2BIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('OCT2DEC')).toBe(true)
    expect(builtinWasmEnabledNames.has('OCT2HEX')).toBe(true)
    expect(builtinWasmEnabledNames.has('BITAND')).toBe(true)
    expect(builtinWasmEnabledNames.has('BITOR')).toBe(true)
    expect(builtinWasmEnabledNames.has('BITXOR')).toBe(true)
    expect(builtinWasmEnabledNames.has('BITLSHIFT')).toBe(true)
    expect(builtinWasmEnabledNames.has('BITRSHIFT')).toBe(true)
    expect(builtinWasmEnabledNames.has('FLOOR.MATH')).toBe(true)
    expect(builtinWasmEnabledNames.has('FLOOR.PRECISE')).toBe(true)
    expect(builtinWasmEnabledNames.has('CEILING.MATH')).toBe(true)
    expect(builtinWasmEnabledNames.has('CEILING.PRECISE')).toBe(true)
    expect(builtinWasmEnabledNames.has('ISO.CEILING')).toBe(true)
    expect(builtinWasmEnabledNames.has('TRUNC')).toBe(true)
    expect(builtinWasmEnabledNames.has('MROUND')).toBe(true)
    expect(builtinWasmEnabledNames.has('SERIESSUM')).toBe(true)
    expect(builtinWasmEnabledNames.has('SQRTPI')).toBe(true)
    expect(builtinWasmEnabledNames.has('CONVERT')).toBe(true)
    expect(builtinWasmEnabledNames.has('EUROCONVERT')).toBe(true)
    expect(builtinWasmEnabledNames.has('SINH')).toBe(true)
    expect(builtinWasmEnabledNames.has('COSH')).toBe(true)
    expect(builtinWasmEnabledNames.has('TANH')).toBe(true)
    expect(builtinWasmEnabledNames.has('ASINH')).toBe(true)
    expect(builtinWasmEnabledNames.has('ACOSH')).toBe(true)
    expect(builtinWasmEnabledNames.has('ATANH')).toBe(true)
    expect(builtinWasmEnabledNames.has('ACOT')).toBe(true)
    expect(builtinWasmEnabledNames.has('ACOTH')).toBe(true)
    expect(builtinWasmEnabledNames.has('COT')).toBe(true)
    expect(builtinWasmEnabledNames.has('COTH')).toBe(true)
    expect(builtinWasmEnabledNames.has('CSC')).toBe(true)
    expect(builtinWasmEnabledNames.has('CSCH')).toBe(true)
    expect(builtinWasmEnabledNames.has('SEC')).toBe(true)
    expect(builtinWasmEnabledNames.has('SECH')).toBe(true)
    expect(builtinWasmEnabledNames.has('SIGN')).toBe(true)
    expect(builtinWasmEnabledNames.has('EVEN')).toBe(true)
    expect(builtinWasmEnabledNames.has('ODD')).toBe(true)
    expect(builtinWasmEnabledNames.has('FACT')).toBe(true)
    expect(builtinWasmEnabledNames.has('FACTDOUBLE')).toBe(true)
    expect(builtinWasmEnabledNames.has('COMBIN')).toBe(true)
    expect(builtinWasmEnabledNames.has('COMBINA')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERMUT')).toBe(true)
    expect(builtinWasmEnabledNames.has('PERMUTATIONA')).toBe(true)
    expect(builtinWasmEnabledNames.has('GCD')).toBe(true)
    expect(builtinWasmEnabledNames.has('LCM')).toBe(true)
    expect(builtinWasmEnabledNames.has('PRODUCT')).toBe(true)
    expect(builtinWasmEnabledNames.has('QUOTIENT')).toBe(true)
    expect(builtinWasmEnabledNames.has('GEOMEAN')).toBe(true)
    expect(builtinWasmEnabledNames.has('HARMEAN')).toBe(true)
    expect(builtinWasmEnabledNames.has('SUMSQ')).toBe(true)
    expect(builtinWasmEnabledNames.has('BESSELI')).toBe(true)
    expect(builtinWasmEnabledNames.has('BESSELJ')).toBe(true)
    expect(builtinWasmEnabledNames.has('BESSELK')).toBe(true)
    expect(builtinWasmEnabledNames.has('BESSELY')).toBe(true)
    expect(builtinWasmEnabledNames.has('BETA.INV')).toBe(true)
    expect(builtinWasmEnabledNames.has('F.DIST.RT')).toBe(true)
    expect(builtinWasmEnabledNames.has('LET')).toBe(false)
  })

  it('tracks JS-only higher-order builtins separately from native coverage', () => {
    expect(builtinJsSpecialNames.has('LET')).toBe(true)
    expect(builtinJsSpecialNames.has('LAMBDA')).toBe(true)
    expect(builtinJsSpecialNames.has('MAP')).toBe(true)
    expect(builtinJsSpecialNames.has('INDIRECT')).toBe(true)
    expect(builtinJsSpecialNames.has('GETPIVOTDATA')).toBe(true)
    expect(builtinJsSpecialNames.has('GROUPBY')).toBe(true)
    expect(builtinJsSpecialNames.has('PIVOTBY')).toBe(true)
    expect(builtinJsSpecialNames.has('MULTIPLE.OPERATIONS')).toBe(true)
    expect(builtinJsSpecialNames.has('TEXTSPLIT')).toBe(false)
    expect(builtinJsSpecialNames.has('EXPAND')).toBe(false)
    expect(builtinJsSpecialNames.has('TRIMRANGE')).toBe(false)
  })

  it('exposes array-runtime backlog metadata for non-native families', () => {
    expect(getBuiltinCapability('FILTER')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('COUNTBLANK')).toMatchObject({
      category: 'aggregation',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('MODE.MULT')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BYROW')).toMatchObject({
      category: 'lambda',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('EXPAND')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('CORREL')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COVARIANCE.S')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('FORECAST')).toMatchObject({
      category: 'lookup-reference',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('GROWTH')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('LINEST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('LOGEST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('MEDIAN')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PERCENTILE.EXC')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('QUARTILE.INC')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('RANK')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TREND')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('RANK.AVG')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('STEYX')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PROB')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TRIMMEAN')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TRIMRANGE')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('TEXTSPLIT')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'implemented',
      wasmStatus: 'production',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('DATEDIF')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('RATE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CUMIPMT')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CUMPRINC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DAYS360')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('WORKDAY.INTL')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('NETWORKDAYS.INTL')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DISC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('INTRATE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('RECEIVED')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPDAYBS')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPDAYS')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPDAYSNC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPNCD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPNUM')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('COUPPCD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PRICEDISC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('YIELDDISC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PRICEMAT')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('YIELDMAT')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ODDLPRICE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ODDLYIELD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PRICE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('YIELD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DURATION')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('MDURATION')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TBILLPRICE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TBILLYIELD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TBILLEQ')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('IRR')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('MIRR')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('XNPV')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('XIRR')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ISOWEEKNUM')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TIMEVALUE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('YEARFRAC')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('NUMBERVALUE')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TEXT')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('VALUETOTEXT')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CHAR')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CODE')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ASC')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('JIS')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DBCS')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BAHTTEXT')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ODDFPRICE')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ODDFYIELD')).toMatchObject({
      category: 'date-time',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DAVERAGE')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DCOUNT')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DGET')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DSTDEV')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DVARP')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('UNICODE')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('UNICHAR')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CLEAN')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CHOOSE')).toMatchObject({
      category: 'lookup-reference',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TEXTBEFORE')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TEXTAFTER')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('TEXTJOIN')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('LENB')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('SEARCHB')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('REPLACEB')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ADDRESS')).toMatchObject({
      category: 'lookup-reference',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DOLLAR')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DOLLARDE')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DOLLARFR')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BASE')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DECIMAL')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CONVERT')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('EUROCONVERT')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('ACOSH')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PRODUCT')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('SUMSQ')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BIN2DEC')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('DEC2HEX')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('HEX2DEC')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('OCT2HEX')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BITAND')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BITRSHIFT')).toMatchObject({
      category: 'math',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CHISQ.INV')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('CHISQ.TEST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('F.TEST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('Z.TEST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('BETA.DIST')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('F.INV.RT')).toMatchObject({
      category: 'statistical',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('PHONETIC')).toMatchObject({
      category: 'text',
      jsStatus: 'implemented',
      wasmStatus: 'production',
    })
    expect(getBuiltinCapability('GETPIVOTDATA')).toMatchObject({
      category: 'lookup-reference',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
    })
    expect(getBuiltinCapability('GROUPBY')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('PIVOTBY')).toMatchObject({
      category: 'dynamic-array',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
      needsArrayRuntime: true,
    })
    expect(getBuiltinCapability('MULTIPLE.OPERATIONS')).toMatchObject({
      category: 'lookup-reference',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
    })
  })
})
