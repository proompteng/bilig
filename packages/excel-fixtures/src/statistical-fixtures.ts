import type {
  ExcelExpectedValue,
  ExcelFixtureCase,
  ExcelFixtureExpectedOutput,
  ExcelFixtureFamily,
  ExcelFixtureInputCell,
} from './index.js'

const excelFixtureIdPattern = /^[a-z][a-z0-9-]*:[a-z0-9-]+$/

function createExcelFixtureId(family: ExcelFixtureFamily, slug: string): string {
  const normalizedSlug = slug.trim().toLowerCase()
  const id = `${family}:${normalizedSlug}`
  if (!excelFixtureIdPattern.test(id)) {
    throw new Error(`Invalid Excel fixture id: ${id}`)
  }
  return id
}

function numberExpected(value: number): ExcelExpectedValue {
  return { kind: 'number', value }
}

function input(address: string, value: ExcelFixtureInputCell['input'], note?: string): ExcelFixtureInputCell {
  return note === undefined ? { address, input: value } : { address, input: value, note }
}

function output(address: string, expected: ExcelFixtureExpectedOutput['expected'], note?: string): ExcelFixtureExpectedOutput {
  return note === undefined ? { address, expected } : { address, expected, note }
}

function fixture(
  slug: string,
  title: string,
  formula: string,
  inputs: ExcelFixtureInputCell[],
  outputs: ExcelFixtureExpectedOutput[],
  notes?: string,
): ExcelFixtureCase {
  const base = {
    id: createExcelFixtureId('statistical', slug),
    family: 'statistical' as const,
    title,
    formula,
    inputs,
    outputs,
    sheetName: 'Sheet1',
  }
  return notes === undefined ? base : { ...base, notes }
}

function regressionInputs(): ExcelFixtureInputCell[] {
  return [input('A1', 5), input('A2', 8), input('A3', 11), input('B1', 1), input('B2', 2), input('B3', 3)]
}

function databaseFixtureInputs(): ExcelFixtureInputCell[] {
  return [
    input('A1', 'Age'),
    input('B1', 'Height'),
    input('C1', 'Yield'),
    input('A2', 10),
    input('B2', 100),
    input('C2', 5),
    input('A3', 12),
    input('B3', 110),
    input('C3', 7),
    input('A4', 12),
    input('B4', 120),
    input('C4', 9),
    input('A5', 15),
    input('B5', 130),
    input('C5', 11),
    input('E1', 'Age'),
    input('E2', 12),
    input('F1', 'Age'),
    input('F2', 15),
  ]
}

export const canonicalStatisticalFixtures: readonly ExcelFixtureCase[] = [
  fixture(
    'standardize-basic',
    'STANDARDIZE normalizes a value by mean and standard deviation',
    '=STANDARDIZE(1,0,1)',
    [],
    [output('C1', numberExpected(1))],
  ),
  fixture(
    'confidence-norm-basic',
    'CONFIDENCE.NORM returns the normal confidence interval half-width',
    '=CONFIDENCE.NORM(0.05,1,100)',
    [],
    [output('C1', numberExpected(0.1959963984540054))],
  ),
  fixture(
    'mode-basic',
    'MODE returns the most frequent numeric value',
    '=MODE(A1:A6)',
    [input('A1', 1), input('A2', 2), input('A3', 2), input('A4', 3), input('A5', 3), input('A6', 3)],
    [output('C1', numberExpected(3))],
  ),
  fixture(
    'mode-sngl-basic',
    'MODE.SNGL returns the single most frequent numeric value',
    '=MODE.SNGL(A1:A6)',
    [input('A1', 1), input('A2', 2), input('A3', 2), input('A4', 3), input('A5', 3), input('A6', 3)],
    [output('C1', numberExpected(3))],
  ),
  fixture(
    'stdev-basic',
    'STDEV returns the sample standard deviation of a numeric range',
    '=STDEV(A1:A4)',
    [input('A1', 1), input('A2', 2), input('A3', 3), input('A4', 4)],
    [output('C1', numberExpected(Math.sqrt(5 / 3)))],
  ),
  fixture(
    'stdeva-basic',
    'STDEVA includes booleans and text-as-zero in direct arguments',
    '=STDEVA(2,TRUE(),"skip")',
    [],
    [output('C1', numberExpected(1))],
  ),
  fixture(
    'var-basic',
    'VAR returns the sample variance of a numeric range',
    '=VAR(A1:A4)',
    [input('A1', 1), input('A2', 2), input('A3', 3), input('A4', 4)],
    [output('C1', numberExpected(5 / 3))],
  ),
  fixture(
    'vara-basic',
    'VARA includes booleans and text-as-zero in direct arguments',
    '=VARA(2,TRUE(),"skip")',
    [],
    [output('C1', numberExpected(1))],
  ),
  fixture(
    'skew-basic',
    'SKEW returns zero for a symmetric sample',
    '=SKEW(A1:A5)',
    [input('A1', 2), input('A2', 3), input('A3', 4), input('A4', 5), input('A5', 6)],
    [output('C1', numberExpected(0))],
  ),
  fixture(
    'kurt-basic',
    'KURT returns the excess kurtosis of the sample',
    '=KURT(A1:A5)',
    [input('A1', 1), input('A2', 2), input('A3', 3), input('A4', 4), input('A5', 5)],
    [output('C1', numberExpected(-1.2))],
  ),
  fixture(
    'normdist-basic',
    'NORMDIST returns the cumulative normal distribution',
    '=NORMDIST(1,0,1,TRUE)',
    [],
    [output('C1', numberExpected(0.8413447460685429))],
  ),
  fixture(
    'norminv-basic',
    'NORMINV inverts the normal cumulative distribution',
    '=NORMINV(0.8413447460685429,0,1)',
    [],
    [output('C1', numberExpected(1))],
  ),
  fixture(
    'normsdist-basic',
    'NORMSDIST returns the standard normal cumulative distribution',
    '=NORMSDIST(1)',
    [],
    [output('C1', numberExpected(0.8413447460685429))],
  ),
  fixture(
    'normsinv-basic',
    'NORMSINV returns the inverse standard normal quantile',
    '=NORMSINV(0.001)',
    [],
    [output('C1', numberExpected(-3.090232306167813))],
  ),
  fixture('loginv-basic', 'LOGINV exponentiates the inverse normal quantile', '=LOGINV(0.5,0,1)', [], [output('C1', numberExpected(1))]),
  fixture(
    'lognormdist-basic',
    'LOGNORMDIST returns the log-normal cumulative distribution',
    '=LOGNORMDIST(1,0,1)',
    [],
    [output('C1', numberExpected(0.5))],
  ),
  fixture('correl-basic', 'CORREL returns perfect linear correlation', '=CORREL(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(1)),
  ]),
  fixture('covar-basic', 'COVAR returns population covariance for paired ranges', '=COVAR(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(2)),
  ]),
  fixture('covariance-p-basic', 'COVARIANCE.P returns population covariance', '=COVARIANCE.P(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(2)),
  ]),
  fixture('covariance-s-basic', 'COVARIANCE.S returns sample covariance', '=COVARIANCE.S(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(3)),
  ]),
  fixture('pearson-basic', 'PEARSON returns perfect linear correlation', '=PEARSON(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(1)),
  ]),
  fixture('intercept-basic', 'INTERCEPT returns the fitted regression intercept', '=INTERCEPT(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(2)),
  ]),
  fixture('slope-basic', 'SLOPE returns the fitted regression slope', '=SLOPE(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(3)),
  ]),
  fixture('rsq-basic', 'RSQ squares the paired correlation coefficient', '=RSQ(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(1)),
  ]),
  fixture('steyx-basic', 'STEYX returns zero for an exact linear fit', '=STEYX(A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(0)),
  ]),
  fixture(
    'rank-basic',
    'RANK returns the top-tied ordinal position',
    '=RANK(20,A1:A4)',
    [input('A1', 10), input('A2', 20), input('A3', 20), input('A4', 30)],
    [output('C1', numberExpected(2))],
  ),
  fixture(
    'rank-eq-basic',
    'RANK.EQ returns the same ordinal rank as RANK',
    '=RANK.EQ(20,A1:A4)',
    [input('A1', 10), input('A2', 20), input('A3', 20), input('A4', 30)],
    [output('C1', numberExpected(2))],
  ),
  fixture(
    'rank-avg-basic',
    'RANK.AVG averages tied ordinal positions',
    '=RANK.AVG(20,A1:A4)',
    [input('A1', 10), input('A2', 20), input('A3', 20), input('A4', 30)],
    [output('C1', numberExpected(2.5))],
  ),
  fixture(
    'median-basic',
    'MEDIAN returns the middle value of a sorted data set',
    '=MEDIAN(A1:A8)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(7.5))],
  ),
  fixture(
    'small-basic',
    'SMALL returns the k-th smallest numeric value',
    '=SMALL(A1:A8,3)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(4))],
  ),
  fixture(
    'large-basic',
    'LARGE returns the k-th largest numeric value',
    '=LARGE(A1:A8,2)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(10))],
  ),
  fixture(
    'percentile-basic',
    'PERCENTILE aliases the inclusive percentile interpolation',
    '=PERCENTILE(A1:A8,0.25)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(3.5))],
  ),
  fixture(
    'percentile-inc-basic',
    'PERCENTILE.INC interpolates inclusive percentiles',
    '=PERCENTILE.INC(A1:A8,0.25)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(3.5))],
  ),
  fixture(
    'percentile-exc-basic',
    'PERCENTILE.EXC interpolates exclusive percentiles',
    '=PERCENTILE.EXC(A1:A8,0.25)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(2.5))],
  ),
  fixture(
    'percentrank-basic',
    'PERCENTRANK aliases the inclusive percentage-rank interpolation',
    '=PERCENTRANK(A1:A8,8)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(0.571))],
  ),
  fixture(
    'percentrank-inc-basic',
    'PERCENTRANK.INC returns an inclusive percentage rank',
    '=PERCENTRANK.INC(A1:A8,8)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(0.571))],
  ),
  fixture(
    'percentrank-exc-basic',
    'PERCENTRANK.EXC returns an exclusive percentage rank',
    '=PERCENTRANK.EXC(A1:A8,8)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(0.555))],
  ),
  fixture(
    'quartile-basic',
    'QUARTILE aliases the inclusive quartile interpolation',
    '=QUARTILE(A1:A8,1)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(3.5))],
  ),
  fixture(
    'quartile-inc-basic',
    'QUARTILE.INC returns the inclusive first quartile',
    '=QUARTILE.INC(A1:A8,1)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(3.5))],
  ),
  fixture(
    'quartile-exc-basic',
    'QUARTILE.EXC returns the exclusive first quartile',
    '=QUARTILE.EXC(A1:A8,1)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(2.5))],
  ),
  fixture(
    'mode-mult-basic',
    'MODE.MULT spills each modal value in ascending order',
    '=MODE.MULT(A1:A6)',
    [input('A1', 1), input('A2', 2), input('A3', 2), input('A4', 3), input('A5', 3), input('A6', 4)],
    [output('C1', numberExpected(2)), output('C2', numberExpected(3))],
  ),
  fixture(
    'frequency-basic',
    'FREQUENCY spills histogram bucket counts vertically',
    '=FREQUENCY(A1:A6,B1:B3)',
    [
      input('A1', 79),
      input('A2', 85),
      input('A3', 78),
      input('A4', 85),
      input('A5', 50),
      input('A6', 81),
      input('B1', 60),
      input('B2', 80),
      input('B3', 90),
    ],
    [output('C1', numberExpected(1)), output('C2', numberExpected(2)), output('C3', numberExpected(3)), output('C4', numberExpected(0))],
  ),
  fixture(
    't-dist-basic',
    'T.DIST returns the cumulative Student-t probability',
    '=T.DIST(1,1,TRUE)',
    [],
    [output('C1', numberExpected(0.75))],
  ),
  fixture('t-inv-2t-basic', 'T.INV.2T returns the symmetric Student-t cutoff', '=T.INV.2T(0.5,1)', [], [output('C1', numberExpected(1))]),
  fixture(
    'confidence-t-basic',
    'CONFIDENCE.T returns the Student-t confidence radius',
    '=CONFIDENCE.T(0.5,2,4)',
    [],
    [output('C1', numberExpected(0.764892328404345))],
  ),
  fixture(
    'gamma-inv-basic',
    'GAMMA.INV inverts the cumulative gamma distribution',
    '=GAMMA.INV(0.08030139707139418,3,2)',
    [],
    [output('C1', numberExpected(2))],
  ),
  fixture(
    't-test-basic',
    'T.TEST returns a paired two-tailed probability',
    '=T.TEST(A1:A3,B1:B3,2,1)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('B1', 1), input('B2', 3), input('B3', 3)],
    [output('C1', numberExpected(1))],
  ),
  fixture('forecast-basic', 'FORECAST projects a new y value from paired samples', '=FORECAST(4,A1:A3,B1:B3)', regressionInputs(), [
    output('C1', numberExpected(14)),
  ]),
  fixture(
    'forecast-linear-basic',
    'FORECAST.LINEAR aliases the linear forecast helper',
    '=FORECAST.LINEAR(4,A1:A3,B1:B3)',
    regressionInputs(),
    [output('C1', numberExpected(14))],
  ),
  fixture(
    'trend-basic',
    'TREND spills fitted linear predictions over the requested new x range',
    '=TREND(A1:A3,B1:B3,D1:D2)',
    [...regressionInputs(), input('D1', 4), input('D2', 5)],
    [output('C1', numberExpected(14)), output('C2', numberExpected(17))],
  ),
  fixture(
    'growth-basic',
    'GROWTH spills fitted exponential predictions over the requested new x range',
    '=GROWTH(A1:A3,B1:B3,D1:D2)',
    [input('A1', 2), input('A2', 4), input('A3', 8), input('B1', 1), input('B2', 2), input('B3', 3), input('D1', 4), input('D2', 5)],
    [output('C1', numberExpected(16)), output('C2', numberExpected(32))],
  ),
  fixture(
    'linest-basic',
    'LINEST spills the slope and intercept for a paired linear regression',
    '=LINEST(A1:A3,B1:B3)',
    regressionInputs(),
    [output('C1', numberExpected(3)), output('D1', numberExpected(2))],
  ),
  fixture(
    'logest-basic',
    'LOGEST spills the exponential growth factor and intercept constant',
    '=LOGEST(A1:A3,B1:B3)',
    [input('A1', 2), input('A2', 4), input('A3', 8), input('B1', 1), input('B2', 2), input('B3', 3)],
    [output('C1', numberExpected(2)), output('D1', numberExpected(1))],
  ),
  fixture(
    'prob-basic',
    'PROB sums matching discrete probabilities across an inclusive lower and upper bound',
    '=PROB(A1:A4,B1:B4,2,3)',
    [
      input('A1', 1),
      input('A2', 2),
      input('A3', 3),
      input('A4', 4),
      input('B1', 0.1),
      input('B2', 0.2),
      input('B3', 0.3),
      input('B4', 0.4),
    ],
    [output('C1', numberExpected(0.5))],
  ),
  fixture(
    'trimmean-basic',
    'TRIMMEAN trims an even count of outer values before averaging',
    '=TRIMMEAN(A1:A8,0.25)',
    [input('A1', 1), input('A2', 2), input('A3', 4), input('A4', 7), input('A5', 8), input('A6', 9), input('A7', 10), input('A8', 12)],
    [output('C1', numberExpected(40 / 6))],
  ),
  fixture(
    'daverage-basic',
    'DAVERAGE averages numeric field values from matching database rows',
    '=DAVERAGE(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(8))],
  ),
  fixture(
    'dcount-basic',
    'DCOUNT counts numeric field values from matching database rows',
    '=DCOUNT(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(2))],
  ),
  fixture(
    'dcounta-basic',
    'DCOUNTA counts non-empty field values from matching database rows',
    '=DCOUNTA(A1:C5,"Height",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(2))],
  ),
  fixture('dget-basic', 'DGET extracts a single matching database field value', '=DGET(A1:C5,"Height",F1:F2)', databaseFixtureInputs(), [
    output('H1', numberExpected(130)),
  ]),
  fixture(
    'dmax-basic',
    'DMAX returns the maximum numeric field value from matching rows',
    '=DMAX(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(9))],
  ),
  fixture(
    'dmin-basic',
    'DMIN returns the minimum numeric field value from matching rows',
    '=DMIN(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(7))],
  ),
  fixture(
    'dproduct-basic',
    'DPRODUCT multiplies numeric field values from matching rows',
    '=DPRODUCT(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(63))],
  ),
  fixture(
    'dstdev-basic',
    'DSTDEV returns the sample standard deviation of matching field values',
    '=DSTDEV(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(Math.SQRT2))],
  ),
  fixture(
    'dstdevp-basic',
    'DSTDEVP returns the population standard deviation of matching field values',
    '=DSTDEVP(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(1))],
  ),
  fixture('dsum-basic', 'DSUM sums numeric field values from matching rows', '=DSUM(A1:C5,"Yield",E1:E2)', databaseFixtureInputs(), [
    output('H1', numberExpected(16)),
  ]),
  fixture(
    'dvar-basic',
    'DVAR returns the sample variance of matching field values',
    '=DVAR(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(2))],
  ),
  fixture(
    'dvarp-basic',
    'DVARP returns the population variance of matching field values',
    '=DVARP(A1:C5,"Yield",E1:E2)',
    databaseFixtureInputs(),
    [output('H1', numberExpected(1))],
  ),
]
