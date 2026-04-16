import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type {
  ExcelExpectedValue,
  ExcelFixtureCase,
  ExcelFixtureExpectedOutput,
  ExcelFixtureFamily,
  ExcelFixtureInputCell,
} from './index.js'

export interface TextFixtureCase {
  name: string
  args: readonly CellValue[]
  expected: CellValue
  note?: string
}

export interface TextFixtureGroup {
  builtin:
    | 'LEN'
    | 'LENB'
    | 'CONCAT'
    | 'LEFT'
    | 'LEFTB'
    | 'RIGHT'
    | 'RIGHTB'
    | 'MID'
    | 'MIDB'
    | 'TRIM'
    | 'UPPER'
    | 'LOWER'
    | 'FIND'
    | 'FINDB'
    | 'SEARCH'
    | 'SEARCHB'
    | 'REPLACEB'
    | 'ASC'
    | 'JIS'
    | 'DBCS'
    | 'VALUE'
    | 'TEXTBEFORE'
    | 'TEXTAFTER'
    | 'TEXTJOIN'
    | 'TEXTSPLIT'
  cases: readonly TextFixtureCase[]
}

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

function stringExpected(value: string): ExcelExpectedValue {
  return { kind: 'string', value }
}

function errorExpected(code: ErrorCode, display: string): ExcelExpectedValue {
  return { kind: 'error', code, display }
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
    id: createExcelFixtureId('text', slug),
    family: 'text' as const,
    title,
    formula,
    inputs,
    outputs,
    sheetName: 'Sheet1',
  }
  return notes === undefined ? base : { ...base, notes }
}

export const TEXT_FIXTURE_METADATA = {
  source: 'excel-web-like text builtin tranche',
  version: 2,
  builtins: [
    'LEN',
    'LENB',
    'CONCAT',
    'LEFT',
    'LEFTB',
    'RIGHT',
    'RIGHTB',
    'MID',
    'MIDB',
    'TRIM',
    'UPPER',
    'LOWER',
    'FIND',
    'FINDB',
    'SEARCH',
    'SEARCHB',
    'REPLACEB',
    'ASC',
    'JIS',
    'DBCS',
    'VALUE',
    'TEXTBEFORE',
    'TEXTAFTER',
    'TEXTJOIN',
    'TEXTSPLIT',
  ] as const,
} as const

export const TEXT_FIXTURES: readonly TextFixtureGroup[] = [
  {
    builtin: 'LEN',
    cases: [
      { name: 'counts plain string length', args: [text('hello')], expected: number(5) },
      { name: 'coerces booleans to text', args: [bool(true)], expected: number(4) },
      { name: 'treats empty as empty string', args: [empty()], expected: number(0) },
    ],
  },
  {
    builtin: 'LENB',
    cases: [
      { name: 'counts ASCII bytes', args: [text('hello')], expected: number(5) },
      { name: 'counts UTF-8 bytes', args: [text('é')], expected: number(2) },
    ],
  },
  {
    builtin: 'CONCAT',
    cases: [
      {
        name: 'joins mixed scalar values',
        args: [text('alpha'), number(2), empty()],
        expected: text('alpha2'),
      },
      {
        name: 'coerces booleans to uppercase logical text',
        args: [bool(false), text('-ok')],
        expected: text('FALSE-ok'),
      },
    ],
  },
  {
    builtin: 'LEFT',
    cases: [
      { name: 'defaults to one character', args: [text('alpha')], expected: text('a') },
      {
        name: 'takes requested prefix length',
        args: [text('alpha'), number(3)],
        expected: text('alp'),
      },
      {
        name: 'zero length returns empty string',
        args: [text('alpha'), empty()],
        expected: text(''),
      },
    ],
  },
  {
    builtin: 'LEFTB',
    cases: [
      {
        name: 'takes requested byte prefix length',
        args: [text('alpha'), number(3)],
        expected: text('alp'),
      },
    ],
  },
  {
    builtin: 'RIGHT',
    cases: [
      { name: 'defaults to one character', args: [text('alpha')], expected: text('a') },
      {
        name: 'takes requested suffix length',
        args: [text('alpha'), number(2)],
        expected: text('ha'),
      },
      {
        name: 'large suffix returns whole string',
        args: [text('alpha'), number(99)],
        expected: text('alpha'),
      },
    ],
  },
  {
    builtin: 'RIGHTB',
    cases: [
      {
        name: 'takes requested byte suffix length',
        args: [text('alpha'), number(2)],
        expected: text('ha'),
      },
    ],
  },
  {
    builtin: 'MID',
    cases: [
      {
        name: 'extracts substring from one-based start',
        args: [text('alphabet'), number(2), number(3)],
        expected: text('lph'),
      },
      {
        name: 'start beyond end returns empty string',
        args: [text('alpha'), number(9), number(2)],
        expected: text(''),
      },
      {
        name: 'zero count returns empty string',
        args: [text('alpha'), number(2), empty()],
        expected: text(''),
      },
    ],
  },
  {
    builtin: 'MIDB',
    cases: [
      {
        name: 'extracts byte-based substring from one-based start',
        args: [text('alphabet'), number(2), number(3)],
        expected: text('lph'),
      },
    ],
  },
  {
    builtin: 'TRIM',
    cases: [
      {
        name: 'collapses internal spaces',
        args: [text('  alpha   beta  ')],
        expected: text('alpha beta'),
      },
      {
        name: 'leaves clean strings alone',
        args: [text('alpha beta')],
        expected: text('alpha beta'),
      },
    ],
  },
  {
    builtin: 'UPPER',
    cases: [{ name: 'uppercases text', args: [text('Alpha beta')], expected: text('ALPHA BETA') }],
  },
  {
    builtin: 'LOWER',
    cases: [{ name: 'lowercases text', args: [text('Alpha BETA')], expected: text('alpha beta') }],
  },
  {
    builtin: 'FIND',
    cases: [
      {
        name: 'finds first case-sensitive position',
        args: [text('ph'), text('alphabet')],
        expected: number(3),
      },
      {
        name: 'respects one-based start',
        args: [text('a'), text('bananas'), number(3)],
        expected: number(4),
      },
      {
        name: 'empty needle returns start position',
        args: [text(''), text('alpha'), number(3)],
        expected: number(3),
      },
    ],
  },
  {
    builtin: 'FINDB',
    cases: [
      {
        name: 'finds first byte-based position',
        args: [text('ph'), text('alphabet')],
        expected: number(3),
      },
      {
        name: 'respects byte-based one-based start',
        args: [text('d'), text('abcdef'), number(3)],
        expected: number(4),
      },
    ],
  },
  {
    builtin: 'SEARCH',
    cases: [
      {
        name: 'searches case-insensitively',
        args: [text('PH'), text('alphabet')],
        expected: number(3),
      },
      {
        name: 'supports wildcard question mark',
        args: [text('b?d'), text('ABCD')],
        expected: number(2),
      },
      { name: 'supports escaped wildcard', args: [text('~*'), text('a*b')], expected: number(2) },
    ],
  },
  {
    builtin: 'SEARCHB',
    cases: [
      {
        name: 'searches case-insensitively with byte positions',
        args: [text('PH'), text('alphabet')],
        expected: number(3),
      },
      {
        name: 'supports wildcard question mark',
        args: [text('b?d'), text('ABCD')],
        expected: number(2),
      },
    ],
  },
  {
    builtin: 'REPLACEB',
    cases: [
      {
        name: 'replaces a byte-based span',
        args: [text('alphabet'), number(3), number(2), text('Z')],
        expected: text('alZabet'),
      },
    ],
  },
  {
    builtin: 'ASC',
    cases: [
      {
        name: 'converts full-width latin text to half-width',
        args: [text('ＡＢＣ　１２３')],
        expected: text('ABC 123'),
      },
      {
        name: 'converts full-width katakana to half-width pairs',
        args: [text('ガギグゲゴ')],
        expected: text('ｶﾞｷﾞｸﾞｹﾞｺﾞ'),
      },
    ],
  },
  {
    builtin: 'JIS',
    cases: [
      {
        name: 'converts half-width latin text to full-width',
        args: [text('ABC 123')],
        expected: text('ＡＢＣ　１２３'),
      },
      {
        name: 'converts half-width katakana pairs to full-width',
        args: [text('ｶﾞｷﾞｸﾞｹﾞｺﾞ')],
        expected: text('ガギグゲゴ'),
      },
    ],
  },
  {
    builtin: 'DBCS',
    cases: [
      {
        name: 'converts half-width latin text to double-byte full-width',
        args: [text('ABC 123')],
        expected: text('ＡＢＣ　１２３'),
      },
    ],
  },
  {
    builtin: 'VALUE',
    cases: [
      { name: 'parses trimmed numeric text', args: [text(' 42 ')], expected: number(42) },
      { name: 'coerces booleans to numbers', args: [bool(true)], expected: number(1) },
      { name: 'treats empty as zero', args: [empty()], expected: number(0) },
    ],
  },
  {
    builtin: 'TEXTBEFORE',
    cases: [
      {
        name: 'returns text before the first delimiter',
        args: [text('alpha-beta'), text('-')],
        expected: text('alpha'),
      },
      {
        name: 'supports negative instance search',
        args: [text('alpha-beta-gamma'), text('-'), number(-1)],
        expected: text('alpha-beta'),
      },
    ],
  },
  {
    builtin: 'TEXTAFTER',
    cases: [
      {
        name: 'returns text after the first delimiter',
        args: [text('alpha-beta'), text('-')],
        expected: text('beta'),
      },
      {
        name: 'supports negative instance search',
        args: [text('alpha-beta-gamma'), text('-'), number(-1)],
        expected: text('gamma'),
      },
    ],
  },
  {
    builtin: 'TEXTJOIN',
    cases: [
      {
        name: 'joins text with a delimiter while ignoring empty items',
        args: [text('-'), bool(true), text('alpha'), empty(), text('beta')],
        expected: text('alpha-beta'),
      },
      {
        name: 'retains empty strings when ignore-empty is false',
        args: [text('|'), bool(false), text('a'), text(''), text('b')],
        expected: text('a||b'),
      },
    ],
  },
]

export const canonicalTextFixtures: readonly ExcelFixtureCase[] = [
  fixture('len-counts-plain-string-length', 'LEN counts plain string length', '=LEN("hello")', [], [output('A1', numberExpected(5))]),
  fixture(
    'len-coerces-booleans-to-text',
    'LEN coerces booleans to text',
    '=LEN(A1)',
    [input('A1', true)],
    [output('A2', numberExpected(4))],
  ),
  fixture(
    'len-treats-empty-as-empty-string',
    'LEN treats empty as empty string',
    '=LEN(A1)',
    [input('A1', null)],
    [output('A2', numberExpected(0))],
  ),
  fixture(
    'concat-joins-mixed-scalar-values',
    'CONCAT joins mixed scalar values',
    '=CONCAT("alpha",2,A1)',
    [input('A1', null)],
    [output('A2', stringExpected('alpha2'))],
  ),
  fixture(
    'concat-coerces-booleans-to-uppercase-logical-text',
    'CONCAT coerces booleans to uppercase logical text',
    '=CONCAT(A1,"-ok")',
    [input('A1', false)],
    [output('A2', stringExpected('FALSE-ok'))],
  ),
  fixture('left-defaults-to-one-character', 'LEFT defaults to one character', '=LEFT("alpha")', [], [output('A1', stringExpected('a'))]),
  fixture(
    'left-takes-requested-prefix-length',
    'LEFT takes requested prefix length',
    '=LEFT("alpha",3)',
    [],
    [output('A1', stringExpected('alp'))],
  ),
  fixture(
    'left-zero-length-returns-empty-string',
    'LEFT zero length returns empty string',
    '=LEFT("alpha",A1)',
    [input('A1', null)],
    [output('A2', stringExpected(''))],
  ),
  fixture('right-defaults-to-one-character', 'RIGHT defaults to one character', '=RIGHT("alpha")', [], [output('A1', stringExpected('a'))]),
  fixture(
    'right-takes-requested-suffix-length',
    'RIGHT takes requested suffix length',
    '=RIGHT("alpha",2)',
    [],
    [output('A1', stringExpected('ha'))],
  ),
  fixture(
    'right-large-suffix-returns-whole-string',
    'RIGHT large suffix returns whole string',
    '=RIGHT("alpha",99)',
    [],
    [output('A1', stringExpected('alpha'))],
  ),
  fixture(
    'mid-extracts-substring-from-one-based-start',
    'MID extracts substring from one-based start',
    '=MID("alphabet",2,3)',
    [],
    [output('A1', stringExpected('lph'))],
  ),
  fixture(
    'mid-start-beyond-end-returns-empty-string',
    'MID start beyond end returns empty string',
    '=MID("alpha",9,2)',
    [],
    [output('A1', stringExpected(''))],
  ),
  fixture(
    'mid-zero-count-returns-empty-string',
    'MID zero count returns empty string',
    '=MID("alpha",2,A1)',
    [input('A1', null)],
    [output('A2', stringExpected(''))],
  ),
  fixture(
    'trim-collapses-internal-spaces',
    'TRIM collapses internal spaces',
    '=TRIM("  alpha   beta  ")',
    [],
    [output('A1', stringExpected('alpha beta'))],
  ),
  fixture(
    'trim-leaves-clean-strings-alone',
    'TRIM leaves clean strings alone',
    '=TRIM("alpha beta")',
    [],
    [output('A1', stringExpected('alpha beta'))],
  ),
  fixture('upper-uppercases-text', 'UPPER uppercases text', '=UPPER("Alpha beta")', [], [output('A1', stringExpected('ALPHA BETA'))]),
  fixture('lower-lowercases-text', 'LOWER lowercases text', '=LOWER("Alpha BETA")', [], [output('A1', stringExpected('alpha beta'))]),
  fixture(
    'find-finds-first-case-sensitive-position',
    'FIND finds first case-sensitive position',
    '=FIND("ph","alphabet")',
    [],
    [output('A1', numberExpected(3))],
  ),
  fixture(
    'find-respects-one-based-start',
    'FIND respects one-based start',
    '=FIND("a","bananas",3)',
    [],
    [output('A1', numberExpected(4))],
  ),
  fixture(
    'find-empty-needle-returns-start-position',
    'FIND empty needle returns start position',
    '=FIND("","alpha",3)',
    [],
    [output('A1', numberExpected(3))],
  ),
  fixture(
    'search-searches-case-insensitively',
    'SEARCH searches case-insensitively',
    '=SEARCH("PH","alphabet")',
    [],
    [output('A1', numberExpected(3))],
  ),
  fixture(
    'search-supports-wildcard-question-mark',
    'SEARCH supports wildcard question mark',
    '=SEARCH("b?d","ABCD")',
    [],
    [output('A1', numberExpected(2))],
  ),
  fixture(
    'search-supports-escaped-wildcard',
    'SEARCH supports escaped wildcard',
    '=SEARCH("~*","a*b")',
    [],
    [output('A1', numberExpected(2))],
  ),
  fixture(
    'asc-basic',
    'ASC converts full-width text to half-width',
    '=ASC("ＡＢＣ　１２３")',
    [],
    [output('A1', stringExpected('ABC 123'))],
  ),
  fixture(
    'jis-basic',
    'JIS converts half-width text to full-width',
    '=JIS("ABC 123")',
    [],
    [output('A1', stringExpected('ＡＢＣ　１２３'))],
  ),
  fixture(
    'dbcs-basic',
    'DBCS converts half-width text to double-byte full-width',
    '=DBCS("ABC 123")',
    [],
    [output('A1', stringExpected('ＡＢＣ　１２３'))],
  ),
  fixture(
    'value-parses-trimmed-numeric-text',
    'VALUE parses trimmed numeric text',
    '=VALUE(" 42 ")',
    [],
    [output('A1', numberExpected(42))],
  ),
  fixture(
    'value-coerces-booleans-to-numbers',
    'VALUE coerces booleans to numbers',
    '=VALUE(A1)',
    [input('A1', true)],
    [output('A2', numberExpected(1))],
  ),
  fixture('value-treats-empty-as-zero', 'VALUE treats empty as zero', '=VALUE(A1)', [input('A1', null)], [output('A2', numberExpected(0))]),
  fixture(
    'textbefore-basic',
    'TEXTBEFORE returns the substring before the delimiter',
    '=TEXTBEFORE("alpha-beta","-")',
    [],
    [output('A1', stringExpected('alpha'))],
  ),
  fixture(
    'textafter-basic',
    'TEXTAFTER returns the substring after the delimiter',
    '=TEXTAFTER("alpha-beta","-")',
    [],
    [output('A1', stringExpected('beta'))],
  ),
  fixture(
    'textjoin-basic',
    'TEXTJOIN joins scalar and empty values with ignore-empty enabled',
    '=TEXTJOIN("-",TRUE,A1:A3)',
    [input('A1', 'alpha'), input('A2', null), input('A3', 'beta')],
    [output('A4', stringExpected('alpha-beta'))],
  ),
  fixture(
    'textsplit-basic',
    'TEXTSPLIT spills a ragged row/column split with #N/A padding',
    '=TEXTSPLIT(A1,",","|")',
    [input('A1', 'red,blue|green')],
    [
      output('A2', stringExpected('red')),
      output('B2', stringExpected('blue')),
      output('A3', stringExpected('green')),
      output('B3', errorExpected(ErrorCode.NA, '#N/A')),
    ],
  ),
]

function empty(): CellValue {
  return { tag: ValueTag.Empty }
}

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function bool(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

export function textValueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}
