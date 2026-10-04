import { builtinCapabilityManifest, type BuiltinCapabilityCategory } from './builtin-capabilities.js'

export interface FormulaHelpArg {
  readonly label: string
  readonly optional?: boolean
}

export interface FormulaHelpEntry {
  readonly kind: 'function'
  readonly name: string
  readonly category: BuiltinCapabilityCategory
  readonly summary: string
  readonly args: readonly FormulaHelpArg[]
  readonly variadic?: boolean
}

const FALLBACK_ARGS: readonly FormulaHelpArg[] = [{ label: '…' }]

const CURATED_HELP: Record<
  string,
  {
    readonly summary: string
    readonly args: readonly FormulaHelpArg[]
    readonly variadic?: boolean
  }
> = {
  SUBSTITUTE: {
    summary: 'Replace matching text, optionally at a specific occurrence.',
    args: [{ label: 'text' }, { label: 'old_text' }, { label: 'new_text' }, { label: '[instance_num]', optional: true }],
  },
  SUM: {
    summary: 'Add numbers, ranges, and spill results.',
    args: [{ label: 'number1' }, { label: '[number2]', optional: true }],
    variadic: true,
  },
  AVERAGE: {
    summary: 'Return the arithmetic mean for the supplied values.',
    args: [{ label: 'number1' }, { label: '[number2]', optional: true }],
    variadic: true,
  },
  COUNT: {
    summary: 'Count numeric values in cells, ranges, and arguments.',
    args: [{ label: 'value1' }, { label: '[value2]', optional: true }],
    variadic: true,
  },
  COUNTA: {
    summary: 'Count non-empty values in cells, ranges, and arguments.',
    args: [{ label: 'value1' }, { label: '[value2]', optional: true }],
    variadic: true,
  },
  MIN: {
    summary: 'Return the smallest numeric value from the supplied inputs.',
    args: [{ label: 'number1' }, { label: '[number2]', optional: true }],
    variadic: true,
  },
  MAX: {
    summary: 'Return the largest numeric value from the supplied inputs.',
    args: [{ label: 'number1' }, { label: '[number2]', optional: true }],
    variadic: true,
  },
  IF: {
    summary: 'Choose between two results based on a logical test.',
    args: [{ label: 'logical_test' }, { label: 'value_if_true' }, { label: '[value_if_false]', optional: true }],
  },
  IFERROR: {
    summary: 'Replace an error result with a fallback value.',
    args: [{ label: 'value' }, { label: 'value_if_error' }],
  },
  IFNA: {
    summary: 'Replace a #N/A result with a fallback value.',
    args: [{ label: 'value' }, { label: 'value_if_na' }],
  },
  AND: {
    summary: 'Return TRUE only when every argument is truthy.',
    args: [{ label: 'logical1' }, { label: '[logical2]', optional: true }],
    variadic: true,
  },
  OR: {
    summary: 'Return TRUE when any supplied argument is truthy.',
    args: [{ label: 'logical1' }, { label: '[logical2]', optional: true }],
    variadic: true,
  },
  NOT: {
    summary: 'Invert a logical value.',
    args: [{ label: 'logical' }],
  },
  XLOOKUP: {
    summary: 'Look up a value in one array and return the matching item from another.',
    args: [
      { label: 'lookup_value' },
      { label: 'lookup_array' },
      { label: 'return_array' },
      { label: '[if_not_found]', optional: true },
      { label: '[match_mode]', optional: true },
      { label: '[search_mode]', optional: true },
    ],
  },
  XMATCH: {
    summary: 'Return the relative position of a lookup value inside an array.',
    args: [
      { label: 'lookup_value' },
      { label: 'lookup_array' },
      { label: '[match_mode]', optional: true },
      { label: '[search_mode]', optional: true },
    ],
  },
  VLOOKUP: {
    summary: 'Search the first column of a table and return a value from a target column.',
    args: [{ label: 'lookup_value' }, { label: 'table_array' }, { label: 'col_index_num' }, { label: '[range_lookup]', optional: true }],
  },
  HLOOKUP: {
    summary: 'Search the first row of a table and return a value from a target row.',
    args: [{ label: 'lookup_value' }, { label: 'table_array' }, { label: 'row_index_num' }, { label: '[range_lookup]', optional: true }],
  },
  INDEX: {
    summary: 'Return a value or reference at the given row and column position.',
    args: [{ label: 'array' }, { label: 'row_num' }, { label: '[column_num]', optional: true }],
  },
  MATCH: {
    summary: 'Return the relative position of a lookup value in a one-dimensional range.',
    args: [{ label: 'lookup_value' }, { label: 'lookup_array' }, { label: '[match_type]', optional: true }],
  },
  SUMIF: {
    summary: 'Sum cells that match a single condition.',
    args: [{ label: 'range' }, { label: 'criteria' }, { label: '[sum_range]', optional: true }],
  },
  SUMIFS: {
    summary: 'Sum cells that match multiple conditions.',
    args: [
      { label: 'sum_range' },
      { label: 'criteria_range1' },
      { label: 'criteria1' },
      { label: '[criteria_range2]', optional: true },
      { label: '[criteria2]', optional: true },
    ],
    variadic: true,
  },
  COUNTIF: {
    summary: 'Count cells that match a single condition.',
    args: [{ label: 'range' }, { label: 'criteria' }],
  },
  COUNTIFS: {
    summary: 'Count cells that match multiple conditions.',
    args: [
      { label: 'criteria_range1' },
      { label: 'criteria1' },
      { label: '[criteria_range2]', optional: true },
      { label: '[criteria2]', optional: true },
    ],
    variadic: true,
  },
  TEXTJOIN: {
    summary: 'Join multiple values with a delimiter and optional blank skipping.',
    args: [{ label: 'delimiter' }, { label: 'ignore_empty' }, { label: 'text1' }, { label: '[text2]', optional: true }],
    variadic: true,
  },
  CONCAT: {
    summary: 'Concatenate text values without a delimiter.',
    args: [{ label: 'text1' }, { label: '[text2]', optional: true }],
    variadic: true,
  },
  FILTER: {
    summary: 'Return rows or columns that satisfy a filter condition.',
    args: [{ label: 'array' }, { label: 'include' }, { label: '[if_empty]', optional: true }],
  },
  UNIQUE: {
    summary: 'Return unique rows or columns from an array.',
    args: [{ label: 'array' }, { label: '[by_col]', optional: true }, { label: '[exactly_once]', optional: true }],
  },
  SORT: {
    summary: 'Sort rows or columns in ascending or descending order.',
    args: [
      { label: 'array' },
      { label: '[sort_index]', optional: true },
      { label: '[sort_order]', optional: true },
      { label: '[by_col]', optional: true },
    ],
  },
  SORTN: {
    summary: 'Return the first rows after sorting, with optional tie handling.',
    args: [
      { label: 'range' },
      { label: '[n]', optional: true },
      { label: '[display_ties_mode]', optional: true },
      { label: '[sort_column1]', optional: true },
      { label: '[is_ascending1]', optional: true },
    ],
    variadic: true,
  },
  ROUND: {
    summary: 'Round a number to the specified number of digits.',
    args: [{ label: 'number' }, { label: 'num_digits' }],
  },
  ROUNDUP: {
    summary: 'Round a number away from zero.',
    args: [{ label: 'number' }, { label: 'num_digits' }],
  },
  ROUNDDOWN: {
    summary: 'Round a number toward zero.',
    args: [{ label: 'number' }, { label: 'num_digits' }],
  },
  LEFT: {
    summary: 'Return the leftmost characters from a text value.',
    args: [{ label: 'text' }, { label: '[num_chars]', optional: true }],
  },
  RIGHT: {
    summary: 'Return the rightmost characters from a text value.',
    args: [{ label: 'text' }, { label: '[num_chars]', optional: true }],
  },
  MID: {
    summary: 'Return a substring from the middle of a text value.',
    args: [{ label: 'text' }, { label: 'start_num' }, { label: 'num_chars' }],
  },
  DATE: {
    summary: 'Build a date serial from year, month, and day components.',
    args: [{ label: 'year' }, { label: 'month' }, { label: 'day' }],
  },
  TODAY: {
    summary: 'Return the current date in the workbook volatile context.',
    args: [],
  },
  NOW: {
    summary: 'Return the current date and time in the workbook volatile context.',
    args: [],
  },
  YEAR: {
    summary: 'Extract the year from a date serial.',
    args: [{ label: 'serial_number' }],
  },
  MONTH: {
    summary: 'Extract the month from a date serial.',
    args: [{ label: 'serial_number' }],
  },
  DAY: {
    summary: 'Extract the day of month from a date serial.',
    args: [{ label: 'serial_number' }],
  },
}

const CATEGORY_SUMMARIES: Record<BuiltinCapabilityCategory, string> = {
  aggregation: 'Combine values across ranges and arrays.',
  logical: 'Branch or test logical conditions.',
  information: 'Inspect value types, sheet state, or cell metadata.',
  text: 'Manipulate and search text values.',
  'date-time': 'Compute or extract date and time values.',
  'lookup-reference': 'Look up positions, references, or matching values.',
  statistical: 'Compute descriptive and inferential statistics.',
  'dynamic-array': 'Produce or transform spill-aware array results.',
  lambda: 'Compose spreadsheet functions with lambda semantics.',
  math: 'Perform scalar math, rounding, and numeric transforms.',
}

export const functionHelpEntries: readonly FormulaHelpEntry[] = builtinCapabilityManifest
  .map((capability) => {
    const curated = CURATED_HELP[capability.name]
    return Object.assign(
      {
        kind: `function` as const,
        name: capability.name,
        category: capability.category,
        summary: curated?.summary ?? CATEGORY_SUMMARIES[capability.category],
        args: curated?.args ?? FALLBACK_ARGS,
      },
      curated?.variadic ? { variadic: true } : {},
    )
  })
  .toSorted((left, right) => left.name.localeCompare(right.name))
