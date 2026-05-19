import * as fc from 'fast-check'

export type FuzzJsonValue = null | boolean | number | string | FuzzJsonValue[] | { readonly [key: string]: FuzzJsonValue }

export interface FuzzCellRangeRef {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

export interface FuzzWorkbookSnapshotCell {
  readonly address: string
  readonly row?: number
  readonly col?: number
  readonly value?: null | boolean | number | string
  readonly formula?: string
  readonly format?: string
}

export interface FuzzWorkbookSnapshotSheet {
  readonly id?: number
  readonly name: string
  readonly order: number
  readonly cells: readonly FuzzWorkbookSnapshotCell[]
}

export interface FuzzWorkbookSnapshot {
  readonly version: 1
  readonly workbook: {
    readonly name: string
  }
  readonly sheets: readonly FuzzWorkbookSnapshotSheet[]
}

export interface FuzzCellSnapshot {
  readonly sheetName: string
  readonly address: string
  readonly value:
    | { readonly tag: 0 }
    | { readonly tag: 1; readonly value: number }
    | { readonly tag: 2; readonly value: boolean }
    | { readonly tag: 3; readonly value: string; readonly stringId: number }
    | { readonly tag: 4; readonly code: number }
  readonly flags: number
  readonly version: number
}

export interface FuzzStructuralAxisEdit {
  readonly axis: 'row' | 'column'
  readonly sheetName: string
  readonly start: number
  readonly count: number
  readonly mode: 'insert' | 'delete'
}

const sheetNames = ['Sheet1', 'Sheet2', 'Revenue', 'Inputs', 'Summary'] as const
const safeText = fc.string({ minLength: 0, maxLength: 24 })

export const fuzzSheetNameArbitrary: fc.Arbitrary<string> = fc.constantFrom(...sheetNames)

export const fuzzWorkbookAddressArbitrary: fc.Arbitrary<string> = fc
  .record({
    row: fc.integer({ min: 0, max: 49 }),
    col: fc.integer({ min: 0, max: 25 }),
  })
  .map(({ row, col }) => `${columnLabel(col)}${String(row + 1)}`)

export const fuzzLiteralInputArbitrary: fc.Arbitrary<null | boolean | number | string> = fc.oneof(
  fc.constant(null),
  fc.boolean(),
  fc.double({ noNaN: true, noDefaultInfinity: true, min: -1_000_000, max: 1_000_000 }).map((value) => (Object.is(value, -0) ? 0 : value)),
  safeText,
)

export const fuzzCellRangeRefArbitrary: fc.Arbitrary<FuzzCellRangeRef> = fc
  .record({
    sheetName: fuzzSheetNameArbitrary,
    startRow: fc.integer({ min: 0, max: 24 }),
    startCol: fc.integer({ min: 0, max: 12 }),
    rowSpan: fc.integer({ min: 0, max: 6 }),
    colSpan: fc.integer({ min: 0, max: 6 }),
  })
  .map(({ sheetName, startRow, startCol, rowSpan, colSpan }) => ({
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(startRow + rowSpan, startCol + colSpan),
  }))

export const fuzzCellValueArbitrary: fc.Arbitrary<FuzzCellSnapshot['value']> = fc.oneof(
  fc.constant({ tag: 0 as const }),
  fc.double({ noNaN: true, noDefaultInfinity: true, min: -1_000_000, max: 1_000_000 }).map((value) => ({
    tag: 1 as const,
    value: Object.is(value, -0) ? 0 : value,
  })),
  fc.boolean().map((value) => ({ tag: 2 as const, value })),
  safeText.map((value) => ({ tag: 3 as const, value, stringId: stableStringId(value) })),
  fc.integer({ min: 0, max: 8 }).map((code) => ({ tag: 4 as const, code })),
)

export const fuzzCellSnapshotArbitrary: fc.Arbitrary<FuzzCellSnapshot> = fc.record({
  sheetName: fuzzSheetNameArbitrary,
  address: fuzzWorkbookAddressArbitrary,
  value: fuzzCellValueArbitrary,
  flags: fc.integer({ min: 0, max: 16 }),
  version: fc.integer({ min: 0, max: 10_000 }),
})

const fuzzWorkbookSnapshotCellArbitrary: fc.Arbitrary<FuzzWorkbookSnapshotCell> = fc
  .record({
    row: fc.integer({ min: 0, max: 12 }),
    col: fc.integer({ min: 0, max: 8 }),
    value: fc.option(fuzzLiteralInputArbitrary, { nil: undefined }),
    formula: fc.option(fc.constantFrom('A1+1', 'SUM(A1:A2)', 'IF(A1>0,1,0)'), { nil: undefined }),
    format: fc.option(fc.constantFrom('0.00', '0%', '@', 'yyyy-mm-dd'), { nil: undefined }),
  })
  .map(({ row, col, value, formula, format }) => {
    const cell: {
      address: string
      row: number
      col: number
      value?: null | boolean | number | string
      formula?: string
      format?: string
    } = {
      address: formatAddress(row, col),
      row,
      col,
    }
    if (value !== undefined) {
      cell.value = value
    }
    if (formula !== undefined) {
      cell.formula = formula
    }
    if (format !== undefined) {
      cell.format = format
    }
    return cell
  })

export const fuzzWorkbookSnapshotArbitrary: fc.Arbitrary<FuzzWorkbookSnapshot> = fc
  .uniqueArray(
    fc.record({
      id: fc.integer({ min: 0, max: 10_000 }),
      name: fuzzSheetNameArbitrary,
      cells: fc.array(fuzzWorkbookSnapshotCellArbitrary, { minLength: 0, maxLength: 16 }),
    }),
    { minLength: 1, maxLength: 3, selector: (sheet) => sheet.name },
  )
  .map((sheets) => ({
    version: 1 as const,
    workbook: { name: 'Fuzz Workbook' },
    sheets: sheets.map((sheet, order) => ({
      id: sheet.id,
      name: sheet.name,
      order,
      cells: sheet.cells,
    })),
  }))

export const fuzzCellStylePatchArbitrary: fc.Arbitrary<Record<string, unknown>> = fc.record(
  {
    font: fc.option(
      fc.record({
        bold: fc.option(fc.boolean(), { nil: undefined }),
        italic: fc.option(fc.boolean(), { nil: undefined }),
        color: fc.option(fc.constantFrom('#111827', '#2563eb', '#dc2626'), { nil: undefined }),
      }),
      { nil: undefined },
    ),
    fill: fc.option(fc.record({ backgroundColor: fc.constantFrom('#ffffff', '#fef3c7', '#dcfce7') }), { nil: undefined }),
    alignment: fc.option(fc.record({ horizontal: fc.constantFrom('left', 'center', 'right') }), { nil: undefined }),
  },
  { requiredKeys: [] },
)

export const fuzzStructuralAxisEditArbitrary: fc.Arbitrary<FuzzStructuralAxisEdit> = fc.record({
  axis: fc.constantFrom('row', 'column'),
  sheetName: fuzzSheetNameArbitrary,
  start: fc.integer({ min: 0, max: 24 }),
  count: fc.integer({ min: 1, max: 8 }),
  mode: fc.constantFrom('insert', 'delete'),
})

export const fuzzJsonValueArbitrary: fc.Arbitrary<FuzzJsonValue> = fc.letrec<{ value: FuzzJsonValue }>((tie) => ({
  value: fc.oneof(
    fc.constant(null),
    fc.boolean(),
    fc.double({ noNaN: true, noDefaultInfinity: true, min: -10_000, max: 10_000 }),
    safeText,
    fc.array(tie('value'), { minLength: 0, maxLength: 4 }),
    fc.dictionary(fc.string({ minLength: 1, maxLength: 12 }), tie('value'), { maxKeys: 4 }),
  ) as fc.Arbitrary<FuzzJsonValue>,
})).value

export function corruptRecord(record: Record<string, unknown>, key: string, value: unknown = Symbol('corrupt')): Record<string, unknown> {
  const next = { ...record }
  if (typeof value === 'symbol') {
    delete next[key]
    return next
  }
  next[key] = value
  return next
}

export function cloneJsonValue<T>(value: T): T {
  return structuredClone(value)
}

export function formatAddress(row: number, col: number): string {
  return `${columnLabel(col)}${String(row + 1)}`
}

function columnLabel(index: number): string {
  let current = index + 1
  let label = ''
  while (current > 0) {
    current -= 1
    label = String.fromCharCode(65 + (current % 26)) + label
    current = Math.floor(current / 26)
  }
  return label
}

function stableStringId(value: string): number {
  let hash = 0
  for (let index = 0; index < value.length; index += 1) {
    hash = (hash * 31 + value.charCodeAt(index)) >>> 0
  }
  return hash
}
