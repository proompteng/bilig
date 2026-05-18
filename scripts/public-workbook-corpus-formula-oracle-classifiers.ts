import { parseRangeAddress } from '../packages/formula/src/addressing.js'
import { evaluateAst } from '../packages/formula/src/js-evaluator.js'
import { parseFormula } from '../packages/formula/src/parser.js'
import { ValueTag } from '../packages/protocol/src/enums.js'
import type { CellValue, LiteralInput, WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { cellValuesMatchOracle, formatCellValue } from './public-workbook-corpus-workbook.ts'

export const localeDecimalCommaFormulaOracleUnsupportedClassification =
  'xlsx.publicCorpus.formulaOracle:localeDecimalCommaTextCoercionUnsupported'

export interface FormulaOracleMismatchDetail {
  readonly sheetName: string
  readonly address: string
  readonly expected: CellValue
  readonly actual: CellValue
  readonly message: string
}

export interface UnsupportedFormulaOracleClassification {
  readonly unsupported: boolean
  readonly classifications: readonly string[]
  readonly evidence: readonly string[]
}

export function classifyUnsupportedLocaleDecimalCommaFormulaOracle(
  snapshot: WorkbookSnapshot,
  validation: {
    readonly mismatches: readonly string[]
    readonly mismatchDetails: readonly FormulaOracleMismatchDetail[]
  },
): UnsupportedFormulaOracleClassification {
  if (validation.mismatchDetails.length === 0 || validation.mismatchDetails.length !== validation.mismatches.length) {
    return emptyUnsupportedFormulaOracleClassification()
  }
  const explanations = validation.mismatchDetails.flatMap((mismatch) => {
    const explanation = evaluateMismatchWithDecimalCommaTextCoercion(snapshot, mismatch)
    return explanation ? [explanation] : []
  })
  if (explanations.length !== validation.mismatchDetails.length) {
    return emptyUnsupportedFormulaOracleClassification()
  }
  return {
    unsupported: true,
    classifications: [localeDecimalCommaFormulaOracleUnsupportedClassification],
    evidence: [
      `Formula oracle mismatches matched cached Excel values when text operands were evaluated with decimal-comma numeric coercion: ${String(
        explanations.length,
      )} mismatches.`,
      ...explanations
        .slice(0, 25)
        .map(
          (explanation) =>
            `locale-decimal-comma-formula=${explanation.sheetName}!${explanation.address} cached ${formatCellValue(
              explanation.expected,
            )} refs ${explanation.references.join(', ')}`,
        ),
    ],
  }
}

function emptyUnsupportedFormulaOracleClassification(): UnsupportedFormulaOracleClassification {
  return { unsupported: false, classifications: [], evidence: [] }
}

function evaluateMismatchWithDecimalCommaTextCoercion(
  snapshot: WorkbookSnapshot,
  mismatch: FormulaOracleMismatchDetail,
): { readonly sheetName: string; readonly address: string; readonly expected: CellValue; readonly references: readonly string[] } | null {
  const cellsBySheet = indexSnapshotCells(snapshot)
  const formulaCell = cellsBySheet.get(mismatch.sheetName)?.get(canonicalFormulaAddress(mismatch.address))
  if (!formulaCell?.formula) {
    return null
  }
  const coercedReferences: string[] = []
  const resolveCellValue = (sheetName: string, address: string): CellValue => {
    const value = cellsBySheet.get(sheetName)?.get(canonicalFormulaAddress(address))?.value
    return coerceDecimalCommaTextCellValue(value, `${sheetName}!${address}`, coercedReferences)
  }
  try {
    const actual = evaluateAst(parseFormula(formulaCell.formula), {
      sheetName: mismatch.sheetName,
      currentAddress: mismatch.address,
      resolveCell: resolveCellValue,
      resolveRange: (sheetName, start, end, refKind) =>
        resolveSnapshotRangeValuesWithDecimalCommaTextCoercion({
          sheetName,
          start,
          end,
          refKind,
          resolveCellValue,
        }),
    })
    return coercedReferences.length > 0 && cellValuesMatchOracle(actual, mismatch.expected)
      ? { sheetName: mismatch.sheetName, address: mismatch.address, expected: mismatch.expected, references: coercedReferences }
      : null
  } catch {
    return null
  }
}

function indexSnapshotCells(snapshot: WorkbookSnapshot): Map<string, Map<string, WorkbookSnapshot['sheets'][number]['cells'][number]>> {
  const cellsBySheet = new Map<string, Map<string, WorkbookSnapshot['sheets'][number]['cells'][number]>>()
  for (const sheet of snapshot.sheets) {
    cellsBySheet.set(sheet.name, new Map(sheet.cells.map((cell) => [canonicalFormulaAddress(cell.address), cell])))
  }
  return cellsBySheet
}

function coerceDecimalCommaTextCellValue(value: LiteralInput | undefined, reference: string, coercedReferences: string[]): CellValue {
  if (typeof value !== 'string') {
    return literalInputToCellValue(value)
  }
  const numeric = parseDecimalCommaNumericText(value)
  if (numeric === undefined) {
    return { tag: ValueTag.String, value, stringId: 0 }
  }
  coercedReferences.push(`${reference}=${value}`)
  return { tag: ValueTag.Number, value: numeric }
}

function literalInputToCellValue(value: LiteralInput | undefined): CellValue {
  if (value === undefined || value === null) {
    return { tag: ValueTag.Empty }
  }
  switch (typeof value) {
    case 'number':
      return { tag: ValueTag.Number, value }
    case 'boolean':
      return { tag: ValueTag.Boolean, value }
    case 'string':
      return { tag: ValueTag.String, value, stringId: 0 }
    case 'bigint':
    case 'function':
    case 'object':
    case 'symbol':
    case 'undefined':
      return { tag: ValueTag.Empty }
  }
}

function parseDecimalCommaNumericText(value: string): number | undefined {
  const trimmed = value.trim()
  if (!/^[+-]?\d+,\d+(?:[eE][+-]?\d+)?$/u.test(trimmed)) {
    return undefined
  }
  const numeric = Number(trimmed.replace(',', '.'))
  return Number.isFinite(numeric) ? numeric : undefined
}

function resolveSnapshotRangeValuesWithDecimalCommaTextCoercion(args: {
  readonly sheetName: string
  readonly start: string
  readonly end: string
  readonly refKind: 'cells' | 'rows' | 'cols'
  readonly resolveCellValue: (sheetName: string, address: string) => CellValue
}): CellValue[] {
  if (args.refKind !== 'cells') {
    return []
  }
  const range = parseRangeAddress(`${args.start}:${args.end}`)
  if (range.kind !== 'cells') {
    return []
  }
  const rowStart = Math.min(range.start.row, range.end.row)
  const rowEnd = Math.max(range.start.row, range.end.row)
  const colStart = Math.min(range.start.col, range.end.col)
  const colEnd = Math.max(range.start.col, range.end.col)
  const cellCount = (rowEnd - rowStart + 1) * (colEnd - colStart + 1)
  if (cellCount > 100_000) {
    throw new Error('Locale decimal-comma formula classifier range budget exceeded')
  }
  const values: CellValue[] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      values.push(args.resolveCellValue(args.sheetName, addressFromCoordinates(row, col)))
    }
  }
  return values
}

function addressFromCoordinates(row: number, col: number): string {
  let column = ''
  let remaining = col + 1
  while (remaining > 0) {
    const mod = (remaining - 1) % 26
    column = String.fromCharCode(65 + mod) + column
    remaining = Math.floor((remaining - 1) / 26)
  }
  return `${column}${String(row + 1)}`
}

function canonicalFormulaAddress(address: string): string {
  return address.replaceAll('$', '').toUpperCase()
}
