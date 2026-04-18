import { ValueTag } from './enums.js'
import type { CellValue } from './types.js'

const DATE_LIKE_HEADER_PATTERN = /(?:date|month|quarter|week|day|year|as of|opened|closed|created|updated)/iu
const DATE_LIKE_FORMULA_PATTERN = /\b(?:DATE|DATEVALUE|EDATE|EOMONTH|TODAY|NOW|WORKDAY(?:\.INTL)?)\s*\(/iu

export function isDateLikeHeaderValue(value: CellValue): boolean {
  return value.tag === ValueTag.String && DATE_LIKE_HEADER_PATTERN.test(value.value)
}

export function isLikelyExcelDateSerialValue(value: CellValue): boolean {
  return value.tag === ValueTag.Number && Number.isInteger(value.value) && value.value >= 20_000 && value.value <= 90_000
}

export function formulaLooksDateLike(formula: string | undefined): boolean {
  return typeof formula === 'string' && DATE_LIKE_FORMULA_PATTERN.test(formula)
}
