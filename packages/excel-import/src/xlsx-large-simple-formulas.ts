import { normalizeImportedFormulaSource } from './xlsx-formula-translation.js'
import { readWorksheetSharedFormulas } from './xlsx-formulas.js'
import { formulaReferencesExternalWorkbook, formulaReferencesVolatileFunction } from './xlsx-import-warnings.js'

const formulaElementPattern =
  /<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/\2>/u
const formulaOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u

export { readWorksheetSharedFormulas as readLargeSimpleSheetSharedFormulas }

export function readLargeSimpleCellFormula(cellXml: string, sharedFormula?: string): string | null | undefined {
  const formulaXml = formulaElementPattern.exec(cellXml)?.[0]
  if (!formulaXml) {
    return undefined
  }
  const openingTag = formulaOpeningTagPattern.exec(formulaXml)?.[0]
  if (!openingTag) {
    return null
  }
  const formulaType = readXmlAttribute(openingTag, 't')
  if (formulaType === 'array' || formulaType === 'dataTable') {
    return null
  }
  if (formulaType === 'shared') {
    return normalizeLargeSimpleFormula(sharedFormula)
  }
  if (openingTag.endsWith('/>')) {
    return null
  }
  const rawFormula = decodeXmlText(formulaElementPattern.exec(cellXml)?.[3] ?? '').trim()
  return normalizeLargeSimpleFormula(rawFormula)
}

function normalizeLargeSimpleFormula(rawFormula: string | undefined): string | null {
  if (rawFormula === undefined || rawFormula.length === 0) {
    return null
  }
  const formula = normalizeImportedFormulaSource(rawFormula)
  return formulaReferencesExternalWorkbook(formula) ||
    formulaReferencesVolatileFunction(formula) ||
    formulaReferencesStructuredTable(formula)
    ? null
    : formula
}

function formulaReferencesStructuredTable(formula: string): boolean {
  return /\[[#@\w]/u.test(formula)
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}
