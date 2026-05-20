import type {
  CellRangeRef,
  LiteralInput,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookValidationComparisonOperator,
  WorkbookValidationErrorStyle,
  WorkbookValidationListSourceSnapshot,
} from '@bilig/protocol'
import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'
import {
  attributeNameMatches,
  decodeAscii,
  decodeBytes,
  decodeCellAddress,
  encodeCellAddress,
  isAsciiWhitespace,
  isXmlNameByte,
  skipAsciiWhitespace,
} from './xlsx-large-simple-xml-byte-utils.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39

export function readLargeSimpleDataValidationsFromBytes(
  sheetName: string,
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): WorkbookDataValidationSnapshot[] | null {
  const validations: WorkbookDataValidationSnapshot[] = []
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] === slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
    if (tagEnd === null) {
      return null
    }
    if (tag.localName !== 'dataValidation') {
      index = tagEnd + 1
      continue
    }
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing
      ? { start: contentStart, end: contentStart }
      : findClosingTag(bytes, contentStart, 'dataValidation', endIndex)
    if (!closing) {
      return null
    }
    const parsed = parseDataValidation(sheetName, bytes, tag.endIndex, tagEnd, contentStart, closing.start)
    if (parsed === null) {
      return null
    }
    validations.push(...parsed)
    index = selfClosing ? tagEnd + 1 : closing.end
  }
  return validations
}

export function countLargeSimpleDataValidationsFromBytes(
  sheetName: string,
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): number | null {
  let count = 0
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] === slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
    if (tagEnd === null) {
      return null
    }
    if (tag.localName !== 'dataValidation') {
      index = tagEnd + 1
      continue
    }
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing
      ? { start: contentStart, end: contentStart }
      : findClosingTag(bytes, contentStart, 'dataValidation', endIndex)
    if (!closing) {
      return null
    }
    const parsedCount = countDataValidation(sheetName, bytes, tag.endIndex, tagEnd, contentStart, closing.start)
    if (parsedCount === null) {
      return null
    }
    count += parsedCount
    index = selfClosing ? tagEnd + 1 : closing.end
  }
  return count
}

function parseDataValidation(
  sheetName: string,
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  contentStart: number,
  contentEnd: number,
): WorkbookDataValidationSnapshot[] | null {
  const formula1 = readElementText(bytes, contentStart, contentEnd, 'formula1')
  const formula2 = readElementText(bytes, contentStart, contentEnd, 'formula2')
  if (formula1 === null || formula2 === null || !hasOnlyFormulaChildren(bytes, contentStart, contentEnd)) {
    return null
  }
  const rule = parseValidationRule(
    sheetName,
    readAttribute(bytes, nameEnd, tagEnd, 'type') ?? 'any',
    readAttribute(bytes, nameEnd, tagEnd, 'operator'),
    formula1,
    formula2,
  )
  const sqref = readAttribute(bytes, nameEnd, tagEnd, 'sqref')
  if (!rule || !sqref) {
    return null
  }
  return sqref
    .trim()
    .split(/\s+/u)
    .flatMap((rangeRef) => {
      const range = parseSqrefRange(sheetName, rangeRef)
      if (!range) {
        return []
      }
      const validation: WorkbookDataValidationSnapshot = { range, rule: cloneValidationRule(rule) }
      const allowBlank = parseBooleanAttribute(readAttribute(bytes, nameEnd, tagEnd, 'allowBlank'))
      if (allowBlank !== undefined) {
        validation.allowBlank = allowBlank
      }
      const showDropDown = parseBooleanAttribute(readAttribute(bytes, nameEnd, tagEnd, 'showDropDown'))
      if (showDropDown !== undefined) {
        validation.showDropdown = !showDropDown
      }
      const promptTitle = readAttribute(bytes, nameEnd, tagEnd, 'promptTitle')
      if (promptTitle !== null) {
        validation.promptTitle = promptTitle
      }
      const promptMessage = readAttribute(bytes, nameEnd, tagEnd, 'prompt')
      if (promptMessage !== null) {
        validation.promptMessage = promptMessage
      }
      const errorStyle = parseErrorStyle(readAttribute(bytes, nameEnd, tagEnd, 'errorStyle'))
      if (errorStyle !== undefined) {
        validation.errorStyle = errorStyle
      }
      const errorTitle = readAttribute(bytes, nameEnd, tagEnd, 'errorTitle')
      if (errorTitle !== null) {
        validation.errorTitle = errorTitle
      }
      const errorMessage = readAttribute(bytes, nameEnd, tagEnd, 'error')
      if (errorMessage !== null) {
        validation.errorMessage = errorMessage
      }
      return [validation]
    })
}

function countDataValidation(
  sheetName: string,
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  contentStart: number,
  contentEnd: number,
): number | null {
  const formula1 = readElementText(bytes, contentStart, contentEnd, 'formula1')
  const formula2 = readElementText(bytes, contentStart, contentEnd, 'formula2')
  if (formula1 === null || formula2 === null || !hasOnlyFormulaChildren(bytes, contentStart, contentEnd)) {
    return null
  }
  if (
    !parseValidationRule(
      sheetName,
      readAttribute(bytes, nameEnd, tagEnd, 'type') ?? 'any',
      readAttribute(bytes, nameEnd, tagEnd, 'operator'),
      formula1,
      formula2,
    )
  ) {
    return null
  }
  const sqref = readAttribute(bytes, nameEnd, tagEnd, 'sqref')
  if (!sqref) {
    return null
  }
  let count = 0
  for (const rangeRef of sqref.trim().split(/\s+/u)) {
    if (parseSqrefRange(sheetName, rangeRef)) {
      count += 1
    }
  }
  return count
}

function parseValidationRule(
  sheetName: string,
  type: string,
  operatorValue: string | null,
  formula1: string | undefined,
  formula2: string | undefined,
): WorkbookDataValidationRuleSnapshot | null {
  switch (type) {
    case 'list': {
      if (!formula1) {
        return null
      }
      const listFormula = parseListFormula(sheetName, formula1)
      return listFormula ? { kind: 'list', ...listFormula } : null
    }
    case 'any':
      return { kind: 'any' }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength': {
      const operator = parseComparisonOperator(operatorValue) ?? 'between'
      const values = [formula1, formula2].flatMap((formula) => (formula === undefined ? [] : [parseScalarFormulaValue(formula)]))
      return { kind: type, operator, values }
    }
    default:
      return null
  }
}

function parseListFormula(
  sheetName: string,
  formula: string,
): Pick<Extract<WorkbookDataValidationRuleSnapshot, { kind: 'list' }>, 'values' | 'source'> | null {
  const trimmed = formula.trim()
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return { values: trimmed.slice(1, -1).replaceAll('""', '"').split(',') }
  }
  const expression = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  const sameSheetSource = parseSourceReference(sheetName, expression)
  if (sameSheetSource) {
    return { source: sameSheetSource }
  }
  const sheetReference = parseSheetReference(expression)
  if (sheetReference) {
    const source = parseSourceReference(sheetReference.sheetName, sheetReference.reference)
    return source ? { source } : null
  }
  const structured = parseStructuredReference(expression)
  if (structured) {
    return { source: structured }
  }
  return /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(expression) ? { source: { kind: 'named-range', name: expression } } : null
}

function parseSourceReference(sheetName: string, reference: string): WorkbookValidationListSourceSnapshot | null {
  const parts = reference.split(':')
  if (parts.length === 1) {
    const address = normalizeAddress(parts[0] ?? '')
    return address ? { kind: 'cell-ref', sheetName, address } : null
  }
  if (parts.length === 2) {
    const startAddress = normalizeAddress(parts[0] ?? '')
    const endAddress = normalizeAddress(parts[1] ?? '')
    return startAddress && endAddress ? { kind: 'range-ref', sheetName, startAddress, endAddress } : null
  }
  return null
}

function parseSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  const quoted = parseQuotedSheetReference(value)
  if (quoted) {
    return quoted
  }
  const separatorIndex = value.indexOf('!')
  if (separatorIndex <= 0) {
    return null
  }
  const sheetName = value.slice(0, separatorIndex).trim()
  const reference = value.slice(separatorIndex + 1).trim()
  return sheetName.length > 0 && reference.length > 0 ? { sheetName, reference } : null
}

function parseQuotedSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  if (!value.startsWith("'")) {
    return null
  }
  let sheetName = ''
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'" && value[index + 1] === "'") {
      sheetName += "'"
      index += 1
      continue
    }
    if (character === "'" && value[index + 1] === '!') {
      const reference = value.slice(index + 2).trim()
      return sheetName.trim().length > 0 && reference.length > 0 ? { sheetName, reference } : null
    }
    sheetName += character
  }
  return null
}

function parseStructuredReference(value: string): WorkbookValidationListSourceSnapshot | null {
  const match = /^([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]$/u.exec(value)
  return match ? { kind: 'structured-ref', tableName: match[1] ?? '', columnName: match[2] ?? '' } : null
}

function parseSqrefRange(sheetName: string, value: string): CellRangeRef | null {
  const parts = value.replaceAll('$', '').split(':')
  if (parts.length < 1 || parts.length > 2) {
    return null
  }
  const start = decodeCellAddress(parts[0] ?? '')
  const end = decodeCellAddress(parts[1] ?? parts[0] ?? '')
  return start && end
    ? {
        sheetName,
        startAddress: encodeCellAddress(start.row, start.column),
        endAddress: encodeCellAddress(end.row, end.column),
      }
    : null
}

function normalizeAddress(address: string): string | null {
  const decoded = decodeCellAddress(address.trim().replaceAll('$', ''))
  return decoded ? encodeCellAddress(decoded.row, decoded.column) : null
}

function parseScalarFormulaValue(formula: string): LiteralInput {
  const trimmed = formula.trim()
  if (/^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/u.test(trimmed)) {
    const numberValue = Number(trimmed)
    if (Number.isFinite(numberValue)) {
      return numberValue
    }
  }
  if (/^TRUE$/iu.test(trimmed)) {
    return true
  }
  if (/^FALSE$/iu.test(trimmed)) {
    return false
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replaceAll('""', '"')
  }
  return trimmed
}

function parseBooleanAttribute(value: string | null): boolean | undefined {
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  return undefined
}

function parseComparisonOperator(value: string | null): WorkbookValidationComparisonOperator | null {
  if (value === null) {
    return null
  }
  switch (value) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual':
      return value
    default:
      return null
  }
}

function parseErrorStyle(value: string | null): WorkbookValidationErrorStyle | undefined {
  if (value === null) {
    return undefined
  }
  switch (value) {
    case 'stop':
    case 'warning':
    case 'information':
      return value
    default:
      return undefined
  }
}

function cloneValidationRule(rule: WorkbookDataValidationRuleSnapshot): WorkbookDataValidationRuleSnapshot {
  return structuredClone(rule)
}

function hasOnlyFormulaChildren(bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const closing = bytes[index + 1] === slash
    const tag = readXmlTagName(bytes, index + (closing ? 2 : 1))
    if (!tag) {
      index += 1
      continue
    }
    if (!closing && tag.localName !== 'formula1' && tag.localName !== 'formula2') {
      return false
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
    if (tagEnd === null) {
      return false
    }
    index = tagEnd + 1
  }
  return true
}

function readElementText(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  elementName: 'formula1' | 'formula2',
): string | undefined | null {
  const tag = findNextOpeningTag(bytes, startIndex, endIndex, elementName)
  if (!tag) {
    return undefined
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null) {
    return null
  }
  if (isSelfClosingTag(bytes, tagEnd)) {
    return ''
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? decodeXmlText(decodeBytes(bytes, tagEnd + 1, closing.start).trim()) : null
}

function findNextOpeningTag(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  localName: string,
): { readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] === slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === localName) {
      return { start: index, nameEnd: tag.endIndex }
    }
    index += 1
  }
  return null
}

function findClosingTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] !== slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 2)
    if (tag?.localName === localName) {
      const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
      return tagEnd === null ? null : { start: index, end: tagEnd + 1 }
    }
    index += 1
  }
  return null
}

function findTagEnd(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuote || byte === singleQuote) {
      quote = byte
      continue
    }
    if (byte === greaterThan) {
      return index
    }
  }
  return null
}

function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === slash || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === 58) {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart
    ? null
    : { localName: readKnownXmlLocalName(bytes, localNameStart, index) ?? decodeAscii(bytes, localNameStart, index), endIndex: index }
}

function readAttribute(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeXmlText(decodeBytes(bytes, range.start, range.end)) : null
}

function readXmlAttributeRangeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      index += 1
      continue
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      index += 1
      continue
    }
    const valueStart = index + 1
    index = valueStart
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    const valueEnd = index
    if (attributeNameMatches(bytes, nameStart, nameEnd, attributeName)) {
      return { start: valueStart, end: valueEnd }
    }
    index += 1
  }
  return null
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
