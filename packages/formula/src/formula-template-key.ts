import { lexFormula, type Token } from './lexer.js'

const CELL_REF_RE = /^(\$?)([A-Z]+)(\$?)([1-9][0-9]*)$/
const COLUMN_REF_RE = /^(\$?)([A-Z]+)$/
const ROW_REF_RE = /^(\$?)([1-9][0-9]*)$/

interface ParsedCellReference {
  colAbsolute: boolean
  rowAbsolute: boolean
  col: number
  row: number
}

interface ParsedAxisReference {
  absolute: boolean
  index: number
}

interface ParsedReferenceToken {
  kind: 'cell' | 'row' | 'col'
  key: string
}

interface TemplateReferenceMatch {
  key: string
  nextIndex: number
}

export function buildRelativeFormulaTemplateTokenKey(source: string, ownerRow: number, ownerCol: number): string {
  const fastKey = tryBuildFastRelativeFormulaTemplateTokenKey(source, ownerRow, ownerCol)
  if (fastKey !== undefined) {
    return fastKey
  }
  const tokens = lexFormula(source.startsWith('=') ? source.slice(1) : source)
  const keyParts: string[] = []

  for (let index = 0; index < tokens.length; ) {
    const token = tokens[index]!
    if (token.kind === 'eof') {
      keyParts.push('eof')
      break
    }

    if (token.kind === 'identifier' && tokens[index + 1]?.kind === 'lparen') {
      keyParts.push(`fn:${token.value.toUpperCase()}`)
      index += 1
      continue
    }

    const explicitSheetMatch = matchTemplateReferenceWithExplicitSheet(tokens, index, ownerRow, ownerCol)
    if (explicitSheetMatch) {
      keyParts.push(explicitSheetMatch.key)
      index = explicitSheetMatch.nextIndex
      continue
    }

    const referenceMatch = matchTemplateReference(tokens, index, ownerRow, ownerCol)
    if (referenceMatch) {
      keyParts.push(referenceMatch.key)
      index = referenceMatch.nextIndex
      continue
    }

    keyParts.push(`tok:${token.kind}:${JSON.stringify(token.value)}`)
    index += 1
  }

  return keyParts.join('|')
}

function tryBuildFastRelativeFormulaTemplateTokenKey(source: string, ownerRow: number, ownerCol: number): string | undefined {
  const input = source.startsWith('=') ? source.slice(1) : source
  if (
    input.includes("'") ||
    input.includes('"') ||
    input.includes('!') ||
    input.includes('#') ||
    input.includes('.') ||
    input.includes('[') ||
    input.includes(']')
  ) {
    return undefined
  }

  const keyParts: string[] = []
  let index = 0
  while (index < input.length) {
    index = skipWhitespace(input, index)
    if (index >= input.length) {
      break
    }

    const char = input[index]!
    if (isIdentifierStart(char)) {
      const identifier = readIdentifier(input, index)
      const nextIndex = skipWhitespace(input, identifier.nextIndex)
      if (input[nextIndex] === '(') {
        keyParts.push(`fn:${identifier.text.toUpperCase()}`)
        index = identifier.nextIndex
        continue
      }
      const referenceMatch = matchFastTemplateReference(input, index, ownerRow, ownerCol)
      if (!referenceMatch) {
        return undefined
      }
      keyParts.push(referenceMatch.key)
      index = referenceMatch.nextIndex
      continue
    }

    if (isDigit(char)) {
      const rowRangeMatch = matchFastTemplateReference(input, index, ownerRow, ownerCol)
      if (rowRangeMatch?.kind === 'range') {
        keyParts.push(rowRangeMatch.key)
        index = rowRangeMatch.nextIndex
        continue
      }
      const number = readNumber(input, index)
      keyParts.push(`tok:number:${JSON.stringify(number.text)}`)
      index = number.nextIndex
      continue
    }

    const tokenKind = fastTokenKindForChar(char)
    if (!tokenKind) {
      return undefined
    }
    keyParts.push(`tok:${tokenKind}:${JSON.stringify(char)}`)
    index += 1
  }
  keyParts.push('eof')
  return keyParts.join('|')
}

interface FastToken {
  readonly text: string
  readonly nextIndex: number
}

interface FastParsedReference {
  readonly kind: 'cell' | 'row' | 'col'
  readonly key: string
  readonly nextIndex: number
}

interface FastTemplateReferenceMatch {
  readonly key: string
  readonly nextIndex: number
  readonly kind: 'single' | 'range'
}

function matchFastTemplateReference(
  input: string,
  index: number,
  ownerRow: number,
  ownerCol: number,
): FastTemplateReferenceMatch | undefined {
  const startRef = parseFastReferenceToken(input, index, ownerRow, ownerCol, false)
  if (!startRef) {
    return undefined
  }
  const afterStart = skipWhitespace(input, startRef.nextIndex)
  if (input[afterStart] === ':') {
    const endRef = parseFastReferenceToken(input, skipWhitespace(input, afterStart + 1), ownerRow, ownerCol, true)
    if (!endRef || endRef.kind !== startRef.kind) {
      return undefined
    }
    const refKind = startRef.kind === 'cell' ? 'cells' : startRef.kind === 'row' ? 'rows' : 'cols'
    return {
      key: `range:${refKind}:.:${startRef.key}:${endRef.key}`,
      nextIndex: endRef.nextIndex,
      kind: 'range',
    }
  }
  if (startRef.kind === 'row') {
    return undefined
  }
  return {
    key: `${startRef.kind}:.:${startRef.key}`,
    nextIndex: startRef.nextIndex,
    kind: 'single',
  }
}

function parseFastReferenceToken(
  input: string,
  index: number,
  ownerRow: number,
  ownerCol: number,
  allowStandaloneRow: boolean,
): FastParsedReference | undefined {
  const token = isDigit(input[index]!) ? readNumber(input, index) : readIdentifier(input, index)
  const upperValue = token.text.toUpperCase()
  const cell = parseCellReferenceParts(upperValue)
  if (cell) {
    return {
      kind: 'cell',
      key: buildRelativeCellReferenceKey(cell, ownerRow, ownerCol),
      nextIndex: token.nextIndex,
    }
  }
  const column = parseAxisReferenceParts(upperValue, 'column')
  if (column) {
    return {
      kind: 'col',
      key: buildRelativeAxisReferenceKey(column, ownerCol),
      nextIndex: token.nextIndex,
    }
  }
  if (!allowStandaloneRow) {
    return undefined
  }
  const row = parseAxisReferenceParts(token.text, 'row')
  if (!row) {
    return undefined
  }
  return {
    kind: 'row',
    key: buildRelativeAxisReferenceKey(row, ownerRow),
    nextIndex: token.nextIndex,
  }
}

function readIdentifier(input: string, index: number): FastToken {
  let nextIndex = index
  while (nextIndex < input.length) {
    const char = input[nextIndex]!
    if (!isIdentifierPart(char)) {
      break
    }
    nextIndex += 1
  }
  return {
    text: input.slice(index, nextIndex),
    nextIndex,
  }
}

function readNumber(input: string, index: number): FastToken {
  let nextIndex = index
  while (nextIndex < input.length && isDigit(input[nextIndex]!)) {
    nextIndex += 1
  }
  return {
    text: input.slice(index, nextIndex),
    nextIndex,
  }
}

function skipWhitespace(input: string, index: number): number {
  let nextIndex = index
  while (nextIndex < input.length && /\s/.test(input[nextIndex]!)) {
    nextIndex += 1
  }
  return nextIndex
}

function isIdentifierStart(char: string): boolean {
  return (char >= 'A' && char <= 'Z') || (char >= 'a' && char <= 'z') || char === '$' || char === '_'
}

function isIdentifierPart(char: string): boolean {
  return isIdentifierStart(char) || isDigit(char)
}

function isDigit(char: string): boolean {
  return char >= '0' && char <= '9'
}

function fastTokenKindForChar(char: string): string | undefined {
  switch (char) {
    case '(':
      return 'lparen'
    case ')':
      return 'rparen'
    case '+':
      return 'plus'
    case '-':
      return 'minus'
    case '*':
      return 'star'
    case '/':
      return 'slash'
    case '^':
      return 'caret'
    case ',':
      return 'comma'
    case ';':
      return 'semicolon'
    case '&':
      return 'ampersand'
    default:
      return undefined
  }
}

function matchTemplateReferenceWithExplicitSheet(
  tokens: readonly Token[],
  index: number,
  ownerRow: number,
  ownerCol: number,
): TemplateReferenceMatch | undefined {
  const sheetToken = tokens[index]
  if ((sheetToken?.kind !== 'identifier' && sheetToken?.kind !== 'quotedIdentifier') || tokens[index + 1]?.kind !== 'bang') {
    return undefined
  }
  return matchTemplateReference(tokens, index + 2, ownerRow, ownerCol, sheetToken.value)
}

function matchTemplateReference(
  tokens: readonly Token[],
  index: number,
  ownerRow: number,
  ownerCol: number,
  sheetName?: string,
): TemplateReferenceMatch | undefined {
  const startToken = tokens[index]
  if (!startToken) {
    return undefined
  }

  const allowStandaloneRow = sheetName !== undefined || tokens[index + 1]?.kind === 'colon'
  const startRef = parseReferenceToken(startToken, ownerRow, ownerCol, allowStandaloneRow)
  if (!startRef) {
    return undefined
  }

  if (tokens[index + 1]?.kind === 'colon') {
    const endToken = tokens[index + 2]!
    const endRef = parseReferenceToken(endToken, ownerRow, ownerCol, true)
    if (!endRef || endRef.kind !== startRef.kind) {
      return undefined
    }
    const refKind = startRef.kind === 'cell' ? 'cells' : startRef.kind === 'row' ? 'rows' : 'cols'
    return {
      key: `range:${refKind}:${templateSheetKey(sheetName)}:${startRef.key}:${endRef.key}`,
      nextIndex: index + 3,
    }
  }

  if (tokens[index + 1]?.kind === 'hash' && startRef.kind === 'cell') {
    return {
      key: `spill:${templateSheetKey(sheetName)}:${startRef.key}`,
      nextIndex: index + 2,
    }
  }

  return {
    key: `${startRef.kind}:${templateSheetKey(sheetName)}:${startRef.key}`,
    nextIndex: index + 1,
  }
}

function parseReferenceToken(
  token: Token,
  ownerRow: number,
  ownerCol: number,
  allowStandaloneRow: boolean,
): ParsedReferenceToken | undefined {
  if (token.kind !== 'identifier' && token.kind !== 'number') {
    return undefined
  }
  const rawValue = token.value
  const upperValue = rawValue.toUpperCase()

  const cell = parseCellReferenceParts(upperValue)
  if (cell) {
    return {
      kind: 'cell',
      key: buildRelativeCellReferenceKey(cell, ownerRow, ownerCol),
    }
  }

  const column = parseAxisReferenceParts(upperValue, 'column')
  if (column) {
    return {
      kind: 'col',
      key: buildRelativeAxisReferenceKey(column, ownerCol),
    }
  }

  if (!allowStandaloneRow) {
    return undefined
  }
  const row = parseAxisReferenceParts(rawValue, 'row')
  if (!row) {
    return undefined
  }
  return {
    kind: 'row',
    key: buildRelativeAxisReferenceKey(row, ownerRow),
  }
}

function templateSheetKey(sheetName: string | undefined): string {
  return sheetName === undefined ? '.' : JSON.stringify(sheetName)
}

function buildRelativeCellReferenceKey(parsed: ParsedCellReference, ownerRow: number, ownerCol: number): string {
  const colKey = parsed.colAbsolute ? `ac${parsed.col}` : `rc${parsed.col - ownerCol}`
  const rowKey = parsed.rowAbsolute ? `ar${parsed.row}` : `rr${parsed.row - ownerRow}`
  return `${colKey}:${rowKey}`
}

function buildRelativeAxisReferenceKey(parsed: ParsedAxisReference, ownerIndex: number): string {
  return parsed.absolute ? `a${parsed.index}` : `r${parsed.index - ownerIndex}`
}

function parseCellReferenceParts(ref: string): ParsedCellReference | undefined {
  const match = CELL_REF_RE.exec(ref)
  if (!match) {
    return undefined
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match
  return {
    colAbsolute: colAbsolute === '$',
    rowAbsolute: rowAbsolute === '$',
    col: columnToIndex(columnText!),
    row: Number.parseInt(rowText!, 10) - 1,
  }
}

function parseAxisReferenceParts(ref: string, kind: 'row' | 'column'): ParsedAxisReference | undefined {
  const match = (kind === 'row' ? ROW_REF_RE : COLUMN_REF_RE).exec(ref.toUpperCase())
  if (!match) {
    return undefined
  }
  return kind === 'row'
    ? {
        absolute: match[1] === '$',
        index: Number.parseInt(match[2]!, 10) - 1,
      }
    : {
        absolute: match[1] === '$',
        index: columnToIndex(match[2]!),
      }
}

function columnToIndex(column: string): number {
  let value = 0
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64)
  }
  return value - 1
}
