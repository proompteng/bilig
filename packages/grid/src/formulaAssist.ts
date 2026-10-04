import { functionHelpEntries, type FormulaHelpArg, type FormulaHelpEntry, type BuiltinCapabilityCategory } from '@bilig/formula'
import type { WorkbookDefinedNameSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'

export type { FormulaHelpArg, FormulaHelpEntry } from '@bilig/formula'

export interface DefinedNameSuggestion {
  readonly kind: 'defined-name'
  readonly name: string
  readonly summary: string
  readonly insertText: string
}

export interface FunctionSuggestion {
  readonly kind: 'function'
  readonly name: string
  readonly category: BuiltinCapabilityCategory
  readonly summary: string
  readonly signature: string
}

export type FormulaSuggestion = FunctionSuggestion | DefinedNameSuggestion

export interface FormulaReplaceResult {
  readonly value: string
  readonly caret: number
}

export interface FormulaAssistState {
  readonly tokenStart: number | null
  readonly tokenEnd: number | null
  readonly suggestions: readonly FormulaSuggestion[]
  readonly activeFunction: {
    readonly entry: FormulaHelpEntry
    readonly activeArgumentIndex: number
    readonly signature: string
  } | null
}

const IDENTIFIER_PATTERN = /[A-Za-z0-9_.]/
const COMMON_FUNCTIONS = new Set([
  'SUM',
  'AVERAGE',
  'COUNT',
  'COUNTA',
  'MIN',
  'MAX',
  'IF',
  'IFERROR',
  'IFNA',
  'AND',
  'OR',
  'NOT',
  'XLOOKUP',
  'XMATCH',
  'VLOOKUP',
  'HLOOKUP',
  'INDEX',
  'MATCH',
  'SUMIF',
  'SUMIFS',
  'COUNTIF',
  'COUNTIFS',
  'TEXTJOIN',
  'CONCAT',
  'FILTER',
  'UNIQUE',
  'SORT',
  'SORTN',
  'ROUND',
  'DATE',
  'TODAY',
  'NOW',
])

const functionHelpByName = new Map(functionHelpEntries.map((entry) => [entry.name, entry]))

function isIdentifierCharacter(char: string | undefined): boolean {
  return char !== undefined && IDENTIFIER_PATTERN.test(char)
}

function isEscapedDoubleQuote(source: string, index: number): boolean {
  return source[index] === '"' && source[index + 1] === '"'
}

function isEscapedSingleQuote(source: string, index: number): boolean {
  return source[index] === "'" && source[index + 1] === "'"
}

function previousNonWhitespaceChar(source: string, index: number): string | null {
  for (let cursor = index; cursor >= 0; cursor -= 1) {
    const char = source[cursor]
    if (char && !/\s/.test(char)) {
      return char
    }
  }
  return null
}

function findCalleeBeforeParen(source: string, parenIndex: number): string | null {
  let cursor = parenIndex - 1
  while (cursor >= 0 && /\s/.test(source[cursor]!)) {
    cursor -= 1
  }
  const end = cursor + 1
  while (cursor >= 0 && isIdentifierCharacter(source[cursor])) {
    cursor -= 1
  }
  const start = cursor + 1
  if (start >= end) {
    return null
  }
  const name = source.slice(start, end).toUpperCase()
  const previous = previousNonWhitespaceChar(source, start - 1)
  if (previous === '!' || previous === ']' || previous === '#') {
    return null
  }
  return name
}

function formatArgumentLabel(arg: FormulaHelpArg): string {
  return arg.optional ? `[${arg.label.replace(/^\[(.*)\]$/, '$1')}]` : arg.label
}

function formatFormulaSignature(entry: FormulaHelpEntry): string {
  if (entry.args.length === 0) {
    return `${entry.name}()`
  }
  const labels = entry.args.map(formatArgumentLabel)
  if (entry.variadic && entry.args.at(-1)?.label !== '…') {
    labels.push('…')
  }
  return `${entry.name}(${labels.join(', ')})`
}

function formatDefinedNameSummary(value: WorkbookDefinedNameValueSnapshot): string {
  if (value === null) {
    return 'Defined name = blank'
  }
  if (typeof value === 'string') {
    return `Defined name = "${value}"`
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return `Defined name = ${String(value)}`
  }
  switch (value.kind) {
    case 'scalar':
      return `Defined name = ${String(value.value)}`
    case 'cell-ref':
      return `${value.sheetName}!${value.address}`
    case 'range-ref':
      return `${value.sheetName}!${value.startAddress}:${value.endAddress}`
    case 'structured-ref':
      return `${value.tableName}[${value.columnName}]`
    case 'formula':
      return `=${value.formula}`
  }
}

export function resolveNameBoxDisplayValue(input: {
  readonly sheetName: string
  readonly address: string
  readonly selectionLabel?: string | undefined
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[] | undefined
}): string {
  const match = input.definedNames?.find((entry) => {
    const value = entry.value
    if (!value || typeof value !== 'object' || !('kind' in value)) {
      return false
    }
    switch (value.kind) {
      case 'cell-ref':
        return value.sheetName === input.sheetName && value.address.toUpperCase() === input.address.toUpperCase()
      case 'range-ref':
        return (
          value.sheetName === input.sheetName &&
          input.selectionLabel !== undefined &&
          `${value.startAddress.toUpperCase()}:${value.endAddress.toUpperCase()}` === input.selectionLabel.toUpperCase()
        )
      case 'scalar':
      case 'structured-ref':
      case 'formula':
        return false
    }
  })
  return match?.name ?? input.selectionLabel ?? input.address
}

function buildDefinedNameSuggestions(
  prefix: string,
  definedNames: readonly WorkbookDefinedNameSnapshot[],
): readonly DefinedNameSuggestion[] {
  return definedNames
    .filter((entry) => entry.name.toUpperCase().startsWith(prefix))
    .toSorted((left, right) => left.name.localeCompare(right.name))
    .map((entry) => ({
      kind: 'defined-name' as const,
      name: entry.name,
      summary: formatDefinedNameSummary(entry.value),
      insertText: entry.name,
    }))
}

function buildFunctionSuggestions(prefix: string): readonly FunctionSuggestion[] {
  const candidates =
    prefix.length === 0
      ? functionHelpEntries.filter((entry) => COMMON_FUNCTIONS.has(entry.name))
      : functionHelpEntries.filter((entry) => entry.name.startsWith(prefix))
  return candidates
    .toSorted((left, right) => {
      const leftCommon = COMMON_FUNCTIONS.has(left.name) ? 0 : 1
      const rightCommon = COMMON_FUNCTIONS.has(right.name) ? 0 : 1
      if (leftCommon !== rightCommon) {
        return leftCommon - rightCommon
      }
      return left.name.localeCompare(right.name)
    })
    .filter((entry, index, entries) => entries.findIndex((candidate) => candidate.name === entry.name) === index)
    .map((entry) => ({
      kind: 'function' as const,
      name: entry.name,
      category: entry.category,
      summary: entry.summary,
      signature: formatFormulaSignature(entry),
    }))
}

function resolveTokenBounds(value: string, caret: number): { start: number; end: number; prefix: string } | null {
  if (!value.startsWith('=') || caret < 1) {
    return null
  }
  let start = caret
  while (start > 1 && isIdentifierCharacter(value[start - 1])) {
    start -= 1
  }
  const prefix = value.slice(start, caret).toUpperCase()
  const previous = previousNonWhitespaceChar(value, start - 1)
  const allowedPrevious =
    previous === null ||
    previous === '=' ||
    previous === '(' ||
    previous === ',' ||
    previous === '+' ||
    previous === '-' ||
    previous === '*' ||
    previous === '/' ||
    previous === '^' ||
    previous === '&' ||
    previous === '<' ||
    previous === '>'
  if (prefix.length === 0 && previous !== '=') {
    return null
  }
  if (!allowedPrevious) {
    return null
  }
  return { start, end: caret, prefix }
}

function resolveActiveFunction(value: string, caret: number) {
  if (!value.startsWith('=') || caret <= 1) {
    return null
  }
  const stack: Array<{ name: string | null; argIndex: number }> = []
  let insideString = false
  let insideQuotedIdentifier = false
  let bracketDepth = 0
  const end = Math.min(caret, value.length)

  for (let index = 1; index < end; index += 1) {
    const char = value[index]!
    if (insideString) {
      if (char === '"' && isEscapedDoubleQuote(value, index)) {
        index += 1
        continue
      }
      if (char === '"') {
        insideString = false
      }
      continue
    }
    if (insideQuotedIdentifier) {
      if (char === "'" && isEscapedSingleQuote(value, index)) {
        index += 1
        continue
      }
      if (char === "'") {
        insideQuotedIdentifier = false
      }
      continue
    }
    if (char === '"') {
      insideString = true
      continue
    }
    if (char === "'") {
      insideQuotedIdentifier = true
      continue
    }
    if (char === '[') {
      bracketDepth += 1
      continue
    }
    if (char === ']' && bracketDepth > 0) {
      bracketDepth -= 1
      continue
    }
    if (bracketDepth > 0) {
      continue
    }
    if (char === '(') {
      stack.push({ name: findCalleeBeforeParen(value, index), argIndex: 0 })
      continue
    }
    if (char === ')') {
      stack.pop()
      continue
    }
    if (char === ',' && stack.length > 0) {
      const top = stack.at(-1)
      if (top && top.name) {
        top.argIndex += 1
      }
    }
  }

  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const active = stack[index]
    if (!active?.name) {
      continue
    }
    const entry = functionHelpByName.get(active.name)
    if (!entry) {
      return null
    }
    return {
      entry,
      activeArgumentIndex: active.argIndex,
      signature: formatFormulaSignature(entry),
    }
  }
  return null
}

export function resolveFormulaAssistState(input: {
  readonly value: string
  readonly caret: number
  readonly definedNames?: readonly WorkbookDefinedNameSnapshot[]
}): FormulaAssistState {
  const token = resolveTokenBounds(input.value, input.caret)
  const prefix = token?.prefix ?? ''
  const functionSuggestions = token ? buildFunctionSuggestions(prefix) : []
  const definedNameSuggestions = token && input.definedNames ? buildDefinedNameSuggestions(prefix, input.definedNames) : []
  return {
    tokenStart: token?.start ?? null,
    tokenEnd: token?.end ?? null,
    suggestions: [...definedNameSuggestions, ...functionSuggestions].slice(0, 12),
    activeFunction: resolveActiveFunction(input.value, input.caret),
  }
}

export function applyFormulaSuggestion(input: {
  readonly value: string
  readonly tokenStart: number
  readonly tokenEnd: number
  readonly suggestion: FormulaSuggestion
}): FormulaReplaceResult {
  const before = input.value.slice(0, input.tokenStart)
  const after = input.value.slice(input.tokenEnd)
  if (input.suggestion.kind === 'defined-name') {
    const value = `${before}${input.suggestion.insertText}${after}`
    return {
      value,
      caret: before.length + input.suggestion.insertText.length,
    }
  }
  if (after.startsWith('(')) {
    const value = `${before}${input.suggestion.name}${after}`
    return {
      value,
      caret: before.length + input.suggestion.name.length + 1,
    }
  }
  const insertText = `${input.suggestion.name}()`
  return {
    value: `${before}${insertText}${after}`,
    caret: before.length + input.suggestion.name.length + 1,
  }
}
