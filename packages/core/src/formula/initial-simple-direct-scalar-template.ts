import type { ParsedCellReferenceInfo, ParsedDependencyReference } from '@bilig/formula'

type InitialSimpleRowRelativeBinaryOperator = '+' | '-' | '*' | '/'

type InitialSimpleRowRelativeBinaryTemplateRight =
  | {
      readonly kind: 'cell'
      readonly colOffset: number
    }
  | {
      readonly kind: 'number'
      readonly text: string
    }

export interface InitialSimpleRowRelativeBinaryTemplateShape {
  readonly templateKey: string
  readonly leftColOffset: number
  readonly operator: InitialSimpleRowRelativeBinaryOperator
  readonly right: InitialSimpleRowRelativeBinaryTemplateRight
}

interface InitialRelativeCellToken {
  readonly colOffset: number
  readonly ref: ParsedCellReferenceInfo
  readonly dep: ParsedDependencyReference
  readonly symbolicRef: string
  readonly next: number
}

export interface InitialSimpleRowRelativeBinaryTemplateMatch extends InitialSimpleRowRelativeBinaryTemplateShape {
  readonly symbolicRefs: string[]
  readonly parsedDeps: ParsedDependencyReference[]
  readonly parsedSymbolicRefs: ParsedCellReferenceInfo[]
}

const INITIAL_SINGLE_COLUMN_LABELS = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
] as const

function initialReadNumberLiteral(source: string, start: number): { readonly text: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    next += 1
  }
  if (next < source.length && source.charCodeAt(next) === 46) {
    const fractionStart = next + 1
    next = fractionStart
    while (next < source.length) {
      const code = source.charCodeAt(next)
      if (code < 48 || code > 57) {
        break
      }
      next += 1
    }
    if (next === fractionStart) {
      return undefined
    }
  }
  return next === start ? undefined : { text: source.slice(start, next), next }
}

function initialIndexToColumn(index: number): string {
  if (index >= 0 && index < INITIAL_SINGLE_COLUMN_LABELS.length) {
    return INITIAL_SINGLE_COLUMN_LABELS[index]!
  }
  let value = index + 1
  let label = ''
  while (value > 0) {
    value -= 1
    label = String.fromCharCode(65 + (value % 26)) + label
    value = Math.floor(value / 26)
  }
  return label
}

function initialSourceStartsWith(source: string, expected: string, start: number): boolean {
  if (start + expected.length > source.length) {
    return false
  }
  for (let index = 0; index < expected.length; index += 1) {
    if (source.charCodeAt(start + index) !== expected.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function initialSourceMatchesColumn(source: string, expected: string, start: number): boolean {
  if (start + expected.length > source.length) {
    return false
  }
  for (let index = 0; index < expected.length; index += 1) {
    const sourceCode = source.charCodeAt(start + index)
    const expectedCode = expected.charCodeAt(index)
    if (sourceCode !== expectedCode && sourceCode !== expectedCode + 32) {
      return false
    }
  }
  return true
}

function initialCellToken(col: number, ownerRow: number, address: string, next: number, ownerCol: number): InitialRelativeCellToken {
  return {
    colOffset: col - ownerCol,
    ref: initialParsedCellRef(address, ownerRow, col),
    dep: initialParsedCellDependency(address, ownerRow, col),
    symbolicRef: address,
    next,
  }
}

function initialParsedCellRef(address: string, row: number, col: number): ParsedCellReferenceInfo {
  return {
    address,
    row,
    col,
    rowAbsolute: false,
    colAbsolute: false,
  }
}

function initialParsedCellDependency(address: string, row: number, col: number): ParsedDependencyReference {
  return {
    kind: 'cell',
    address,
    row,
    col,
    rowAbsolute: false,
    colAbsolute: false,
  }
}

function initialReadRelativeCellToken(
  source: string,
  start: number,
  ownerRow: number,
  ownerCol: number,
  ownerRowText: string,
): InitialRelativeCellToken | undefined {
  let next = start
  let oneBasedCol = 0
  let hasLowercaseColumn = false
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code >= 65 && code <= 90) {
      oneBasedCol = oneBasedCol * 26 + code - 64
      next += 1
      continue
    }
    if (code >= 97 && code <= 122) {
      oneBasedCol = oneBasedCol * 26 + code - 96
      hasLowercaseColumn = true
      next += 1
      continue
    }
    break
  }
  if (next === start || !source.startsWith(ownerRowText, next)) {
    return undefined
  }
  const rowEnd = next + ownerRowText.length
  if (rowEnd < source.length) {
    const trailingCode = source.charCodeAt(rowEnd)
    if (trailingCode >= 48 && trailingCode <= 57) {
      return undefined
    }
  }
  const col = oneBasedCol - 1
  if (col < 0) {
    return undefined
  }
  const rawColumn = source.slice(start, next)
  const address = `${hasLowercaseColumn ? rawColumn.toUpperCase() : rawColumn}${ownerRowText}`
  return initialCellToken(col, ownerRow, address, rowEnd, ownerCol)
}

function initialReadExpectedRelativeCellEnd(source: string, start: number, ownerRowText: string, column: string): number {
  if (!initialSourceMatchesColumn(source, column, start)) {
    return -1
  }
  const rowStart = start + column.length
  if (!initialSourceStartsWith(source, ownerRowText, rowStart)) {
    return -1
  }
  const rowEnd = rowStart + ownerRowText.length
  if (rowEnd < source.length) {
    const trailingCode = source.charCodeAt(rowEnd)
    if (trailingCode >= 48 && trailingCode <= 57) {
      return -1
    }
  }
  return rowEnd
}

function initialBinaryOperator(operator: string | undefined): InitialSimpleRowRelativeBinaryOperator | undefined {
  switch (operator) {
    case undefined:
      return undefined
    case '+':
    case '-':
    case '*':
    case '/':
      return operator
    default:
      return undefined
  }
}

export function tryMatchInitialSimpleRowRelativeBinaryTemplate(
  source: string,
  ownerRow: number,
  ownerCol: number,
): InitialSimpleRowRelativeBinaryTemplateMatch | undefined {
  let index = source.charCodeAt(0) === 61 ? 1 : 0
  const ownerRowText = String(ownerRow + 1)
  const left = initialReadRelativeCellToken(source, index, ownerRow, ownerCol, ownerRowText)
  if (!left) {
    return undefined
  }
  index = left.next
  const operator = initialBinaryOperator(source[index])
  if (!operator) {
    return undefined
  }
  index += 1
  const rightCell = initialReadRelativeCellToken(source, index, ownerRow, ownerCol, ownerRowText)
  if (rightCell) {
    return rightCell.next === source.length
      ? {
          leftColOffset: left.colOffset,
          operator,
          templateKey: `c${left.colOffset}${operator}c${rightCell.colOffset}`,
          right: {
            kind: 'cell',
            colOffset: rightCell.colOffset,
          },
          symbolicRefs: [left.symbolicRef, rightCell.symbolicRef],
          parsedDeps: [left.dep, rightCell.dep],
          parsedSymbolicRefs: [left.ref, rightCell.ref],
        }
      : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  return rightNumber && rightNumber.next === source.length
    ? {
        leftColOffset: left.colOffset,
        operator,
        templateKey: `c${left.colOffset}${operator}n${rightNumber.text}`,
        right: {
          kind: 'number',
          text: rightNumber.text,
        },
        symbolicRefs: [left.symbolicRef],
        parsedDeps: [left.dep],
        parsedSymbolicRefs: [left.ref],
      }
    : undefined
}

export function tryMatchInitialSimpleRowRelativeBinaryTemplateShape(
  source: string,
  ownerRow: number,
  ownerCol: number,
  shape: InitialSimpleRowRelativeBinaryTemplateShape,
): InitialSimpleRowRelativeBinaryTemplateMatch | undefined {
  let index = source.charCodeAt(0) === 61 ? 1 : 0
  const ownerRowText = String(ownerRow + 1)
  const leftCol = ownerCol + shape.leftColOffset
  if (leftCol < 0) {
    return undefined
  }
  const leftColumn = initialIndexToColumn(leftCol)
  const leftEnd = initialReadExpectedRelativeCellEnd(source, index, ownerRowText, leftColumn)
  if (leftEnd < 0) {
    return undefined
  }
  index = leftEnd
  if (source[index] !== shape.operator) {
    return undefined
  }
  index += 1
  const leftAddress = `${leftColumn}${ownerRowText}`
  const leftRef = initialParsedCellRef(leftAddress, ownerRow, leftCol)
  const leftDep = initialParsedCellDependency(leftAddress, ownerRow, leftCol)
  if (shape.right.kind === 'cell') {
    const rightCol = ownerCol + shape.right.colOffset
    if (rightCol < 0) {
      return undefined
    }
    const rightColumn = initialIndexToColumn(rightCol)
    const rightEnd = initialReadExpectedRelativeCellEnd(source, index, ownerRowText, rightColumn)
    if (rightEnd < 0 || rightEnd !== source.length) {
      return undefined
    }
    const rightAddress = `${rightColumn}${ownerRowText}`
    const rightRef = initialParsedCellRef(rightAddress, ownerRow, rightCol)
    const rightDep = initialParsedCellDependency(rightAddress, ownerRow, rightCol)
    return {
      leftColOffset: shape.leftColOffset,
      operator: shape.operator,
      templateKey: shape.templateKey,
      right: shape.right,
      symbolicRefs: [leftAddress, rightAddress],
      parsedDeps: [leftDep, rightDep],
      parsedSymbolicRefs: [leftRef, rightRef],
    }
  }
  if (!initialSourceStartsWith(source, shape.right.text, index) || index + shape.right.text.length !== source.length) {
    return undefined
  }
  return {
    leftColOffset: shape.leftColOffset,
    operator: shape.operator,
    templateKey: shape.templateKey,
    right: shape.right,
    symbolicRefs: [leftAddress],
    parsedDeps: [leftDep],
    parsedSymbolicRefs: [leftRef],
  }
}
