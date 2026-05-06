import type { ParsedCellReferenceInfo } from '@bilig/formula'

interface InitialRelativeCellToken {
  readonly colOffset: number
  readonly ref: ParsedCellReferenceInfo
  readonly next: number
}

export interface InitialSimpleRowRelativeBinaryTemplateMatch {
  readonly templateKey: string
  readonly parsedSymbolicRefs: ParsedCellReferenceInfo[]
}

function initialColumnToIndex(column: string): number {
  let value = 0
  for (let index = 0; index < column.length; index += 1) {
    const code = column.charCodeAt(index)
    value = value * 26 + (code >= 97 && code <= 122 ? code - 96 : code - 64)
  }
  return value - 1
}

function initialReadColumn(source: string, start: number): { readonly column: string; readonly next: number } | undefined {
  let next = start
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (!((code >= 65 && code <= 90) || (code >= 97 && code <= 122))) {
      break
    }
    next += 1
  }
  return next === start ? undefined : { column: source.slice(start, next), next }
}

function initialReadRowNumber(source: string, start: number): { readonly row: number; readonly next: number } | undefined {
  let next = start
  let row = 0
  while (next < source.length) {
    const code = source.charCodeAt(next)
    if (code < 48 || code > 57) {
      break
    }
    row = row * 10 + (code - 48)
    next += 1
  }
  return next === start || row <= 0 ? undefined : { row, next }
}

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

function initialReadRelativeCellToken(
  source: string,
  start: number,
  ownerRow: number,
  ownerCol: number,
): InitialRelativeCellToken | undefined {
  const column = initialReadColumn(source, start)
  if (!column) {
    return undefined
  }
  const row = initialReadRowNumber(source, column.next)
  if (!row || row.row - 1 !== ownerRow) {
    return undefined
  }
  const col = initialColumnToIndex(column.column)
  return col < 0
    ? undefined
    : {
        colOffset: col - ownerCol,
        ref: {
          address: `${column.column.toUpperCase()}${String(row.row)}`,
          row: row.row - 1,
          col,
        },
        next: row.next,
      }
}

export function tryMatchInitialSimpleRowRelativeBinaryTemplate(
  source: string,
  ownerRow: number,
  ownerCol: number,
): InitialSimpleRowRelativeBinaryTemplateMatch | undefined {
  let index = source.charCodeAt(0) === 61 ? 1 : 0
  const left = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (!left) {
    return undefined
  }
  index = left.next
  const operator = source[index]
  if (operator !== '+' && operator !== '-' && operator !== '*' && operator !== '/') {
    return undefined
  }
  index += 1
  const rightCell = initialReadRelativeCellToken(source, index, ownerRow, ownerCol)
  if (rightCell) {
    return rightCell.next === source.length
      ? {
          templateKey: `c${left.colOffset}${operator}c${rightCell.colOffset}`,
          parsedSymbolicRefs: [left.ref, rightCell.ref],
        }
      : undefined
  }
  const rightNumber = initialReadNumberLiteral(source, index)
  return rightNumber && rightNumber.next === source.length
    ? {
        templateKey: `c${left.colOffset}${operator}n${rightNumber.text}`,
        parsedSymbolicRefs: [left.ref],
      }
    : undefined
}
