import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationContext, StackValue } from './js-evaluator.js'
import type { RangeLikeValue } from './runtime-values.js'

interface ArraySpecialCallDeps {
  error: (code: ErrorCode) => CellValue
  emptyValue: () => CellValue
  numberValue: (value: number) => CellValue
  stringValue: (value: string) => CellValue
  stackScalar: (value: CellValue) => StackValue
  toRangeLike: (value: StackValue) => RangeLikeValue
  getRangeCell: (range: RangeLikeValue, row: number, col: number) => CellValue
  getBroadcastShape: (values: readonly StackValue[]) => { rows: number; cols: number } | undefined
  makeArrayStack: (rows: number, cols: number, values: CellValue[]) => StackValue
  applyLambda: (lambdaValue: StackValue, args: StackValue[], context: EvaluationContext) => StackValue
  toPositiveInteger: (value: StackValue | undefined) => number | undefined
  coerceScalarTextArgument: (value: StackValue | undefined) => string | CellValue
  coerceOptionalBooleanArgument: (value: StackValue | undefined, fallback: boolean) => boolean | CellValue
  coerceOptionalMatchModeArgument: (value: StackValue | undefined, fallback: 0 | 1) => 0 | 1 | CellValue
  coerceOptionalPositiveIntegerArgument: (value: StackValue | undefined, fallback: number) => number | CellValue
  coerceOptionalTrimModeArgument: (value: StackValue | undefined, fallback: 0 | 1 | 2 | 3) => 0 | 1 | 2 | 3 | CellValue
  isCellValueError: (value: number | boolean | string | CellValue) => value is CellValue
  isSingleCellValue: (value: StackValue) => CellValue | undefined
}

function indexOfWithMatchMode(text: string, delimiter: string, startIndex: number, matchMode: 0 | 1): number {
  if (matchMode === 1) {
    return text.toLowerCase().indexOf(delimiter.toLowerCase(), startIndex)
  }
  return text.indexOf(delimiter, startIndex)
}

function splitTextByDelimiter(text: string, delimiter: string, matchMode: 0 | 1): string[] {
  if (delimiter === '') {
    return [text]
  }
  const parts: string[] = []
  let cursor = 0
  while (cursor <= text.length) {
    const found = indexOfWithMatchMode(text, delimiter, cursor, matchMode)
    if (found === -1) {
      parts.push(text.slice(cursor))
      break
    }
    parts.push(text.slice(cursor, found))
    cursor = found + delimiter.length
  }
  return parts
}

function isTrimRangeEmptyCell(value: CellValue): boolean {
  return value.tag === ValueTag.Empty
}

export function evaluateArraySpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  deps: ArraySpecialCallDeps,
): StackValue | undefined {
  switch (callee) {
    case 'EXPAND': {
      if (rawArgs.length < 2 || rawArgs.length > 4) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const source = deps.toRangeLike(rawArgs[0]!)
      const rows = deps.coerceOptionalPositiveIntegerArgument(rawArgs[1], source.rows)
      const cols = deps.coerceOptionalPositiveIntegerArgument(rawArgs[2], source.cols)
      if (deps.isCellValueError(rows)) {
        return deps.stackScalar(rows)
      }
      if (deps.isCellValueError(cols)) {
        return deps.stackScalar(cols)
      }
      const padArgument = rawArgs[3]
      const padValue =
        padArgument === undefined
          ? deps.error(ErrorCode.NA)
          : (() => {
              const scalar = deps.isSingleCellValue(padArgument)
              return scalar ?? deps.error(ErrorCode.Value)
            })()
      if (padValue.tag === ValueTag.Error && padArgument !== undefined) {
        const scalar = deps.isSingleCellValue(padArgument)
        if (!scalar) {
          return deps.stackScalar(deps.error(ErrorCode.Value))
        }
      }
      if (rows < source.rows || cols < source.cols) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const values: CellValue[] = []
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < cols; col += 1) {
          values.push(row < source.rows && col < source.cols ? deps.getRangeCell(source, row, col) : padValue)
        }
      }
      return deps.makeArrayStack(rows, cols, values)
    }
    case 'TEXTSPLIT': {
      if (rawArgs.length < 2 || rawArgs.length > 6) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const text = deps.coerceScalarTextArgument(rawArgs[0])
      const columnDelimiter = deps.coerceScalarTextArgument(rawArgs[1])
      const rowDelimiter = rawArgs[2] === undefined ? undefined : deps.coerceScalarTextArgument(rawArgs[2])
      const ignoreEmpty = deps.coerceOptionalBooleanArgument(rawArgs[3], false)
      const matchMode = deps.coerceOptionalMatchModeArgument(rawArgs[4], 0)
      if (deps.isCellValueError(text)) {
        return deps.stackScalar(text)
      }
      if (deps.isCellValueError(columnDelimiter)) {
        return deps.stackScalar(columnDelimiter)
      }
      if (rowDelimiter !== undefined && deps.isCellValueError(rowDelimiter)) {
        return deps.stackScalar(rowDelimiter)
      }
      if (deps.isCellValueError(ignoreEmpty)) {
        return deps.stackScalar(ignoreEmpty)
      }
      if (deps.isCellValueError(matchMode)) {
        return deps.stackScalar(matchMode)
      }
      if (columnDelimiter === '' && rowDelimiter === undefined) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const padArgument = rawArgs[5]
      const padValue =
        padArgument === undefined
          ? deps.error(ErrorCode.NA)
          : (() => {
              const scalar = deps.isSingleCellValue(padArgument)
              return scalar ?? deps.error(ErrorCode.Value)
            })()
      if (padArgument !== undefined && !deps.isSingleCellValue(padArgument)) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }

      const rowSlices = rowDelimiter === undefined || rowDelimiter === '' ? [text] : splitTextByDelimiter(text, rowDelimiter, matchMode)
      const matrix = rowSlices.map((rowSlice) => {
        const parts = columnDelimiter === '' ? [rowSlice] : splitTextByDelimiter(rowSlice, columnDelimiter, matchMode)
        const filtered = ignoreEmpty ? parts.filter((part) => part !== '') : parts
        return filtered.length === 0 ? [] : filtered
      })
      const rows = Math.max(matrix.length, 1)
      const cols = Math.max(1, ...matrix.map((row) => row.length))
      const values: CellValue[] = []
      for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
        const row = matrix[rowIndex] ?? []
        for (let colIndex = 0; colIndex < cols; colIndex += 1) {
          values.push(colIndex < row.length ? deps.stringValue(row[colIndex]!) : padValue)
        }
      }
      return deps.makeArrayStack(rows, cols, values)
    }
    case 'TRIMRANGE': {
      if (rawArgs.length < 1 || rawArgs.length > 3) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const source = deps.toRangeLike(rawArgs[0]!)
      const trimRows = deps.coerceOptionalTrimModeArgument(rawArgs[1], 3)
      const trimCols = deps.coerceOptionalTrimModeArgument(rawArgs[2], 3)
      if (deps.isCellValueError(trimRows)) {
        return deps.stackScalar(trimRows)
      }
      if (deps.isCellValueError(trimCols)) {
        return deps.stackScalar(trimCols)
      }

      let startRow = 0
      let endRow = source.rows - 1
      let startCol = 0
      let endCol = source.cols - 1

      const trimLeadingRows = trimRows === 1 || trimRows === 3
      const trimTrailingRows = trimRows === 2 || trimRows === 3
      const trimLeadingCols = trimCols === 1 || trimCols === 3
      const trimTrailingCols = trimCols === 2 || trimCols === 3

      if (trimLeadingRows) {
        while (startRow <= endRow) {
          let hasNonEmpty = false
          for (let col = 0; col < source.cols; col += 1) {
            if (!isTrimRangeEmptyCell(deps.getRangeCell(source, startRow, col))) {
              hasNonEmpty = true
              break
            }
          }
          if (hasNonEmpty) {
            break
          }
          startRow += 1
        }
      }

      if (trimTrailingRows) {
        while (endRow >= startRow) {
          let hasNonEmpty = false
          for (let col = 0; col < source.cols; col += 1) {
            if (!isTrimRangeEmptyCell(deps.getRangeCell(source, endRow, col))) {
              hasNonEmpty = true
              break
            }
          }
          if (hasNonEmpty) {
            break
          }
          endRow -= 1
        }
      }

      if (startRow > endRow) {
        return deps.makeArrayStack(1, 1, [deps.emptyValue()])
      }

      if (trimLeadingCols) {
        while (startCol <= endCol) {
          let hasNonEmpty = false
          for (let row = startRow; row <= endRow; row += 1) {
            if (!isTrimRangeEmptyCell(deps.getRangeCell(source, row, startCol))) {
              hasNonEmpty = true
              break
            }
          }
          if (hasNonEmpty) {
            break
          }
          startCol += 1
        }
      }

      if (trimTrailingCols) {
        while (endCol >= startCol) {
          let hasNonEmpty = false
          for (let row = startRow; row <= endRow; row += 1) {
            if (!isTrimRangeEmptyCell(deps.getRangeCell(source, row, endCol))) {
              hasNonEmpty = true
              break
            }
          }
          if (hasNonEmpty) {
            break
          }
          endCol -= 1
        }
      }

      if (startCol > endCol) {
        return deps.makeArrayStack(1, 1, [deps.emptyValue()])
      }

      const rows = endRow - startRow + 1
      const cols = endCol - startCol + 1
      const values: CellValue[] = []
      for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          values.push(deps.getRangeCell(source, row, col))
        }
      }
      return deps.makeArrayStack(rows, cols, values)
    }
    case 'MAKEARRAY': {
      if (rawArgs.length !== 3) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rows = deps.toPositiveInteger(rawArgs[0])
      const cols = deps.toPositiveInteger(rawArgs[1])
      if (rows === undefined || cols === undefined) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const lambda = rawArgs[2]!
      const values: CellValue[] = []
      for (let row = 1; row <= rows; row += 1) {
        for (let col = 1; col <= cols; col += 1) {
          const result = deps.applyLambda(
            lambda,
            [deps.stackScalar({ tag: ValueTag.Number, value: row }), deps.stackScalar({ tag: ValueTag.Number, value: col })],
            context,
          )
          const scalar = deps.isSingleCellValue(result)
          if (!scalar) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          values.push(scalar)
        }
      }
      return { kind: 'array', rows, cols, values }
    }
    case 'MAP': {
      if (rawArgs.length < 2) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const lambda = rawArgs[rawArgs.length - 1]!
      const inputs = rawArgs.slice(0, -1)
      const shape = deps.getBroadcastShape(inputs)
      if (!shape) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const ranges = inputs.map(deps.toRangeLike)
      const values: CellValue[] = []
      for (let row = 0; row < shape.rows; row += 1) {
        for (let col = 0; col < shape.cols; col += 1) {
          const lambdaArgs = ranges.map((range) =>
            deps.stackScalar(deps.getRangeCell(range, Math.min(row, range.rows - 1), Math.min(col, range.cols - 1))),
          )
          const result = deps.applyLambda(lambda, lambdaArgs, context)
          const scalar = deps.isSingleCellValue(result)
          if (!scalar) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          values.push(scalar)
        }
      }
      return { kind: 'array', rows: shape.rows, cols: shape.cols, values }
    }
    case 'BYROW':
    case 'BYCOL': {
      if (rawArgs.length !== 2) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const source = deps.toRangeLike(rawArgs[0]!)
      const lambda = rawArgs[1]!
      const values: CellValue[] = []
      if (callee === 'BYROW') {
        for (let row = 0; row < source.rows; row += 1) {
          const rowValues: CellValue[] = []
          for (let col = 0; col < source.cols; col += 1) {
            rowValues.push(deps.getRangeCell(source, row, col))
          }
          const result = deps.applyLambda(
            lambda,
            [{ kind: 'range', values: rowValues, rows: 1, cols: source.cols, refKind: 'cells' }],
            context,
          )
          const scalar = deps.isSingleCellValue(result)
          if (!scalar) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          values.push(scalar)
        }
        return { kind: 'array', rows: source.rows, cols: 1, values }
      }
      for (let col = 0; col < source.cols; col += 1) {
        const colValues: CellValue[] = []
        for (let row = 0; row < source.rows; row += 1) {
          colValues.push(deps.getRangeCell(source, row, col))
        }
        const result = deps.applyLambda(
          lambda,
          [{ kind: 'range', values: colValues, rows: source.rows, cols: 1, refKind: 'cells' }],
          context,
        )
        const scalar = deps.isSingleCellValue(result)
        if (!scalar) {
          return deps.stackScalar(deps.error(ErrorCode.Value))
        }
        values.push(scalar)
      }
      return { kind: 'array', rows: 1, cols: source.cols, values }
    }
    case 'REDUCE':
    case 'SCAN': {
      if (rawArgs.length !== 2 && rawArgs.length !== 3) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const hasInitial = rawArgs.length === 3
      let accumulator = hasInitial ? rawArgs[0]! : deps.stackScalar(deps.emptyValue())
      const source = deps.toRangeLike(rawArgs[hasInitial ? 1 : 0]!)
      const lambda = rawArgs[hasInitial ? 2 : 1]!
      const scanValues: CellValue[] = []
      for (const cell of source.values) {
        accumulator = deps.applyLambda(lambda, [accumulator, deps.stackScalar(cell)], context)
        if (callee === 'SCAN') {
          const scalar = deps.isSingleCellValue(accumulator)
          if (!scalar) {
            return deps.stackScalar(deps.error(ErrorCode.Value))
          }
          scanValues.push(scalar)
        }
      }
      return callee === 'SCAN' ? { kind: 'array', rows: source.rows, cols: source.cols, values: scanValues } : accumulator
    }
    default:
      return undefined
  }
}
