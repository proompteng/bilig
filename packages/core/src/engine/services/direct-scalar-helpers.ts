import type { RuntimeDirectScalarDescriptor } from '../runtime-state.js'

export const ROW_PAIR_LEFT_PLUS_RIGHT = 1
export const ROW_PAIR_LEFT_MINUS_RIGHT = 2
export const ROW_PAIR_RIGHT_MINUS_LEFT = 3
export const ROW_PAIR_LEFT_TIMES_RIGHT = 4
export const ROW_PAIR_LEFT_DIV_RIGHT = 5
export const ROW_PAIR_RIGHT_DIV_LEFT = 6

export function directScalarLiteralNumericValue(value: unknown): number | undefined {
  if (value === null) {
    return 0
  }
  switch (typeof value) {
    case 'number':
      return Object.is(value, -0) ? 0 : value
    case 'boolean':
      return value ? 1 : 0
    case 'string':
    case 'bigint':
    case 'function':
    case 'object':
    case 'symbol':
    case 'undefined':
      return undefined
  }
  return undefined
}

export function singleInputAffineDirectScalar(
  directScalar: RuntimeDirectScalarDescriptor,
  inputCellIndex: number,
): { readonly scale: number; readonly offset: number } | null {
  if (directScalar.kind === 'abs') {
    return null
  }
  const leftIsInput = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === inputCellIndex
  const rightIsInput = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === inputCellIndex
  const leftLiteral = directScalar.left.kind === 'literal-number' ? directScalar.left.value : undefined
  const rightLiteral = directScalar.right.kind === 'literal-number' ? directScalar.right.value : undefined
  if (leftIsInput && rightLiteral !== undefined) {
    const resultOffset = directScalar.resultOffset ?? 0
    switch (directScalar.operator) {
      case '+':
        return { scale: 1, offset: rightLiteral + resultOffset }
      case '-':
        return { scale: 1, offset: -rightLiteral + resultOffset }
      case '*':
        return { scale: rightLiteral, offset: resultOffset }
      case '/':
        return rightLiteral === 0 ? null : { scale: 1 / rightLiteral, offset: resultOffset }
    }
  }
  if (rightIsInput && leftLiteral !== undefined) {
    const resultOffset = directScalar.resultOffset ?? 0
    switch (directScalar.operator) {
      case '+':
        return { scale: 1, offset: leftLiteral + resultOffset }
      case '-':
        return { scale: -1, offset: leftLiteral + resultOffset }
      case '*':
        return { scale: leftLiteral, offset: resultOffset }
      case '/':
        return null
    }
  }
  return null
}

export function rowPairDirectScalarCode(
  directScalar: RuntimeDirectScalarDescriptor,
  leftCellIndex: number,
  rightCellIndex: number,
): number {
  if (directScalar.kind === 'abs' || directScalar.resultOffset !== undefined) {
    return 0
  }
  const leftOperandIsLeft = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === leftCellIndex
  const leftOperandIsRight = directScalar.left.kind === 'cell' && directScalar.left.cellIndex === rightCellIndex
  const rightOperandIsLeft = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === leftCellIndex
  const rightOperandIsRight = directScalar.right.kind === 'cell' && directScalar.right.cellIndex === rightCellIndex
  if (leftOperandIsLeft && rightOperandIsRight) {
    switch (directScalar.operator) {
      case '+':
        return ROW_PAIR_LEFT_PLUS_RIGHT
      case '-':
        return ROW_PAIR_LEFT_MINUS_RIGHT
      case '*':
        return ROW_PAIR_LEFT_TIMES_RIGHT
      case '/':
        return ROW_PAIR_LEFT_DIV_RIGHT
    }
  }
  if (leftOperandIsRight && rightOperandIsLeft) {
    switch (directScalar.operator) {
      case '+':
        return ROW_PAIR_LEFT_PLUS_RIGHT
      case '-':
        return ROW_PAIR_RIGHT_MINUS_LEFT
      case '*':
        return ROW_PAIR_LEFT_TIMES_RIGHT
      case '/':
        return ROW_PAIR_RIGHT_DIV_LEFT
    }
  }
  return 0
}

export function evaluateRowPairDirectScalarCode(code: number, leftValue: number, rightValue: number): number | undefined {
  switch (code) {
    case ROW_PAIR_LEFT_PLUS_RIGHT:
      return leftValue + rightValue
    case ROW_PAIR_LEFT_MINUS_RIGHT:
      return leftValue - rightValue
    case ROW_PAIR_RIGHT_MINUS_LEFT:
      return rightValue - leftValue
    case ROW_PAIR_LEFT_TIMES_RIGHT:
      return leftValue * rightValue
    case ROW_PAIR_LEFT_DIV_RIGHT:
      return rightValue === 0 ? undefined : leftValue / rightValue
    case ROW_PAIR_RIGHT_DIV_LEFT:
      return leftValue === 0 ? undefined : rightValue / leftValue
    default:
      return undefined
  }
}

export function rowPairDirectScalarCodeNeedsZeroGuard(code: number): boolean {
  return code === ROW_PAIR_LEFT_DIV_RIGHT || code === ROW_PAIR_RIGHT_DIV_LEFT
}
