import { ErrorCode, ValueTag } from './protocol'

const DIRECT_SCALAR_OP_ADD: u8 = 1
const DIRECT_SCALAR_OP_SUB: u8 = 2
const DIRECT_SCALAR_OP_MUL: u8 = 3
const DIRECT_SCALAR_OP_DIV: u8 = 4
const DIRECT_SCALAR_OP_ABS: u8 = 5
const DIRECT_SCALAR_BATCH_REF_NONE: u32 = 0xffffffff

let directScalarOperandTag: u8 = <u8>ValueTag.Empty
let directScalarOperandValue: f64 = 0

function readDirectScalarValueBatchOperand(
  batchRef: u32,
  tag: u8,
  value: f64,
  error: u16,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  let operandTag = tag
  let operandValue = value
  let operandError = error
  if (batchRef != DIRECT_SCALAR_BATCH_REF_NONE) {
    operandTag = outTags[batchRef]
    operandValue = outNumbers[batchRef]
    operandError = outErrors[batchRef]
  }
  if (operandTag == ValueTag.Error) {
    directScalarOperandTag = <u8>ValueTag.Error
    directScalarOperandValue = operandError
    return
  }
  if (operandTag == ValueTag.String) {
    directScalarOperandTag = <u8>ValueTag.Error
    directScalarOperandValue = ErrorCode.Value
    return
  }
  directScalarOperandTag = <u8>ValueTag.Number
  directScalarOperandValue = operandTag == ValueTag.Boolean ? (operandValue != 0 ? 1 : 0) : operandTag == ValueTag.Empty ? 0 : operandValue
}

function writeDirectScalarValueBatchResult(
  index: i32,
  tag: u8,
  value: f64,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  outTags[index] = tag
  if (tag == ValueTag.Error) {
    outNumbers[index] = 0
    outErrors[index] = <u16>value
    return
  }
  outNumbers[index] = value
  outErrors[index] = ErrorCode.None
}

export function evalDirectScalarValueBatch(
  operators: Uint8Array,
  leftBatchRefs: Uint32Array,
  leftTags: Uint8Array,
  leftValues: Float64Array,
  leftErrors: Uint16Array,
  rightBatchRefs: Uint32Array,
  rightTags: Uint8Array,
  rightValues: Float64Array,
  rightErrors: Uint16Array,
  resultOffsets: Float64Array,
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): void {
  for (let index = 0; index < operators.length; index++) {
    const operator = operators[index]
    readDirectScalarValueBatchOperand(
      leftBatchRefs[index],
      leftTags[index],
      leftValues[index],
      leftErrors[index],
      outTags,
      outNumbers,
      outErrors,
    )
    const leftTag = directScalarOperandTag
    const leftValue = directScalarOperandValue
    if (leftTag == ValueTag.Error) {
      writeDirectScalarValueBatchResult(index, leftTag, leftValue, outTags, outNumbers, outErrors)
      continue
    }

    if (operator == DIRECT_SCALAR_OP_ABS) {
      writeDirectScalarValueBatchResult(index, <u8>ValueTag.Number, Math.abs(leftValue), outTags, outNumbers, outErrors)
      continue
    }

    readDirectScalarValueBatchOperand(
      rightBatchRefs[index],
      rightTags[index],
      rightValues[index],
      rightErrors[index],
      outTags,
      outNumbers,
      outErrors,
    )
    const rightTag = directScalarOperandTag
    const rightValue = directScalarOperandValue
    if (rightTag == ValueTag.Error) {
      writeDirectScalarValueBatchResult(index, rightTag, rightValue, outTags, outNumbers, outErrors)
      continue
    }

    let result: f64 = 0
    if (operator == DIRECT_SCALAR_OP_ADD) {
      result = leftValue + rightValue
    } else if (operator == DIRECT_SCALAR_OP_SUB) {
      result = leftValue - rightValue
    } else if (operator == DIRECT_SCALAR_OP_MUL) {
      result = leftValue * rightValue
    } else if (operator == DIRECT_SCALAR_OP_DIV) {
      if (rightValue == 0) {
        writeDirectScalarValueBatchResult(index, <u8>ValueTag.Error, ErrorCode.Div0, outTags, outNumbers, outErrors)
        continue
      }
      result = leftValue / rightValue
    } else {
      writeDirectScalarValueBatchResult(index, <u8>ValueTag.Error, ErrorCode.Value, outTags, outNumbers, outErrors)
      continue
    }
    writeDirectScalarValueBatchResult(index, <u8>ValueTag.Number, result + resultOffsets[index], outTags, outNumbers, outErrors)
  }
}
