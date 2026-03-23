import { applyBuiltin, STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR } from "./builtins";
import { ErrorCode, Opcode, ValueTag } from "./protocol";

export let tags = new Uint8Array(64);
export let numbers = new Float64Array(64);
export let stringIds = new Uint32Array(64);
export let errors = new Uint16Array(64);
let stringLengths = new Uint32Array(64);
let stringOffsets = new Uint32Array(64);
let stringData = new Uint16Array(64);

let programArena = new Uint32Array(64);
let programOffsets = new Uint32Array(64);
let programLengths = new Uint32Array(64);
let programTargets = new Uint32Array(64);
let constantArena = new Float64Array(64);
let constantOffsets = new Uint32Array(64);
let constantLengths = new Uint32Array(64);
export let rangeOffsets = new Uint32Array(64);
export let rangeLengths = new Uint32Array(64);
export let rangeRowCounts = new Uint32Array(64);
export let rangeColCounts = new Uint32Array(64);
export let rangeMembers = new Uint32Array(64);
let formulaForCell = new Uint32Array(64);
let formulaCount = 0;

let outputStringLengths = new Uint32Array(64);
let outputStringOffsets = new Uint32Array(64);
let outputStringData = new Uint16Array(64);
let outputStringCount = 0;
let outputStringDataLength = 0;
let spillRows = new Uint32Array(64);
let spillCols = new Uint32Array(64);
let spillOffsets = new Uint32Array(64);
let spillLengths = new Uint32Array(64);
let spillArrayRows = new Uint32Array(16);
let spillArrayCols = new Uint32Array(16);
let spillArrayOffsets = new Uint32Array(16);
let spillArrayLengths = new Uint32Array(16);
let spillArrayCount = 0;
let spillNumbers = new Float64Array(64);
let spillValueCount = 0;
const OUTPUT_STRING_BASE: f64 = 2147483648.0;
let volatileNowSerial: f64 = NaN;
let volatileRandomValues = new Float64Array(0);
let volatileRandomCursor = 0;
const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff;

export function getOutputStringLengthsPtr(): usize {
  return outputStringLengths.dataStart;
}
export function getOutputStringOffsetsPtr(): usize {
  return outputStringOffsets.dataStart;
}
export function getOutputStringDataPtr(): usize {
  return outputStringData.dataStart;
}
export function getOutputStringCount(): i32 {
  return outputStringCount;
}
export function getOutputStringDataLength(): i32 {
  return outputStringDataLength;
}
export function getSpillResultRowsPtr(): usize {
  return spillRows.dataStart;
}
export function getSpillResultColsPtr(): usize {
  return spillCols.dataStart;
}
export function getSpillResultOffsetsPtr(): usize {
  return spillOffsets.dataStart;
}
export function getSpillResultLengthsPtr(): usize {
  return spillLengths.dataStart;
}
export function getSpillResultNumbersPtr(): usize {
  return spillNumbers.dataStart;
}
export function getSpillResultValueCount(): i32 {
  return spillValueCount;
}

export function resetOutputStrings(): void {
  outputStringCount = 0;
  outputStringDataLength = 0;
}

function resetSpillResults(): void {
  spillArrayCount = 0;
  spillValueCount = 0;
}

function ensureSpillArrayCapacity(nextCapacity: i32): void {
  spillArrayRows = ensureU32(spillArrayRows, nextCapacity);
  spillArrayCols = ensureU32(spillArrayCols, nextCapacity);
  spillArrayOffsets = ensureU32(spillArrayOffsets, nextCapacity);
  spillArrayLengths = ensureU32(spillArrayLengths, nextCapacity);
}

function ensureSpillValueCapacity(nextCapacity: i32): void {
  spillNumbers = ensureF64(spillNumbers, nextCapacity);
}

export function allocateSpillArrayResult(rows: i32, cols: i32): u32 {
  const index = spillArrayCount;
  spillArrayCount += 1;
  ensureSpillArrayCapacity(spillArrayCount);
  const length = rows * cols;
  spillArrayRows[index] = rows;
  spillArrayCols[index] = cols;
  spillArrayOffsets[index] = spillValueCount;
  spillArrayLengths[index] = length;
  spillValueCount += length;
  ensureSpillValueCapacity(spillValueCount);
  return index;
}

export function writeSpillArrayNumber(arrayIndex: u32, offset: i32, value: f64): void {
  const baseOffset = spillArrayOffsets[arrayIndex];
  spillNumbers[baseOffset + offset] = value;
}

export function readSpillArrayLength(arrayIndex: u32): i32 {
  return <i32>spillArrayLengths[arrayIndex];
}

export function readSpillArrayNumber(arrayIndex: u32, offset: i32): f64 {
  return spillNumbers[spillArrayOffsets[arrayIndex] + offset];
}

export function allocateOutputString(length: i32): i32 {
  const index = outputStringCount;
  outputStringCount += 1;
  outputStringLengths = ensureU32(outputStringLengths, outputStringCount);
  outputStringOffsets = ensureU32(outputStringOffsets, outputStringCount);

  outputStringLengths[index] = length;
  outputStringOffsets[index] = outputStringDataLength;

  outputStringDataLength += length;
  outputStringData = ensureU16(outputStringData, outputStringDataLength);

  return index;
}

export function writeOutputStringData(index: i32, offset: i32, char: u16): void {
  const dataOffset = outputStringOffsets[index];
  outputStringData[dataOffset + offset] = char;
}

export function encodeOutputStringId(index: i32): f64 {
  return OUTPUT_STRING_BASE + <f64>index;
}

export function uploadVolatileNowSerial(nowSerial: f64): void {
  volatileNowSerial = nowSerial;
}

export function readVolatileNowSerial(): f64 {
  return volatileNowSerial;
}

export function uploadVolatileRandomValues(values: Float64Array): void {
  volatileRandomValues = values;
  volatileRandomCursor = 0;
}

export function nextVolatileRandomValue(): f64 {
  if (volatileRandomCursor >= volatileRandomValues.length) {
    return NaN;
  }
  const next = volatileRandomValues[volatileRandomCursor];
  volatileRandomCursor += 1;
  return next;
}

const valueStack = new Float64Array(256);
const tagStack = new Uint8Array(256);
const kindStack = new Uint8Array(256);
const rangeIndexStack = new Uint32Array(256);

function toNumeric(kind: u8, tag: u8, value: f64): f64 {
  if (kind == STACK_KIND_RANGE) return NaN;
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function outputStringIndex(value: f64): i32 {
  if (value < OUTPUT_STRING_BASE) {
    return -1;
  }
  return <i32>(value - OUTPUT_STRING_BASE);
}

function isTextLike(tag: u8): bool {
  return tag == ValueTag.String || tag == ValueTag.Empty;
}

function poolString(stringId: i32): string | null {
  if (stringId < 0 || stringId >= stringLengths.length) {
    return null;
  }
  const offset = <i32>stringOffsets[stringId];
  const length = <i32>stringLengths[stringId];
  let text = "";
  for (let index = 0; index < length; index++) {
    text += String.fromCharCode(stringData[offset + index]);
  }
  return text;
}

function outputString(index: i32): string | null {
  if (index < 0 || index >= outputStringLengths.length) {
    return null;
  }
  const offset = <i32>outputStringOffsets[index];
  const length = <i32>outputStringLengths[index];
  let text = "";
  for (let i = 0; i < length; i++) {
    text += String.fromCharCode(outputStringData[offset + i]);
  }
  return text;
}

function scalarText(tag: u8, value: f64): string | null {
  if (tag == ValueTag.Empty) {
    return "";
  }
  if (tag == ValueTag.Number) {
    return value.toString();
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? "TRUE" : "FALSE";
  }
  if (tag == ValueTag.String) {
    const outputIndex = outputStringIndex(value);
    if (outputIndex >= 0) {
      return outputString(outputIndex);
    }
    return poolString(<i32>value);
  }
  return null;
}

function compareText(left: string, right: string): i32 {
  const normalizedLeft = left.toUpperCase();
  const normalizedRight = right.toUpperCase();
  if (normalizedLeft == normalizedRight) {
    return 0;
  }
  return normalizedLeft < normalizedRight ? -1 : 1;
}

function compareScalars(leftTag: u8, leftValue: f64, rightTag: u8, rightValue: f64): i32 {
  if (isTextLike(leftTag) && isTextLike(rightTag)) {
    const leftText = scalarText(leftTag, leftValue);
    const rightText = scalarText(rightTag, rightValue);
    if (leftText == null || rightText == null) {
      return i32.MIN_VALUE;
    }
    return compareText(leftText, rightText);
  }

  const leftNumeric = toNumeric(STACK_KIND_SCALAR, leftTag, leftValue);
  const rightNumeric = toNumeric(STACK_KIND_SCALAR, rightTag, rightValue);
  if (isNaN(leftNumeric) || isNaN(rightNumeric)) {
    return i32.MIN_VALUE;
  }
  if (leftNumeric == rightNumeric) {
    return 0;
  }
  return leftNumeric < rightNumeric ? -1 : 1;
}

function writeConcatenatedString(
  slot: i32,
  leftTag: u8,
  leftValue: f64,
  rightTag: u8,
  rightValue: f64,
): void {
  const leftText = scalarText(leftTag, leftValue);
  const rightText = scalarText(rightTag, rightValue);
  if (leftText == null || rightText == null) {
    writeScalar(slot, <u8>ValueTag.Error, ErrorCode.Value);
    return;
  }
  const outputIndex = allocateOutputString(leftText.length + rightText.length);
  let offset = 0;
  for (let index = 0; index < leftText.length; index++) {
    writeOutputStringData(outputIndex, offset++, <u16>leftText.charCodeAt(index));
  }
  for (let index = 0; index < rightText.length; index++) {
    writeOutputStringData(outputIndex, offset++, <u16>rightText.charCodeAt(index));
  }
  writeScalar(slot, <u8>ValueTag.String, encodeOutputStringId(outputIndex));
}

function ensureU8(buffer: Uint8Array, size: i32): Uint8Array {
  if (buffer.length >= size) return buffer;
  let nextLength = buffer.length;
  while (nextLength < size) nextLength *= 2;
  const next = new Uint8Array(nextLength);
  next.set(buffer);
  return next;
}

function ensureU16(buffer: Uint16Array, size: i32): Uint16Array {
  if (buffer.length >= size) return buffer;
  let nextLength = buffer.length;
  while (nextLength < size) nextLength *= 2;
  const next = new Uint16Array(nextLength);
  next.set(buffer);
  return next;
}

function ensureU32(buffer: Uint32Array, size: i32): Uint32Array {
  if (buffer.length >= size) return buffer;
  let nextLength = buffer.length;
  while (nextLength < size) nextLength *= 2;
  const next = new Uint32Array(nextLength);
  next.set(buffer);
  return next;
}

function ensureF64(buffer: Float64Array, size: i32): Float64Array {
  if (buffer.length >= size) return buffer;
  let nextLength = buffer.length;
  while (nextLength < size) nextLength *= 2;
  const next = new Float64Array(nextLength);
  next.set(buffer);
  return next;
}

export function init(
  cellCapacity: i32,
  formulaCapacity: i32,
  constantCapacity: i32,
  rangeCapacity: i32,
  memberCapacity: i32,
): void {
  ensureCellCapacity(cellCapacity);
  ensureFormulaCapacity(formulaCapacity);
  ensureConstantCapacity(constantCapacity);
  ensureRangeCapacity(rangeCapacity);
  ensureMemberCapacity(memberCapacity);
}

export function ensureCellCapacity(nextCapacity: i32): void {
  tags = ensureU8(tags, nextCapacity);
  numbers = ensureF64(numbers, nextCapacity);
  stringIds = ensureU32(stringIds, nextCapacity);
  errors = ensureU16(errors, nextCapacity);
  formulaForCell = ensureU32(formulaForCell, nextCapacity);
  spillRows = ensureU32(spillRows, nextCapacity);
  spillCols = ensureU32(spillCols, nextCapacity);
  spillOffsets = ensureU32(spillOffsets, nextCapacity);
  spillLengths = ensureU32(spillLengths, nextCapacity);
}

export function ensureFormulaCapacity(nextCapacity: i32): void {
  programOffsets = ensureU32(programOffsets, nextCapacity);
  programLengths = ensureU32(programLengths, nextCapacity);
  programTargets = ensureU32(programTargets, nextCapacity);
  constantOffsets = ensureU32(constantOffsets, nextCapacity);
  constantLengths = ensureU32(constantLengths, nextCapacity);
}

export function ensureConstantCapacity(nextCapacity: i32): void {
  constantArena = ensureF64(constantArena, nextCapacity);
}

export function ensureRangeCapacity(nextCapacity: i32): void {
  rangeOffsets = ensureU32(rangeOffsets, nextCapacity);
  rangeLengths = ensureU32(rangeLengths, nextCapacity);
  rangeRowCounts = ensureU32(rangeRowCounts, nextCapacity);
  rangeColCounts = ensureU32(rangeColCounts, nextCapacity);
}

export function ensureMemberCapacity(nextCapacity: i32): void {
  rangeMembers = ensureU32(rangeMembers, nextCapacity);
}

function ensureStringCapacity(nextCapacity: i32): void {
  stringLengths = ensureU32(stringLengths, nextCapacity);
  stringOffsets = ensureU32(stringOffsets, nextCapacity);
}

function ensureStringDataCapacity(nextCapacity: i32): void {
  stringData = ensureU16(stringData, nextCapacity);
}

export function uploadPrograms(
  programs: Uint32Array,
  offsets: Uint32Array,
  lengths: Uint32Array,
  targets: Uint32Array,
): void {
  programArena = ensureU32(programArena, programs.length);
  programArena.set(programs);
  formulaCount = offsets.length;
  ensureFormulaCapacity(formulaCount);
  formulaForCell.fill(0);
  for (let index = 0; index < formulaCount; index++) {
    programOffsets[index] = offsets[index];
    programLengths[index] = lengths[index];
    programTargets[index] = targets[index];
    if (<i32>targets[index] >= formulaForCell.length) {
      ensureCellCapacity(<i32>targets[index] + 1);
    }
    formulaForCell[targets[index]] = index + 1;
  }
}

export function uploadConstants(
  constants: Float64Array,
  offsets: Uint32Array,
  lengths: Uint32Array,
): void {
  constantArena = ensureF64(constantArena, constants.length);
  constantArena.set(constants);
  ensureFormulaCapacity(offsets.length);
  for (let index = 0; index < offsets.length; index++) {
    constantOffsets[index] = offsets[index];
    constantLengths[index] = lengths[index];
  }
}

export function uploadRangeMembers(
  members: Uint32Array,
  offsets: Uint32Array,
  lengths: Uint32Array,
): void {
  ensureRangeCapacity(offsets.length);
  ensureMemberCapacity(members.length);
  rangeMembers.set(members);
  rangeOffsets.fill(0);
  rangeLengths.fill(0);
  for (let index = 0; index < offsets.length; index++) {
    rangeOffsets[index] = offsets[index];
    rangeLengths[index] = lengths[index];
  }
}

export function uploadRangeShapes(rowCounts: Uint32Array, colCounts: Uint32Array): void {
  ensureRangeCapacity(max<i32>(rowCounts.length, colCounts.length));
  rangeRowCounts.fill(0);
  rangeColCounts.fill(0);
  for (let index = 0; index < rowCounts.length; index++) {
    rangeRowCounts[index] = rowCounts[index];
  }
  for (let index = 0; index < colCounts.length; index++) {
    rangeColCounts[index] = colCounts[index];
  }
}

export function uploadStringLengths(lengths: Uint32Array): void {
  ensureStringCapacity(lengths.length);
  stringLengths.set(lengths);
}

export function uploadStrings(offsets: Uint32Array, lengths: Uint32Array, data: Uint16Array): void {
  ensureStringCapacity(lengths.length);
  ensureStringDataCapacity(data.length);
  stringOffsets.set(offsets);
  stringLengths.set(lengths);
  stringData.set(data);
}

export function writeCells(
  nextTags: Uint8Array,
  nextNumbers: Float64Array,
  nextStringIds: Uint32Array,
  nextErrors: Uint16Array,
): void {
  ensureCellCapacity(nextTags.length);
  tags.set(nextTags);
  numbers.set(nextNumbers);
  stringIds.set(nextStringIds);
  errors.set(nextErrors);
}

function binaryNumeric(op: i32, left: f64, right: f64): f64 {
  if (op == Opcode.Add) return left + right;
  if (op == Opcode.Sub) return left - right;
  if (op == Opcode.Mul) return left * right;
  if (op == Opcode.Div) return right == 0 ? NaN : left / right;
  if (op == Opcode.Pow) return Math.pow(left, right);
  if (op == Opcode.Eq) return left == right ? 1 : 0;
  if (op == Opcode.Neq) return left != right ? 1 : 0;
  if (op == Opcode.Gt) return left > right ? 1 : 0;
  if (op == Opcode.Gte) return left >= right ? 1 : 0;
  if (op == Opcode.Lt) return left < right ? 1 : 0;
  if (op == Opcode.Lte) return left <= right ? 1 : 0;
  return NaN;
}

function writeScalar(slot: i32, tag: u8, value: f64): void {
  kindStack[slot] = STACK_KIND_SCALAR;
  rangeIndexStack[slot] = 0;
  tagStack[slot] = tag;
  valueStack[slot] = value;
}

function evalProgram(cellIndex: i32, formulaIndex: i32): void {
  let sp = 0;
  const start = programOffsets[formulaIndex];
  const end = start + programLengths[formulaIndex];
  const constantBase = constantOffsets[formulaIndex];

  for (let pc = start; pc < end; pc++) {
    const instruction = programArena[pc];
    const opcode = instruction >>> 24;
    const operand = instruction & 0x00ffffff;

    if (opcode == Opcode.PushNumber) {
      writeScalar(sp, <u8>ValueTag.Number, constantArena[constantBase + operand]);
      sp++;
      continue;
    }

    if (opcode == Opcode.PushBoolean) {
      writeScalar(sp, <u8>ValueTag.Boolean, operand == 0 ? 0 : 1);
      sp++;
      continue;
    }

    if (opcode == Opcode.PushString) {
      writeScalar(sp, <u8>ValueTag.String, <f64>operand);
      sp++;
      continue;
    }

    if (opcode == Opcode.PushError) {
      writeScalar(sp, <u8>ValueTag.Error, operand);
      sp++;
      continue;
    }

    if (opcode == Opcode.PushCell) {
      if (operand == UNRESOLVED_WASM_OPERAND) {
        writeScalar(sp, <u8>ValueTag.Error, ErrorCode.Ref);
        sp++;
        continue;
      }
      if (tags[operand] == ValueTag.Error) {
        writeScalar(sp, <u8>ValueTag.Error, errors[operand]);
      } else if (tags[operand] == ValueTag.String) {
        writeScalar(sp, <u8>ValueTag.String, stringIds[operand]);
      } else {
        writeScalar(sp, tags[operand], numbers[operand]);
      }
      sp++;
      continue;
    }

    if (opcode == Opcode.PushRange) {
      kindStack[sp] = STACK_KIND_RANGE;
      rangeIndexStack[sp] = operand;
      tagStack[sp] = ValueTag.Empty;
      valueStack[sp] = 0;
      sp++;
      continue;
    }

    if (
      opcode == Opcode.Add ||
      opcode == Opcode.Sub ||
      opcode == Opcode.Mul ||
      opcode == Opcode.Div ||
      opcode == Opcode.Pow ||
      opcode == Opcode.Concat ||
      opcode == Opcode.Eq ||
      opcode == Opcode.Neq ||
      opcode == Opcode.Gt ||
      opcode == Opcode.Gte ||
      opcode == Opcode.Lt ||
      opcode == Opcode.Lte
    ) {
      const rightTag = tagStack[sp - 1];
      const leftTag = tagStack[sp - 2];
      const rightKind = kindStack[sp - 1];
      const leftKind = kindStack[sp - 2];
      if (rightKind == STACK_KIND_RANGE || leftKind == STACK_KIND_RANGE) {
        writeScalar(sp - 2, <u8>ValueTag.Error, ErrorCode.Value);
        sp--;
        continue;
      }
      if (rightTag == ValueTag.Error || leftTag == ValueTag.Error) {
        writeScalar(
          sp - 2,
          <u8>ValueTag.Error,
          rightTag == ValueTag.Error ? valueStack[sp - 1] : valueStack[sp - 2],
        );
        sp--;
        continue;
      }
      if (opcode == Opcode.Concat) {
        writeConcatenatedString(sp - 2, leftTag, valueStack[sp - 2], rightTag, valueStack[sp - 1]);
        sp--;
        continue;
      }
      if (
        opcode == Opcode.Eq ||
        opcode == Opcode.Neq ||
        opcode == Opcode.Gt ||
        opcode == Opcode.Gte ||
        opcode == Opcode.Lt ||
        opcode == Opcode.Lte
      ) {
        const comparison = compareScalars(
          leftTag,
          valueStack[sp - 2],
          rightTag,
          valueStack[sp - 1],
        );
        if (comparison == i32.MIN_VALUE) {
          writeScalar(sp - 2, <u8>ValueTag.Error, ErrorCode.Value);
          sp--;
          continue;
        }
        let result = 0;
        if (opcode == Opcode.Eq) result = comparison == 0 ? 1 : 0;
        else if (opcode == Opcode.Neq) result = comparison != 0 ? 1 : 0;
        else if (opcode == Opcode.Gt) result = comparison > 0 ? 1 : 0;
        else if (opcode == Opcode.Gte) result = comparison >= 0 ? 1 : 0;
        else if (opcode == Opcode.Lt) result = comparison < 0 ? 1 : 0;
        else result = comparison <= 0 ? 1 : 0;
        writeScalar(sp - 2, <u8>ValueTag.Boolean, result);
        sp--;
        continue;
      }
      const left = toNumeric(leftKind, leftTag, valueStack[sp - 2]);
      const right = toNumeric(rightKind, rightTag, valueStack[sp - 1]);
      if (isNaN(left) || isNaN(right)) {
        writeScalar(sp - 2, <u8>ValueTag.Error, ErrorCode.Value);
        sp--;
        continue;
      }
      const next = binaryNumeric(opcode, left, right);
      if (opcode == Opcode.Div && right == 0) {
        writeScalar(sp - 2, <u8>ValueTag.Error, ErrorCode.Div0);
      } else {
        writeScalar(
          sp - 2,
          <u8>(
            (<i32>opcode >= Opcode.Eq && <i32>opcode <= Opcode.Lte
              ? ValueTag.Boolean
              : ValueTag.Number)
          ),
          next,
        );
      }
      sp--;
      continue;
    }

    if (opcode == Opcode.Neg) {
      if (kindStack[sp - 1] == STACK_KIND_RANGE) {
        writeScalar(sp - 1, <u8>ValueTag.Error, ErrorCode.Value);
        continue;
      }
      if (tagStack[sp - 1] == ValueTag.Error) {
        continue;
      }
      const numeric = toNumeric(kindStack[sp - 1], tagStack[sp - 1], valueStack[sp - 1]);
      if (isNaN(numeric)) {
        writeScalar(sp - 1, <u8>ValueTag.Error, ErrorCode.Value);
      } else {
        writeScalar(sp - 1, <u8>ValueTag.Number, -numeric);
      }
      continue;
    }

    if (opcode == Opcode.Jump) {
      pc = start + operand - 1;
      continue;
    }

    if (opcode == Opcode.JumpIfFalse) {
      if (kindStack[sp - 1] == STACK_KIND_SCALAR && tagStack[sp - 1] == ValueTag.Error) {
        const skipInstruction = programArena[start + operand - 1];
        if (skipInstruction >>> 24 == Opcode.Jump) {
          pc = start + (skipInstruction & 0x00ffffff) - 1;
          continue;
        }
      }
      const numeric = toNumeric(kindStack[sp - 1], tagStack[sp - 1], valueStack[sp - 1]);
      sp--;
      if (isNaN(numeric) || numeric == 0) {
        pc = start + operand - 1;
      }
      continue;
    }

    if (opcode == Opcode.CallBuiltin) {
      const builtinId = operand >>> 8;
      const argc = operand & 0xff;
      sp = applyBuiltin(
        builtinId,
        argc,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
        tags,
        numbers,
        stringIds,
        errors,
        stringOffsets,
        stringLengths,
        stringData,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
        sp,
      );
      continue;
    }

    if (opcode == Opcode.Ret) {
      if (kindStack[sp - 1] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[sp - 1];
        const offset = spillArrayOffsets[arrayIndex];
        const length = spillArrayLengths[arrayIndex];
        spillRows[cellIndex] = spillArrayRows[arrayIndex];
        spillCols[cellIndex] = spillArrayCols[arrayIndex];
        spillOffsets[cellIndex] = offset;
        spillLengths[cellIndex] = length;
        tags[cellIndex] = length > 0 ? <u8>ValueTag.Number : <u8>ValueTag.Empty;
        numbers[cellIndex] = length > 0 ? spillNumbers[offset] : 0;
        stringIds[cellIndex] = 0;
        errors[cellIndex] = ErrorCode.None;
        return;
      }
      const resultTag = tagStack[sp - 1];
      tags[cellIndex] = resultTag;
      if (resultTag == ValueTag.String) {
        const stringValue = valueStack[sp - 1];
        stringIds[cellIndex] =
          stringValue >= OUTPUT_STRING_BASE
            ? 0x80000000 | <u32>(stringValue - OUTPUT_STRING_BASE)
            : <u32>stringValue;
      } else {
        stringIds[cellIndex] = 0;
      }
      if (resultTag == ValueTag.Error) {
        errors[cellIndex] = <u16>valueStack[sp - 1];
        numbers[cellIndex] = 0;
      } else if (resultTag == ValueTag.String) {
        numbers[cellIndex] = 0;
        errors[cellIndex] = ErrorCode.None;
      } else {
        numbers[cellIndex] = valueStack[sp - 1];
        errors[cellIndex] = ErrorCode.None;
      }
      return;
    }
  }
}

export function evalBatch(cellIndices: Uint32Array): void {
  resetOutputStrings();
  resetSpillResults();
  for (let index = 0; index < cellIndices.length; index++) {
    const cellIndex = cellIndices[index];
    spillRows[cellIndex] = 0;
    spillCols[cellIndex] = 0;
    spillOffsets[cellIndex] = 0;
    spillLengths[cellIndex] = 0;
    const formulaIndex = formulaForCell[cellIndex];
    if (formulaIndex == 0) continue;
    evalProgram(cellIndex, formulaIndex - 1);
  }
}

export function getTagsPtr(): usize {
  return changetype<usize>(tags.dataStart);
}

export function getNumbersPtr(): usize {
  return changetype<usize>(numbers.dataStart);
}

export function getStringIdsPtr(): usize {
  return changetype<usize>(stringIds.dataStart);
}

export function getErrorsPtr(): usize {
  return changetype<usize>(errors.dataStart);
}

export function getProgramOffsetsPtr(): usize {
  return changetype<usize>(programOffsets.dataStart);
}

export function getProgramLengthsPtr(): usize {
  return changetype<usize>(programLengths.dataStart);
}

export function getConstantOffsetsPtr(): usize {
  return changetype<usize>(constantOffsets.dataStart);
}

export function getConstantLengthsPtr(): usize {
  return changetype<usize>(constantLengths.dataStart);
}

export function getConstantArenaPtr(): usize {
  return changetype<usize>(constantArena.dataStart);
}

export function getRangeOffsetsPtr(): usize {
  return changetype<usize>(rangeOffsets.dataStart);
}

export function getRangeLengthsPtr(): usize {
  return changetype<usize>(rangeLengths.dataStart);
}

export function getRangeMembersPtr(): usize {
  return changetype<usize>(rangeMembers.dataStart);
}

export function getCellCapacity(): i32 {
  return tags.length;
}

export function getFormulaCapacity(): i32 {
  return programOffsets.length;
}

export function getConstantCapacity(): i32 {
  return constantArena.length;
}

export function getRangeCapacity(): i32 {
  return rangeOffsets.length;
}

export function getMemberCapacity(): i32 {
  return rangeMembers.length;
}
