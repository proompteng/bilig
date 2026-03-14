import { applyBuiltin, STACK_KIND_RANGE, STACK_KIND_SCALAR } from "./builtins";
import { ErrorCode, Opcode, ValueTag } from "./protocol";

export let tags = new Uint8Array(64);
export let numbers = new Float64Array(64);
export let stringIds = new Uint32Array(64);
export let errors = new Uint16Array(64);

let programArena = new Uint32Array(64);
let programOffsets = new Uint32Array(64);
let programLengths = new Uint32Array(64);
let programTargets = new Uint32Array(64);
let constantArena = new Float64Array(64);
let constantOffsets = new Uint32Array(64);
let constantLengths = new Uint32Array(64);
let rangeOffsets = new Uint32Array(64);
let rangeLengths = new Uint32Array(64);
let rangeMembers = new Uint32Array(64);
let formulaForCell = new Uint32Array(64);
let formulaCount = 0;

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
  memberCapacity: i32
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
}

export function ensureMemberCapacity(nextCapacity: i32): void {
  rangeMembers = ensureU32(rangeMembers, nextCapacity);
}

export function uploadPrograms(
  programs: Uint32Array,
  offsets: Uint32Array,
  lengths: Uint32Array,
  targets: Uint32Array
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

export function uploadConstants(constants: Float64Array, offsets: Uint32Array, lengths: Uint32Array): void {
  constantArena = ensureF64(constantArena, constants.length);
  constantArena.set(constants);
  ensureFormulaCapacity(offsets.length);
  for (let index = 0; index < offsets.length; index++) {
    constantOffsets[index] = offsets[index];
    constantLengths[index] = lengths[index];
  }
}

export function uploadRangeMembers(members: Uint32Array, offsets: Uint32Array, lengths: Uint32Array): void {
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

export function writeCells(
  nextTags: Uint8Array,
  nextNumbers: Float64Array,
  nextStringIds: Uint32Array,
  nextErrors: Uint16Array
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

    if (opcode == Opcode.PushCell) {
      if (tags[operand] == ValueTag.Error) {
        writeScalar(sp, <u8>ValueTag.Error, errors[operand]);
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
      opcode == Opcode.Add || opcode == Opcode.Sub || opcode == Opcode.Mul || opcode == Opcode.Div ||
      opcode == Opcode.Pow || opcode == Opcode.Eq || opcode == Opcode.Neq || opcode == Opcode.Gt ||
      opcode == Opcode.Gte || opcode == Opcode.Lt || opcode == Opcode.Lte
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
        writeScalar(sp - 2, <u8>ValueTag.Error, rightTag == ValueTag.Error ? valueStack[sp - 1] : valueStack[sp - 2]);
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
          <u8>(<i32>opcode >= Opcode.Eq && <i32>opcode <= Opcode.Lte ? ValueTag.Boolean : ValueTag.Number),
          next
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
        errors,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        sp
      );
      continue;
    }

    if (opcode == Opcode.Ret) {
      const resultTag = tagStack[sp - 1];
      tags[cellIndex] = resultTag;
      stringIds[cellIndex] = 0;
      if (resultTag == ValueTag.Error) {
        errors[cellIndex] = <u16>valueStack[sp - 1];
        numbers[cellIndex] = 0;
      } else {
        numbers[cellIndex] = valueStack[sp - 1];
        errors[cellIndex] = ErrorCode.None;
      }
      return;
    }
  }
}

export function evalBatch(cellIndices: Uint32Array): void {
  for (let index = 0; index < cellIndices.length; index++) {
    const cellIndex = cellIndices[index];
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
