import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  inputCellScalarValue,
  inputCellTag,
  inputColsFromSlot,
  inputRowsFromSlot,
  toNumberOrNaN,
  toNumberOrZero,
} from "./operands";
import { coerceNumberArg, coercePositiveIntegerArg } from "./builtin-args";
import {
  STACK_KIND_ARRAY,
  STACK_KIND_RANGE,
  STACK_KIND_SCALAR,
  writeArrayResult,
  writeResult,
} from "./result-io";
import { allocateSpillArrayResult, writeSpillArrayNumber } from "./vm";

const AXIS_AGG_SUM: i32 = 1;
const AXIS_AGG_AVERAGE: i32 = 2;
const AXIS_AGG_MIN: i32 = 3;
const AXIS_AGG_MAX: i32 = 4;
const AXIS_AGG_COUNT: i32 = 5;
const AXIS_AGG_COUNTA: i32 = 6;

export function tryApplyArrayFoundationBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): i32 {
  if (builtinId == BuiltinId.Sequence) {
    if (argc < 1 || argc > 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rows = coercePositiveIntegerArg(tagStack[base], valueStack[base], argc >= 1, 1);
    const cols = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], argc >= 2, 1);
    const start = coerceNumberArg(tagStack[base + 2], valueStack[base + 2], argc >= 3, 1);
    const step = coerceNumberArg(tagStack[base + 3], valueStack[base + 3], argc >= 4, 1);
    if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || isNaN(start) || isNaN(step)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const arrayIndex = allocateSpillArrayResult(rows, cols);
    const length = rows * cols;
    for (let index = 0; index < length; index++) {
      writeSpillArrayNumber(arrayIndex, index, start + <f64>index * step);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rows,
      cols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.ByrowSum || builtinId == BuiltinId.BycolSum) && argc == 1) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const byRow = builtinId == BuiltinId.ByrowSum;
    const outputRows = byRow ? sourceRows : 1;
    const outputCols = byRow ? 1 : sourceCols;
    const outerCount = byRow ? sourceRows : sourceCols;
    const innerCount = byRow ? sourceCols : sourceRows;
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);

    for (let outer = 0; outer < outerCount; outer += 1) {
      let sum = 0.0;
      for (let inner = 0; inner < innerCount; inner += 1) {
        const row = byRow ? outer : inner;
        const col = byRow ? inner : outer;
        const tag = inputCellTag(
          base,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        const scalar = inputCellScalarValue(
          base,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (tag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            scalar,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const numeric = toNumberOrNaN(tag, scalar);
        if (!isNaN(numeric)) {
          sum += numeric;
        }
      }
      writeSpillArrayNumber(arrayIndex, outer, sum);
    }

    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.ByrowAggregate || builtinId == BuiltinId.BycolAggregate) &&
    argc == 2
  ) {
    const aggregateCodeValue = toNumberOrNaN(tagStack[base], valueStack[base]);
    if (isNaN(aggregateCodeValue)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const aggregateCode = <i32>aggregateCodeValue;
    if (
      aggregateCode != AXIS_AGG_SUM &&
      aggregateCode != AXIS_AGG_AVERAGE &&
      aggregateCode != AXIS_AGG_MIN &&
      aggregateCode != AXIS_AGG_MAX &&
      aggregateCode != AXIS_AGG_COUNT &&
      aggregateCode != AXIS_AGG_COUNTA
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceSlot = base + 1;
    const sourceKind = kindStack[sourceSlot];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceRows = inputRowsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const byRow = builtinId == BuiltinId.ByrowAggregate;
    const outputRows = byRow ? sourceRows : 1;
    const outputCols = byRow ? 1 : sourceCols;
    const outerCount = byRow ? sourceRows : sourceCols;
    const innerCount = byRow ? sourceCols : sourceRows;
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);

    for (let outer = 0; outer < outerCount; outer += 1) {
      let numericCount = 0;
      let nonEmptyCount = 0;
      let numericTotal = 0.0;
      let extremum = 0.0;
      let extremumSet = false;
      for (let inner = 0; inner < innerCount; inner += 1) {
        const row = byRow ? outer : inner;
        const col = byRow ? inner : outer;
        const tag = inputCellTag(
          sourceSlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        const scalar = inputCellScalarValue(
          sourceSlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (tag == ValueTag.Error) {
          if (aggregateCode == AXIS_AGG_COUNT || aggregateCode == AXIS_AGG_COUNTA) {
            if (aggregateCode == AXIS_AGG_COUNTA) {
              nonEmptyCount += 1;
            }
            continue;
          }
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            scalar,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        if (tag != ValueTag.Empty) {
          nonEmptyCount += 1;
        }
        const numeric = toNumberOrNaN(tag, scalar);
        if (!isNaN(numeric)) {
          numericTotal += numeric;
          if (!extremumSet) {
            extremum = numeric;
            extremumSet = true;
          } else if (aggregateCode == AXIS_AGG_MIN && numeric < extremum) {
            extremum = numeric;
          } else if (aggregateCode == AXIS_AGG_MAX && numeric > extremum) {
            extremum = numeric;
          }
          if (tag != ValueTag.Empty || aggregateCode != AXIS_AGG_COUNT) {
            numericCount += 1;
          }
        }
      }
      if (aggregateCode == AXIS_AGG_COUNT) {
        writeSpillArrayNumber(arrayIndex, outer, <f64>numericCount);
      } else if (aggregateCode == AXIS_AGG_COUNTA) {
        writeSpillArrayNumber(arrayIndex, outer, <f64>nonEmptyCount);
      } else if (aggregateCode == AXIS_AGG_AVERAGE) {
        writeSpillArrayNumber(
          arrayIndex,
          outer,
          numericCount == 0 ? 0.0 : numericTotal / <f64>numericCount,
        );
      } else if (aggregateCode == AXIS_AGG_MIN || aggregateCode == AXIS_AGG_MAX) {
        writeSpillArrayNumber(arrayIndex, outer, extremumSet ? extremum : 0.0);
      } else {
        writeSpillArrayNumber(arrayIndex, outer, numericTotal);
      }
    }

    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.ReduceSum ||
      builtinId == BuiltinId.ScanSum ||
      builtinId == BuiltinId.ReduceProduct ||
      builtinId == BuiltinId.ScanProduct) &&
    (argc == 1 || argc == 2)
  ) {
    const sourceSlot = base + argc - 1;
    const sourceKind = kindStack[sourceSlot];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceRows = inputRowsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const productFold = builtinId == BuiltinId.ReduceProduct || builtinId == BuiltinId.ScanProduct;
    let accumulator =
      argc == 2 ? toNumberOrZero(tagStack[base], valueStack[base]) : productFold ? 1.0 : 0.0;
    const arrayIndex =
      builtinId == BuiltinId.ScanSum || builtinId == BuiltinId.ScanProduct
        ? allocateSpillArrayResult(sourceRows, sourceCols)
        : 0;
    let outputOffset = 0;

    for (let row = 0; row < sourceRows; row += 1) {
      for (let col = 0; col < sourceCols; col += 1) {
        const tag = inputCellTag(
          sourceSlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        const scalar = inputCellScalarValue(
          sourceSlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (tag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            scalar,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const numeric = toNumberOrNaN(tag, scalar);
        if (!isNaN(numeric)) {
          accumulator = productFold ? accumulator * numeric : accumulator + numeric;
        }
        if (builtinId == BuiltinId.ScanSum || builtinId == BuiltinId.ScanProduct) {
          writeSpillArrayNumber(arrayIndex, outputOffset, accumulator);
          outputOffset += 1;
        }
      }
    }

    if (builtinId == BuiltinId.ScanSum || builtinId == BuiltinId.ScanProduct) {
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      accumulator,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.MakearraySum && argc == 2) {
    const rows = coercePositiveIntegerArg(tagStack[base], valueStack[base], true, 1);
    const cols = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], true, 1);
    if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const arrayIndex = allocateSpillArrayResult(rows, cols);
    let outputOffset = 0;
    for (let row = 1; row <= rows; row += 1) {
      for (let col = 1; col <= cols; col += 1) {
        writeSpillArrayNumber(arrayIndex, outputOffset, <f64>(row + col));
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rows,
      cols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return -1;
}
