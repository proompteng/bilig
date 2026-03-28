import { ErrorCode, ValueTag } from "./protocol";
import { excelYearPartFromSerial } from "./date-finance";
import {
  fDistributionCdf,
  regularizedUpperGamma,
  standardNormalCdf,
  studentTCdf,
} from "./distributions";
import {
  inputCellNumeric,
  inputCellScalarValue,
  inputCellTag,
  inputColsFromSlot,
  inputRowsFromSlot,
  toNumberExact,
  toNumberOrNaN,
} from "./operands";
import { meanOf, sampleVarianceOf } from "./statistics-core";

const STACK_KIND_SCALAR: u8 = 0;
const STACK_KIND_RANGE: u8 = 1;
const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff;

export let pairedSampleCount: i32 = 0;
export let pairedSumX: f64 = 0;
export let pairedSumY: f64 = 0;
export let pairedSumXX: f64 = 0;
export let pairedSumYY: f64 = 0;
export let pairedSumXY: f64 = 0;

function normalizeNearZero(value: f64): f64 {
  return Math.abs(value) < 1e-12 ? 0.0 : value;
}

export function collectPairedNumericStats(
  ySlot: i32,
  xSlot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  const yRows = inputRowsFromSlot(ySlot, kindStack, rangeIndexStack, rangeRowCounts);
  const yCols = inputColsFromSlot(ySlot, kindStack, rangeIndexStack, rangeColCounts);
  const xRows = inputRowsFromSlot(xSlot, kindStack, rangeIndexStack, rangeRowCounts);
  const xCols = inputColsFromSlot(xSlot, kindStack, rangeIndexStack, rangeColCounts);
  if (yRows < 1 || yCols < 1 || xRows < 1 || xCols < 1) {
    return ErrorCode.Value;
  }

  const yCount = yRows * yCols;
  const xCount = xRows * xCols;
  if (yCount != xCount || yCount <= 0) {
    return ErrorCode.Value;
  }

  pairedSampleCount = yCount;
  pairedSumX = 0;
  pairedSumY = 0;
  pairedSumXX = 0;
  pairedSumYY = 0;
  pairedSumXY = 0;

  for (let offset = 0; offset < yCount; offset += 1) {
    const yRow = yCols == 0 ? 0 : <i32>Math.floor(<f64>offset / <f64>yCols);
    const yCol = yCols == 0 ? 0 : offset - yRow * yCols;
    const xRow = xCols == 0 ? 0 : <i32>Math.floor(<f64>offset / <f64>xCols);
    const xCol = xCols == 0 ? 0 : offset - xRow * xCols;

    const yTag = inputCellTag(
      ySlot,
      yRow,
      yCol,
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
    if (yTag == ValueTag.Error) {
      return <i32>(
        inputCellScalarValue(
          ySlot,
          yRow,
          yCol,
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
        )
      );
    }

    const xTag = inputCellTag(
      xSlot,
      xRow,
      xCol,
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
    if (xTag == ValueTag.Error) {
      return <i32>(
        inputCellScalarValue(
          xSlot,
          xRow,
          xCol,
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
        )
      );
    }

    const yNumeric = inputCellNumeric(
      ySlot,
      yRow,
      yCol,
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
    const xNumeric = inputCellNumeric(
      xSlot,
      xRow,
      xCol,
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
    if (isNaN(yNumeric) || isNaN(xNumeric)) {
      return ErrorCode.Value;
    }

    pairedSumX += xNumeric;
    pairedSumY += yNumeric;
    pairedSumXX += xNumeric * xNumeric;
    pairedSumYY += yNumeric * yNumeric;
    pairedSumXY += xNumeric * yNumeric;
  }

  return 0;
}

export function pairedCenteredSumSquaresX(): f64 {
  return normalizeNearZero(pairedSumXX - (pairedSumX * pairedSumX) / <f64>pairedSampleCount);
}

export function pairedCenteredSumSquaresY(): f64 {
  return normalizeNearZero(pairedSumYY - (pairedSumY * pairedSumY) / <f64>pairedSampleCount);
}

export function pairedCenteredCrossProducts(): f64 {
  return normalizeNearZero(pairedSumXY - (pairedSumX * pairedSumY) / <f64>pairedSampleCount);
}

export function chiSquareTestPValue(
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
): f64 {
  const actualRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
  const actualCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
  const expectedRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
  const expectedCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
  if (
    actualRows < 1 ||
    actualCols < 1 ||
    expectedRows < 1 ||
    expectedCols < 1 ||
    actualRows != expectedRows ||
    actualCols != expectedCols ||
    (actualRows == 1 && actualCols == 1)
  ) {
    return -ErrorCode.NA;
  }

  let statistic = 0.0;
  for (let row = 0; row < actualRows; row += 1) {
    for (let col = 0; col < actualCols; col += 1) {
      const actualTag = inputCellTag(
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
      const actualRaw = inputCellScalarValue(
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
      if (actualTag == ValueTag.Error) {
        return -actualRaw;
      }
      const expectedTag = inputCellTag(
        base + 1,
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
      const expectedRaw = inputCellScalarValue(
        base + 1,
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
      if (expectedTag == ValueTag.Error) {
        return -expectedRaw;
      }

      const actualValue = toNumberOrNaN(actualTag, actualRaw);
      const expectedValue = toNumberOrNaN(expectedTag, expectedRaw);
      if (isNaN(actualValue) || isNaN(expectedValue) || actualValue < 0.0 || expectedValue < 0.0) {
        return -ErrorCode.Value;
      }
      if (expectedValue == 0.0) {
        return -ErrorCode.Div0;
      }

      const delta = actualValue - expectedValue;
      statistic += (delta * delta) / expectedValue;
    }
  }

  const degrees =
    actualRows > 1 && actualCols > 1
      ? (actualRows - 1) * (actualCols - 1)
      : actualRows > 1
        ? actualRows - 1
        : actualCols - 1;
  return degrees > 0 ? regularizedUpperGamma(<f64>degrees / 2.0, statistic / 2.0) : NaN;
}

export let sampleCollectionErrorCode = 0;

export function collectSampleNumbersFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): Array<f64> | null {
  sampleCollectionErrorCode = 0;
  const values = new Array<f64>();
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    if (tagStack[slot] == ValueTag.Error) {
      sampleCollectionErrorCode = <i32>valueStack[slot];
      return null;
    }
    if (tagStack[slot] != ValueTag.Number) {
      sampleCollectionErrorCode = ErrorCode.Value;
      return null;
    }
    values.push(valueStack[slot]);
    return values;
  }

  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
  if (rows < 1 || cols < 1) {
    sampleCollectionErrorCode = ErrorCode.Value;
    return null;
  }
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const tag = inputCellTag(
        slot,
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
      const raw = inputCellScalarValue(
        slot,
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
        sampleCollectionErrorCode = <i32>raw;
        return null;
      }
      if (tag == ValueTag.Number) {
        values.push(raw);
      }
    }
  }
  return values;
}

export function collectNumericSeriesFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  strict: bool,
): Array<f64> | null {
  const values = collectSampleNumbersFromSlot(
    slot,
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
  if (values === null || !strict) {
    return values;
  }
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
  if (rows < 1 || cols < 1 || values.length != rows * cols) {
    sampleCollectionErrorCode = ErrorCode.Value;
    return null;
  }
  return values;
}

export function collectDateSeriesFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): Array<i32> | null {
  const numericValues = collectNumericSeriesFromSlot(
    slot,
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
    true,
  );
  if (numericValues === null) {
    return null;
  }
  const dates = new Array<i32>();
  for (let index = 0; index < numericValues.length; index += 1) {
    const whole = <i32>unchecked(numericValues[index]);
    if (excelYearPartFromSerial(<u8>ValueTag.Number, <f64>whole) == i32.MIN_VALUE) {
      sampleCollectionErrorCode = ErrorCode.Value;
      return null;
    }
    dates.push(whole);
  }
  return dates;
}

export function collectNumericCellRangeSeriesFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  valueStack: Float64Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellErrors: Uint16Array,
  strict: bool,
): Array<f64> | null {
  sampleCollectionErrorCode = 0;
  if (kindStack[slot] == STACK_KIND_SCALAR) {
    if (tagStack[slot] == ValueTag.Error) {
      sampleCollectionErrorCode = <i32>valueStack[slot];
      return null;
    }
    if (tagStack[slot] != ValueTag.Number) {
      sampleCollectionErrorCode = ErrorCode.Value;
      return null;
    }
    const values = new Array<f64>();
    values.push(valueStack[slot]);
    return values;
  }
  if (kindStack[slot] != STACK_KIND_RANGE) {
    sampleCollectionErrorCode = ErrorCode.Value;
    return null;
  }
  const rangeIndex = rangeIndexStack[slot];
  if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
    sampleCollectionErrorCode = ErrorCode.Ref;
    return null;
  }
  const offset = rangeOffsets[rangeIndex];
  const length = <i32>rangeLengths[rangeIndex];
  const values = new Array<f64>();
  for (let index = 0; index < length; index += 1) {
    const memberIndex = unchecked(rangeMembers[offset + index]);
    const tag = unchecked(cellTags[memberIndex]);
    if (tag == ValueTag.Error) {
      sampleCollectionErrorCode = unchecked(cellErrors[memberIndex]);
      return null;
    }
    if (tag == ValueTag.Number) {
      values.push(unchecked(cellNumbers[memberIndex]));
    } else if (strict) {
      sampleCollectionErrorCode = ErrorCode.Value;
      return null;
    }
  }
  return values;
}

export function collectDateCellRangeSeriesFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  valueStack: Float64Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellErrors: Uint16Array,
): Array<i32> | null {
  const numericValues = collectNumericCellRangeSeriesFromSlot(
    slot,
    kindStack,
    tagStack,
    valueStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellErrors,
    true,
  );
  if (numericValues === null) {
    return null;
  }
  const dates = new Array<i32>();
  for (let index = 0; index < numericValues.length; index += 1) {
    const whole = <i32>unchecked(numericValues[index]);
    if (excelYearPartFromSerial(<u8>ValueTag.Number, <f64>whole) == i32.MIN_VALUE) {
      sampleCollectionErrorCode = ErrorCode.Value;
      return null;
    }
    dates.push(whole);
  }
  return dates;
}

export let orderStatisticErrorCode = 0;

export function collectNumericValuesFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): Array<f64> | null {
  orderStatisticErrorCode = 0;
  const values = new Array<f64>();
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
  if (rows < 1 || cols < 1) {
    orderStatisticErrorCode = ErrorCode.Value;
    return null;
  }
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const tag = inputCellTag(
        slot,
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
      const raw = inputCellScalarValue(
        slot,
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
        orderStatisticErrorCode = <i32>raw;
        return null;
      }
      const numeric = toNumberOrNaN(tag, raw);
      if (isNaN(numeric)) {
        orderStatisticErrorCode = ErrorCode.Value;
        return null;
      }
      values.push(numeric);
    }
  }
  return values;
}

export function collectNumericValuesFromArgs(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): Array<f64> | null {
  orderStatisticErrorCode = 0;
  const values = new Array<f64>();
  for (let index = 0; index < argc; index += 1) {
    const collected = collectNumericValuesFromSlot(
      base + index,
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
    if (collected === null) {
      return null;
    }
    for (let cursor = 0; cursor < collected.length; cursor += 1) {
      values.push(unchecked(collected[cursor]));
    }
  }
  return values;
}

export function fTestPValue(
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
): f64 {
  const first = collectSampleNumbersFromSlot(
    base,
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
  if (first === null) {
    return -(<f64>sampleCollectionErrorCode);
  }
  const second = collectSampleNumbersFromSlot(
    base + 1,
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
  if (second === null) {
    return -(<f64>sampleCollectionErrorCode);
  }
  if (first.length < 2 || second.length < 2) {
    return -(<f64>ErrorCode.Div0);
  }

  const firstVariance = sampleVarianceOf(first);
  const secondVariance = sampleVarianceOf(second);
  if (!(firstVariance > 0.0) || !(secondVariance > 0.0)) {
    return -(<f64>ErrorCode.Div0);
  }

  let numeratorVariance = firstVariance;
  let denominatorVariance = secondVariance;
  let numeratorDegrees = <f64>(first.length - 1);
  let denominatorDegrees = <f64>(second.length - 1);
  if (firstVariance < secondVariance) {
    numeratorVariance = secondVariance;
    denominatorVariance = firstVariance;
    numeratorDegrees = <f64>(second.length - 1);
    denominatorDegrees = <f64>(first.length - 1);
  }

  const upperTail =
    1.0 -
    fDistributionCdf(numeratorVariance / denominatorVariance, numeratorDegrees, denominatorDegrees);
  return isFinite(upperTail) ? min(1.0, upperTail * 2.0) : NaN;
}

export function zTestPValue(
  base: i32,
  argc: i32,
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
): f64 {
  const sample = collectSampleNumbersFromSlot(
    base,
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
  if (sample === null) {
    return -(<f64>sampleCollectionErrorCode);
  }
  if (sample.length == 0) {
    return -(<f64>ErrorCode.Value);
  }
  if (kindStack[base + 1] != STACK_KIND_SCALAR) {
    return -(<f64>ErrorCode.Value);
  }
  if (tagStack[base + 1] == ValueTag.Error) {
    return -valueStack[base + 1];
  }
  const x = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
  if (isNaN(x)) {
    return -(<f64>ErrorCode.Value);
  }

  let sigma = NaN;
  if (argc == 3) {
    if (kindStack[base + 2] != STACK_KIND_SCALAR) {
      return -(<f64>ErrorCode.Value);
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return -valueStack[base + 2];
    }
    sigma = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
  } else {
    sigma = Math.sqrt(sampleVarianceOf(sample));
  }

  if (!(sigma > 0.0)) {
    return -(<f64>ErrorCode.Div0);
  }
  const sampleMean = meanOf(sample);
  const zScore = (sampleMean - x) / (sigma / Math.sqrt(<f64>sample.length));
  return 1.0 - standardNormalCdf(zScore);
}

export function tTestPValue(
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
): f64 {
  const first = collectSampleNumbersFromSlot(
    base,
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
  if (first === null) {
    return -(<f64>sampleCollectionErrorCode);
  }
  const second = collectSampleNumbersFromSlot(
    base + 1,
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
  if (second === null) {
    return -(<f64>sampleCollectionErrorCode);
  }
  if (kindStack[base + 2] != STACK_KIND_SCALAR || kindStack[base + 3] != STACK_KIND_SCALAR) {
    return -(<f64>ErrorCode.Value);
  }
  if (tagStack[base + 2] == ValueTag.Error) {
    return -valueStack[base + 2];
  }
  if (tagStack[base + 3] == ValueTag.Error) {
    return -valueStack[base + 3];
  }
  const tailsRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
  const typeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
  const tails = <i32>tailsRaw;
  const testType = <i32>typeRaw;
  if (
    isNaN(tailsRaw) ||
    isNaN(typeRaw) ||
    tailsRaw != <f64>tails ||
    typeRaw != <f64>testType ||
    (tails != 1 && tails != 2) ||
    (testType != 1 && testType != 2 && testType != 3)
  ) {
    return -(<f64>ErrorCode.Value);
  }

  let statistic = NaN;
  let degreesFreedom = NaN;
  if (testType == 1) {
    if (first.length != second.length) {
      return -(<f64>ErrorCode.NA);
    }
    if (first.length < 2) {
      return -(<f64>ErrorCode.Div0);
    }
    const deltas = new Array<f64>();
    for (let index = 0; index < first.length; index += 1) {
      deltas.push(unchecked(first[index]) - unchecked(second[index]));
    }
    const variance = sampleVarianceOf(deltas);
    if (!(variance > 0.0)) {
      return -(<f64>ErrorCode.Div0);
    }
    statistic = meanOf(deltas) / Math.sqrt(variance / <f64>deltas.length);
    degreesFreedom = <f64>(deltas.length - 1);
  } else {
    if (first.length < 2 || second.length < 2) {
      return -(<f64>ErrorCode.Div0);
    }
    const firstMean = meanOf(first);
    const secondMean = meanOf(second);
    const firstVariance = sampleVarianceOf(first);
    const secondVariance = sampleVarianceOf(second);
    if (!(firstVariance > 0.0) || !(secondVariance > 0.0)) {
      return -(<f64>ErrorCode.Div0);
    }
    if (testType == 2) {
      const pooledVariance =
        (<f64>(first.length - 1) * firstVariance + <f64>(second.length - 1) * secondVariance) /
        <f64>(first.length + second.length - 2);
      if (!(pooledVariance > 0.0)) {
        return -(<f64>ErrorCode.Div0);
      }
      statistic =
        (firstMean - secondMean) /
        Math.sqrt(pooledVariance * (1.0 / <f64>first.length + 1.0 / <f64>second.length));
      degreesFreedom = <f64>(first.length + second.length - 2);
    } else {
      const firstTerm = firstVariance / <f64>first.length;
      const secondTerm = secondVariance / <f64>second.length;
      const denominator = Math.sqrt(firstTerm + secondTerm);
      const welchDenominator =
        (firstTerm * firstTerm) / <f64>(first.length - 1) +
        (secondTerm * secondTerm) / <f64>(second.length - 1);
      if (!(denominator > 0.0) || !(welchDenominator > 0.0)) {
        return -(<f64>ErrorCode.Div0);
      }
      statistic = (firstMean - secondMean) / denominator;
      degreesFreedom = ((firstTerm + secondTerm) * (firstTerm + secondTerm)) / welchDenominator;
    }
  }

  const upperTail = 1.0 - studentTCdf(Math.abs(statistic), degreesFreedom);
  const probability = tails == 1 ? upperTail : min(1.0, upperTail * 2.0);
  return isFinite(probability) ? probability : NaN;
}
