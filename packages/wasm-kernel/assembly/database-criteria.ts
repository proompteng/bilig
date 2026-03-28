import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import { matchesCriteriaValue } from "./criteria";
import { truncToInt } from "./numeric-core";
import { memberScalarValue, rangeMemberAt, toNumberOrNaN } from "./operands";
import { STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from "./result-io";
import { scalarText, trimAsciiWhitespace } from "./text-codec";

function stackScalarTagOrSingleCellRange(
  slot: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
): i32 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return tagStack[slot];
  }
  if (kind != STACK_KIND_RANGE) {
    return -1;
  }
  const rangeIndex = rangeIndexStack[slot];
  if (<i32>rangeLengths[rangeIndex] != 1) {
    return -1;
  }
  return cellTags[rangeMembers[rangeOffsets[rangeIndex]]];
}

function stackScalarValueOrSingleCellRange(
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
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return valueStack[slot];
  }
  if (kind != STACK_KIND_RANGE) {
    return 0;
  }
  const rangeIndex = rangeIndexStack[slot];
  if (<i32>rangeLengths[rangeIndex] != 1) {
    return 0;
  }
  const memberIndex = rangeMembers[rangeOffsets[rangeIndex]];
  return memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors);
}

function writeDatabaseError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.Error,
    error,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

export function databaseBuiltinResult(
  builtinId: u16,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const databaseSlot = base;
  const fieldSlot = base + 1;
  const criteriaSlot = base + 2;
  if (kindStack[databaseSlot] != STACK_KIND_RANGE || kindStack[criteriaSlot] != STACK_KIND_RANGE) {
    return writeDatabaseError(
      base,
      ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const databaseRangeIndex = rangeIndexStack[databaseSlot];
  const criteriaRangeIndex = rangeIndexStack[criteriaSlot];
  const databaseRows = <i32>rangeRowCounts[databaseRangeIndex];
  const databaseCols = <i32>rangeColCounts[databaseRangeIndex];
  const criteriaRows = <i32>rangeRowCounts[criteriaRangeIndex];
  const criteriaCols = <i32>rangeColCounts[criteriaRangeIndex];
  if (databaseRows < 1 || databaseCols < 1 || criteriaRows < 2 || criteriaCols < 1) {
    return writeDatabaseError(
      base,
      ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const fieldTag = stackScalarTagOrSingleCellRange(
    fieldSlot,
    kindStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
  );
  if (fieldTag < 0) {
    return writeDatabaseError(
      base,
      ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  const fieldValue = stackScalarValueOrSingleCellRange(
    fieldSlot,
    kindStack,
    tagStack,
    valueStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  );
  if (fieldTag == ValueTag.Error) {
    return writeDatabaseError(
      base,
      <i32>fieldValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const allowOmittedField = builtinId == BuiltinId.Dcount || builtinId == BuiltinId.Dcounta;
  let omitField = false;
  let fieldIndex = -1;
  if (fieldTag == ValueTag.Empty) {
    if (allowOmittedField) {
      omitField = true;
    } else {
      return writeDatabaseError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
  } else if (fieldTag == ValueTag.Number) {
    const position = truncToInt(<u8>fieldTag, fieldValue);
    if (position == i32.MIN_VALUE || position < 1 || position > databaseCols) {
      return writeDatabaseError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    fieldIndex = position - 1;
  } else {
    const fieldText = scalarText(
      <u8>fieldTag,
      fieldValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (fieldText == null) {
      return writeDatabaseError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const normalizedFieldText = trimAsciiWhitespace(fieldText).toUpperCase();
    if (normalizedFieldText.length == 0 && allowOmittedField) {
      omitField = true;
    } else {
      for (let col = 0; col < databaseCols; col += 1) {
        const headerMemberIndex = rangeMemberAt(
          databaseRangeIndex,
          0,
          col,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
        );
        if (headerMemberIndex == 0xffffffff) {
          continue;
        }
        const headerTag = cellTags[headerMemberIndex];
        const headerValue = memberScalarValue(
          headerMemberIndex,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (headerTag == ValueTag.Error) {
          return writeDatabaseError(
            base,
            <i32>headerValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const headerText = scalarText(
          headerTag,
          headerValue,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (
          headerText != null &&
          trimAsciiWhitespace(headerText).toUpperCase() == normalizedFieldText
        ) {
          fieldIndex = col;
          break;
        }
      }
      if (fieldIndex < 0 && !omitField) {
        return writeDatabaseError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
  }

  let recordCount = 0;
  let numericCount = 0;
  let sum = 0.0;
  let sumSquares = 0.0;
  let product = 1.0;
  let minimum = Infinity;
  let maximum = -Infinity;
  let hasNumeric = false;
  let dgetFound = false;
  let dgetTag = <u8>ValueTag.Empty;
  let dgetValue = 0.0;

  for (let databaseRow = 1; databaseRow < databaseRows; databaseRow += 1) {
    let matchesAnyCriteriaRow = false;
    for (let criteriaRow = 1; criteriaRow < criteriaRows; criteriaRow += 1) {
      let blocked = false;
      let matchesAll = true;
      for (let criteriaCol = 0; criteriaCol < criteriaCols; criteriaCol += 1) {
        const criteriaMemberIndex = rangeMemberAt(
          criteriaRangeIndex,
          criteriaRow,
          criteriaCol,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
        );
        if (criteriaMemberIndex == 0xffffffff) {
          continue;
        }
        const criteriaTag = cellTags[criteriaMemberIndex];
        const criteriaValue = memberScalarValue(
          criteriaMemberIndex,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (criteriaTag == ValueTag.Empty) {
          continue;
        }
        if (criteriaTag == ValueTag.Error) {
          return writeDatabaseError(
            base,
            <i32>criteriaValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }

        const headerMemberIndex = rangeMemberAt(
          criteriaRangeIndex,
          0,
          criteriaCol,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
        );
        if (headerMemberIndex == 0xffffffff) {
          blocked = true;
          continue;
        }
        const headerTag = cellTags[headerMemberIndex];
        const headerValue = memberScalarValue(
          headerMemberIndex,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (headerTag == ValueTag.Error) {
          return writeDatabaseError(
            base,
            <i32>headerValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const headerText = scalarText(
          headerTag,
          headerValue,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (headerText == null) {
          blocked = true;
          continue;
        }
        const normalizedHeaderText = trimAsciiWhitespace(headerText).toUpperCase();
        if (normalizedHeaderText.length == 0) {
          blocked = true;
          continue;
        }

        let criteriaDatabaseCol = -1;
        for (let databaseCol = 0; databaseCol < databaseCols; databaseCol += 1) {
          const databaseHeaderMemberIndex = rangeMemberAt(
            databaseRangeIndex,
            0,
            databaseCol,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
          );
          if (databaseHeaderMemberIndex == 0xffffffff) {
            continue;
          }
          const databaseHeaderTag = cellTags[databaseHeaderMemberIndex];
          const databaseHeaderValue = memberScalarValue(
            databaseHeaderMemberIndex,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
          if (databaseHeaderTag == ValueTag.Error) {
            return writeDatabaseError(
              base,
              <i32>databaseHeaderValue,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          const databaseHeaderText = scalarText(
            databaseHeaderTag,
            databaseHeaderValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (
            databaseHeaderText != null &&
            trimAsciiWhitespace(databaseHeaderText).toUpperCase() == normalizedHeaderText
          ) {
            criteriaDatabaseCol = databaseCol;
            break;
          }
        }
        if (criteriaDatabaseCol < 0) {
          blocked = true;
          continue;
        }

        const databaseMemberIndex = rangeMemberAt(
          databaseRangeIndex,
          databaseRow,
          criteriaDatabaseCol,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
        );
        if (databaseMemberIndex == 0xffffffff) {
          matchesAll = false;
          break;
        }
        if (
          !matchesCriteriaValue(
            cellTags[databaseMemberIndex],
            memberScalarValue(
              databaseMemberIndex,
              cellTags,
              cellNumbers,
              cellStringIds,
              cellErrors,
            ),
            criteriaTag,
            criteriaValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false;
          break;
        }
      }
      if (!blocked && matchesAll) {
        matchesAnyCriteriaRow = true;
        break;
      }
    }

    if (!matchesAnyCriteriaRow) {
      continue;
    }

    recordCount += 1;
    if (omitField) {
      continue;
    }

    const fieldMemberIndex = rangeMemberAt(
      databaseRangeIndex,
      databaseRow,
      fieldIndex,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (fieldMemberIndex == 0xffffffff) {
      continue;
    }
    const memberTag = cellTags[fieldMemberIndex];
    const memberValue = memberScalarValue(
      fieldMemberIndex,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );

    if (builtinId == BuiltinId.Dget) {
      if (dgetFound) {
        return writeDatabaseError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      dgetFound = true;
      dgetTag = memberTag;
      dgetValue = memberValue;
      continue;
    }

    if (builtinId == BuiltinId.Dcount) {
      if (!isNaN(toNumberOrNaN(memberTag, memberValue))) {
        numericCount += 1;
      }
      continue;
    }
    if (builtinId == BuiltinId.Dcounta) {
      if (memberTag != ValueTag.Empty) {
        numericCount += 1;
      }
      continue;
    }

    if (memberTag != ValueTag.Number && memberTag != ValueTag.Boolean) {
      continue;
    }
    const numeric = memberValue;
    numericCount += 1;
    sum += numeric;
    sumSquares += numeric * numeric;
    product *= numeric;
    minimum = min(minimum, numeric);
    maximum = max(maximum, numeric);
    hasNumeric = true;
  }

  if (builtinId == BuiltinId.Dcount || builtinId == BuiltinId.Dcounta) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      omitField ? recordCount : numericCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Dget) {
    if (!dgetFound) {
      return writeDatabaseError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      dgetTag,
      dgetValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Daverage) {
    return numericCount == 0
      ? writeDatabaseError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          sum / <f64>numericCount,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }
  if (builtinId == BuiltinId.Dmax || builtinId == BuiltinId.Dmin || builtinId == BuiltinId.Dsum) {
    const value =
      builtinId == BuiltinId.Dmax
        ? hasNumeric
          ? maximum
          : 0.0
        : builtinId == BuiltinId.Dmin
          ? hasNumeric
            ? minimum
            : 0.0
          : sum;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Dproduct) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      numericCount == 0 ? 0.0 : product,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    numericCount == 0 ||
    ((builtinId == BuiltinId.Dstdev || builtinId == BuiltinId.Dvar) && numericCount < 2)
  ) {
    return writeDatabaseError(
      base,
      ErrorCode.Div0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const mean = sum / numericCount;
  const numericCountF64 = <f64>numericCount;
  let variance = sumSquares - numericCount * mean * mean;
  variance /=
    builtinId == BuiltinId.Dstdev || builtinId == BuiltinId.Dvar
      ? numericCountF64 - 1.0
      : numericCountF64;
  if (variance < 0 && variance > -1e-12) {
    variance = 0;
  }
  const result =
    builtinId == BuiltinId.Dstdev || builtinId == BuiltinId.Dstdevp ? sqrt(variance) : variance;
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.Number,
    result,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}
