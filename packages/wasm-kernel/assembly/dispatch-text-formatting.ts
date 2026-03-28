import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import { scalarArgsOnly, scalarErrorAt } from "./builtin-args";
import { poolString, scalarText } from "./text-codec";
import {
  arrayToTextCell,
  coerceScalarNumberLikeText,
  firstUnicodeCodePoint,
  stringFromUnicodeCodePoint,
  stripControlCharacters,
  toJapaneseFullWidth,
  toJapaneseHalfWidth,
} from "./text-special";
import {
  bahtTextFromNumber,
  containsDateTimeTokens,
  errorLabel,
  formatDateTimePatternText,
  formatNumericPatternText,
  formatTextSectionText,
  jsonQuoteText,
  numberValueParseText,
  splitFormatSectionsText,
  stripFormatDecorationsText,
} from "./text-format";
import {
  indexOfTextWithMode,
  lastIndexOfTextWithMode,
  splitTextByDelimiterWithMode,
} from "./text-ops";
import { valueNumber } from "./comparison";
import { coerceInteger } from "./numeric-core";
import {
  STACK_KIND_RANGE,
  STACK_KIND_SCALAR,
  copySlotResult,
  writeArrayResult,
  writeResult,
  writeStringResult,
} from "./result-io";
import {
  allocateOutputString,
  allocateSpillArrayResult,
  encodeOutputStringId,
  writeOutputStringData,
  writeSpillArrayValue,
} from "./vm";
import { toNumberExact } from "./operands";

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  return -1;
}

export function tryApplyTextFormattingBuiltin(
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Textsplit && argc >= 2 && argc <= 6) {
    if (!scalarArgsOnly(base, argc, kindStack)) {
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

    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const columnDelimiter = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const rowDelimiter =
      argc >= 3
        ? scalarText(
            tagStack[base + 2],
            valueStack[base + 2],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : null;
    const ignoreEmpty = argc >= 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 0;
    const matchModeNumeric =
      argc >= 5
        ? valueNumber(
            tagStack[base + 4],
            valueStack[base + 4],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0;
    if (
      text == null ||
      columnDelimiter == null ||
      (argc >= 3 && rowDelimiter == null) ||
      ignoreEmpty < 0 ||
      !isFinite(matchModeNumeric) ||
      matchModeNumeric != <f64>(<i32>matchModeNumeric)
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

    const matchMode = <i32>matchModeNumeric;
    if (!(matchMode == 0 || matchMode == 1)) {
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
    if (columnDelimiter.length == 0 && argc < 3) {
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

    const padTag = argc >= 6 ? tagStack[base + 5] : <u8>ValueTag.Error;
    const padValue = argc >= 6 ? valueStack[base + 5] : <f64>ErrorCode.NA;

    let rowDelimiterText = "";
    if (argc >= 3) {
      rowDelimiterText = rowDelimiter == null ? "" : rowDelimiter;
    }
    const rowSlices = new Array<string>();
    if (argc < 3 || rowDelimiterText.length == 0) {
      rowSlices.push(text);
    } else {
      const splitRows = splitTextByDelimiterWithMode(text, rowDelimiterText, matchMode);
      for (let rowIndex = 0; rowIndex < splitRows.length; rowIndex += 1) {
        rowSlices.push(splitRows[rowIndex]);
      }
    }

    const matrix = new Array<Array<string>>();
    let maxCols = 1;
    for (let rowIndex = 0; rowIndex < rowSlices.length; rowIndex += 1) {
      const rowSlice = rowSlices[rowIndex];
      const rawParts = new Array<string>();
      if (columnDelimiter.length == 0) {
        rawParts.push(rowSlice);
      } else {
        const splitParts = splitTextByDelimiterWithMode(rowSlice, columnDelimiter, matchMode);
        for (let partIndex = 0; partIndex < splitParts.length; partIndex += 1) {
          rawParts.push(splitParts[partIndex]);
        }
      }
      const filtered = new Array<string>();
      for (let partIndex = 0; partIndex < rawParts.length; partIndex += 1) {
        const part = rawParts[partIndex];
        if (ignoreEmpty == 1 && part.length == 0) {
          continue;
        }
        filtered.push(part);
      }
      matrix.push(filtered);
      if (filtered.length > maxCols) {
        maxCols = filtered.length;
      }
    }

    const rows = max<i32>(1, matrix.length);
    const cols = max<i32>(1, maxCols);
    const arrayIndex = allocateSpillArrayResult(rows, cols);
    let outputOffset = 0;
    for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
      const row = rowIndex < matrix.length ? matrix[rowIndex] : new Array<string>();
      for (let colIndex = 0; colIndex < cols; colIndex += 1) {
        if (colIndex < row.length) {
          const part = row[colIndex];
          const outputStringId = allocateOutputString(part.length);
          for (let index = 0; index < part.length; index += 1) {
            writeOutputStringData(outputStringId, index, <u16>part.charCodeAt(index));
          }
          writeSpillArrayValue(
            arrayIndex,
            outputOffset,
            <u8>ValueTag.String,
            encodeOutputStringId(outputStringId),
          );
        } else {
          writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue);
        }
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

  if (
    (builtinId == BuiltinId.Textbefore || builtinId == BuiltinId.Textafter) &&
    argc >= 2 &&
    argc <= 6
  ) {
    if (!scalarArgsOnly(base, argc, kindStack)) {
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
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const delimiter = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || delimiter == null || delimiter.length == 0) {
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

    const instanceNumeric =
      argc >= 3
        ? valueNumber(
            tagStack[base + 2],
            valueStack[base + 2],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 1;
    const matchModeNumeric =
      argc >= 4
        ? valueNumber(
            tagStack[base + 3],
            valueStack[base + 3],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0;
    const matchEndNumeric =
      argc >= 5
        ? valueNumber(
            tagStack[base + 4],
            valueStack[base + 4],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0;
    if (
      !isFinite(instanceNumeric) ||
      !isFinite(matchModeNumeric) ||
      !isFinite(matchEndNumeric) ||
      instanceNumeric != <f64>(<i32>instanceNumeric) ||
      matchModeNumeric != <f64>(<i32>matchModeNumeric) ||
      matchEndNumeric != <f64>(<i32>matchEndNumeric)
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

    const instance = <i32>instanceNumeric;
    const matchMode = <i32>matchModeNumeric;
    if (instance == 0 || !(matchMode == 0 || matchMode == 1)) {
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
    if (!(<i32>matchEndNumeric == 0 || <i32>matchEndNumeric == 1)) {
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

    let found = -1;
    if (instance > 0) {
      let searchFrom = 0;
      for (let count = 0; count < instance; count += 1) {
        found = indexOfTextWithMode(text, delimiter, searchFrom, matchMode);
        if (found < 0) {
          return argc == 6
            ? copySlotResult(base, base + 5, rangeIndexStack, valueStack, tagStack, kindStack)
            : writeResult(
                base,
                STACK_KIND_SCALAR,
                <u8>ValueTag.Error,
                ErrorCode.NA,
                rangeIndexStack,
                valueStack,
                tagStack,
                kindStack,
              );
        }
        searchFrom = found + delimiter.length;
      }
    } else {
      let searchFrom = text.length;
      for (let count = 0; count < -instance; count += 1) {
        found = lastIndexOfTextWithMode(text, delimiter, searchFrom, matchMode);
        if (found < 0) {
          return argc == 6
            ? copySlotResult(base, base + 5, rangeIndexStack, valueStack, tagStack, kindStack)
            : writeResult(
                base,
                STACK_KIND_SCALAR,
                <u8>ValueTag.Error,
                ErrorCode.NA,
                rangeIndexStack,
                valueStack,
                tagStack,
                kindStack,
              );
        }
        searchFrom = found - 1;
      }
    }

    return writeStringResult(
      base,
      builtinId == BuiltinId.Textbefore
        ? text.slice(0, found)
        : text.slice(found + delimiter.length),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Value && argc == 1) {
    const numeric = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (isNaN(numeric)) {
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      numeric,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Char && argc == 1) {
    const numeric = coerceScalarNumberLikeText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (!isFinite(numeric)) {
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
    const integerCode = <i32>numeric;
    if (integerCode < 1 || integerCode > 255) {
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
    return writeStringResult(
      base,
      String.fromCharCode(integerCode),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Code || builtinId == BuiltinId.Unicode) && argc == 1) {
    if (tagStack[base] == ValueTag.Error) {
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
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || text.length == 0) {
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>firstUnicodeCodePoint(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Unichar && argc == 1) {
    const numeric = coerceScalarNumberLikeText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (!isFinite(numeric)) {
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
    const integerCode = <i32>numeric;
    if (integerCode < 0 || integerCode > 0x10ffff) {
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
    return writeStringResult(
      base,
      stringFromUnicodeCodePoint(integerCode),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Clean && argc == 1) {
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
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
    return writeStringResult(
      base,
      stripControlCharacters(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Asc || builtinId == BuiltinId.Jis || builtinId == BuiltinId.Dbcs) &&
    argc == 1
  ) {
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
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
    const converted =
      builtinId == BuiltinId.Asc ? toJapaneseHalfWidth(text) : toJapaneseFullWidth(text);
    return writeStringResult(base, converted, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Bahttext && argc == 1) {
    const numeric = coerceScalarNumberLikeText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const text = bahtTextFromNumber(numeric);
    if (text == null) {
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
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Numbervalue && argc >= 1 && argc <= 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const decimalSeparator =
      argc >= 2
        ? scalarText(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : ".";
    const groupSeparator =
      argc >= 3
        ? scalarText(
            tagStack[base + 2],
            valueStack[base + 2],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : ",";
    if (text == null || decimalSeparator == null || groupSeparator == null) {
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
    const parsed = numberValueParseText(text, decimalSeparator, groupSeparator);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(parsed) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(parsed) ? ErrorCode.Value : parsed,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Valuetotext && (argc == 1 || argc == 2)) {
    if (
      argc == 2 &&
      kindStack[base + 1] == STACK_KIND_SCALAR &&
      tagStack[base + 1] == ValueTag.Error
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (kindStack[base] == STACK_KIND_SCALAR && tagStack[base] == ValueTag.Error) {
      return writeStringResult(
        base,
        errorLabel(<i32>valueStack[base]),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const format = argc == 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (format == i32.MIN_VALUE || (format != 0 && format != 1)) {
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

    let textResult: string | null = null;
    const tag = tagStack[base];
    if (tag == ValueTag.Empty) {
      textResult = "";
    } else if (tag == ValueTag.Number) {
      textResult = valueStack[base].toString();
    } else if (tag == ValueTag.Boolean) {
      textResult = valueStack[base] != 0 ? "TRUE" : "FALSE";
    } else if (tag == ValueTag.String) {
      const rawText = scalarText(
        tag,
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      textResult = rawText == null ? null : format == 1 ? jsonQuoteText(rawText) : rawText;
    } else if (tag == ValueTag.Error) {
      textResult = errorLabel(<i32>valueStack[base]);
    }
    if (textResult == null) {
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
    return writeStringResult(base, textResult, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Text && argc == 2) {
    const formatText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (formatText == null) {
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
    const sections = splitFormatSectionsText(formatText);
    let section = sections.length > 0 ? sections[0] : "";
    let numeric = NaN;
    let autoNegative = false;

    if (tagStack[base] == ValueTag.String) {
      if (sections.length >= 4) {
        section = sections[3];
      }
      const cleaned = stripFormatDecorationsText(section);
      let hasPlaceholder = false;
      for (let index = 0; index < cleaned.length; index += 1) {
        if (
          cleaned.charCodeAt(index) == 48 ||
          cleaned.charCodeAt(index) == 35 ||
          cleaned.charCodeAt(index) == 63
        ) {
          hasPlaceholder = true;
          break;
        }
      }
      if (cleaned.indexOf("@") >= 0 || (!hasPlaceholder && !containsDateTimeTokens(cleaned))) {
        const textValue = scalarText(
          tagStack[base],
          valueStack[base],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (textValue == null) {
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
        return writeStringResult(
          base,
          formatTextSectionText(textValue, section),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
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

    numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (!isFinite(numeric)) {
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
    if (numeric < 0.0) {
      if (sections.length >= 2) {
        section = sections[1];
      }
      numeric = -numeric;
      autoNegative = sections.length < 2;
    } else if (numeric == 0.0 && sections.length >= 3) {
      section = sections[2];
    }

    const cleaned = stripFormatDecorationsText(section);
    if (containsDateTimeTokens(cleaned)) {
      const formatted = formatDateTimePatternText(<u8>ValueTag.Number, numeric, section);
      if (formatted == null) {
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
      return writeStringResult(base, formatted, rangeIndexStack, valueStack, tagStack, kindStack);
    }

    return writeStringResult(
      base,
      formatNumericPatternText(numeric, section, autoNegative),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Phonetic && argc == 1) {
    if (kindStack[base] == STACK_KIND_RANGE) {
      const rangeIndex = rangeIndexStack[base];
      const rangeLength = <i32>rangeLengths[rangeIndex];
      if (rangeLength <= 0) {
        return writeStringResult(base, "", rangeIndexStack, valueStack, tagStack, kindStack);
      }
      const memberIndex = rangeMembers[rangeOffsets[rangeIndex]];
      const memberTag = cellTags[memberIndex];
      if (memberTag == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          cellErrors[memberIndex],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const memberText =
        memberTag == ValueTag.String
          ? poolString(<i32>cellStringIds[memberIndex], stringOffsets, stringLengths, stringData)
          : arrayToTextCell(
              memberTag,
              cellNumbers[memberIndex],
              false,
              stringOffsets,
              stringLengths,
              stringData,
              outputStringOffsets,
              outputStringLengths,
              outputStringData,
            );
      if (memberText == null) {
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
      return writeStringResult(base, memberText, rangeIndexStack, valueStack, tagStack, kindStack);
    }

    const text = arrayToTextCell(
      tagStack[base],
      valueStack[base],
      false,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
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
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  return -1;
}
