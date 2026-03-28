import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  getTrackedArrayCols as getDynamicArrayCols,
  getTrackedArrayRows as getDynamicArrayRows,
  registerTrackedArrayShape as registerTrackedArrayShapeImpl,
} from "./dynamic-arrays";
import {
  leftBytesText,
  midBytesText,
  poolString,
  replaceBytesText,
  rightBytesText,
  scalarText,
  textLength,
  trimAsciiWhitespace,
  utf8ByteLength,
  bytePositionToCharPositionUtf8,
  charPositionToBytePositionUtf8,
} from "./text-codec";
import {
  coerceScalarNumberLikeText,
  firstUnicodeCodePoint,
  parseNumericText,
  parseTimeValueText,
  stringFromUnicodeCodePoint,
  stripControlCharacters,
  toJapaneseFullWidth,
  toJapaneseHalfWidth,
} from "./text-special";
import {
  formatSignedRadixText,
  isValidBaseText,
  parseBaseText,
  parseSignedRadixText,
  toBaseText,
} from "./radix";
import {
  EXCEL_SECONDS_PER_DAY,
  accruedIssueYearfracValue,
  addMonthsExcelSerial,
  couponDateFromMaturityValue,
  couponDaysByBasisValue,
  couponPeriodDaysValue,
  couponPeriodsRemainingValue,
  couponPriceFromMetricsValue,
  dbDepreciation,
  ddbDepreciation,
  excelDateSerial,
  excelDatedifValue,
  excelDayPartFromSerial,
  excelDays360Value,
  excelIsoWeeknumValue,
  excelMonthPartFromSerial,
  excelSecondOfDay,
  excelSerialWhole,
  excelTimeSerial,
  excelWeekdayFromSerial,
  excelWeeknumFromSerial,
  excelYearPartFromSerial,
  excelYearfracValue,
  macaulayDurationValue,
  maturityIssueYearfracValue,
  oddFirstPriceValue,
  oddFirstYieldValue,
  oddLastPriceValue,
  oddLastYieldValue,
  securityAnnualizedYearfracValue,
  solveCouponYieldValue,
  treasuryBillDaysValue,
  vdbDepreciation,
} from "./date-finance";
import {
  cumulativePeriodicPaymentCalc,
  futureValueCalc,
  hasPositiveAndNegativeSeries,
  interestPaymentCalc,
  mirrCalc,
  periodicCashflowNetPresentValueCalc,
  periodicPaymentCalc,
  presentValueCalc,
  principalPaymentCalc,
  solvePeriodicCashflowRateCalc,
  solveRateCalc,
  solveXirrCalc,
  totalPeriodsCalc,
  xnpvCalc,
} from "./cashflows";
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  besselIValue,
  besselJValue,
  besselKValue,
  besselYValue,
  binomialProbability,
  chiSquareCdf,
  chiSquareDensity,
  erfApprox,
  fDistributionCdf,
  fDistributionDensity,
  gammaDistributionCdf,
  gammaDistributionDensity,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  negativeBinomialProbability,
  poissonProbability,
  regularizedUpperGamma,
  standardNormalCdf,
  standardNormalPdf,
  studentTCdf,
  studentTDensity,
} from "./distributions";
import {
  interpolateSortedPercentRank,
  interpolateSortedPercentile,
  kurtosisOf,
  meanOf,
  modeSingleOf,
  populationVarianceOf,
  sampleVarianceOf,
  skewPopulationOf,
  skewSampleOf,
  sortNumericValues,
  truncateToSignificance,
} from "./statistics-core";
import {
  allocateOutputString,
  allocateSpillArrayResult,
  encodeOutputStringId,
  nextVolatileRandomValue,
  readSpillArrayTag,
  readSpillArrayLength,
  readSpillArrayNumber,
  readVolatileNowSerial,
  writeOutputStringData,
  writeSpillArrayNumber,
  writeSpillArrayValue,
} from "./vm";

export const STACK_KIND_SCALAR: u8 = 0;
export const STACK_KIND_RANGE: u8 = 1;
export const STACK_KIND_ARRAY: u8 = 2;
const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff;

export function registerTrackedArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols);
}

function toNumberOrNaN(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function toNumberOrZero(tag: u8, value: f64): f64 {
  const numeric = toNumberOrNaN(tag, value);
  return isNaN(numeric) ? 0 : numeric;
}

function toNumberExact(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function truncAbs(value: f64): f64 {
  if (!isFinite(value)) {
    return NaN;
  }
  return Math.abs(value < 0.0 ? Math.ceil(value) : Math.floor(value));
}

function factorialCalc(value: f64): f64 {
  if (!isFinite(value) || value < 0.0) {
    return NaN;
  }
  const truncated = <i32>Math.floor(value);
  let result = 1.0;
  for (let index = 2; index <= truncated; index += 1) {
    result *= <f64>index;
  }
  return result;
}

function doubleFactorialCalc(value: f64): f64 {
  if (!isFinite(value) || value < 0.0) {
    return NaN;
  }
  const truncated = <i32>Math.floor(value);
  let result = 1.0;
  for (let index = truncated; index >= 2; index -= 2) {
    result *= <f64>index;
  }
  return result;
}

function gcdPairCalc(left: f64, right: f64): f64 {
  let a = truncAbs(left);
  let b = truncAbs(right);
  while (b != 0.0) {
    const next = a % b;
    a = b;
    b = next;
  }
  return a;
}

function lcmPairCalc(left: f64, right: f64): f64 {
  const a = truncAbs(left);
  const b = truncAbs(right);
  if (a == 0.0 || b == 0.0) {
    return 0.0;
  }
  return Math.abs((a * b) / gcdPairCalc(a, b));
}

function evenCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0;
  const rounded = Math.ceil(Math.abs(value) / 2.0) * 2.0;
  return sign * rounded;
}

function oddCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0;
  const rounded = Math.ceil(Math.abs(value));
  const odd = rounded % 2.0 == 0.0 ? rounded + 1.0 : rounded;
  return sign * odd;
}

const MAX_SAFE_INTEGER_F64: f64 = 9007199254740991.0;

function volatileNowResult(): f64 {
  return readVolatileNowSerial();
}

function arrayToTextCell(
  tag: u8,
  value: f64,
  strict: bool,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Number) {
    return value == Math.trunc(value) ? (<i64>value).toString() : value.toString();
  }
  const text = scalarText(
    tag,
    value,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (text == null) {
    return null;
  }
  if (!strict || tag != ValueTag.String) {
    return text;
  }
  return '"' + substituteText(text, '"', '""') + '"';
}

function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric;
}

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  return -1;
}

function roundToDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits);
    return Math.round(value * factor) / factor;
  }
  const factor = Math.pow(10.0, <f64>-digits);
  return Math.round(value / factor) * factor;
}

function roundTowardZeroDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits);
    return Math.trunc(value * factor) / factor;
  }
  const factor = Math.pow(10.0, <f64>-digits);
  return Math.trunc(value / factor) * factor;
}

function columnLabelText(column: i32): string | null {
  if (column < 1) {
    return null;
  }
  let current = column;
  let label = "";
  while (current > 0) {
    const offset = (current - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    current = (current - 1) / 26;
  }
  return label;
}

function escapeSheetNameText(value: string): string {
  let output = "";
  for (let index = 0; index < value.length; index++) {
    const char = value.charAt(index);
    output += char;
    if (char == "'") {
      output += "'";
    }
  }
  return output;
}

function formatThousandsText(integerPart: string): string {
  let output = "";
  let groupCount = 0;
  for (let index = integerPart.length - 1; index >= 0; index--) {
    output = integerPart.charAt(index) + output;
    groupCount += 1;
    if (index > 0 && groupCount % 3 == 0) {
      output = "," + output;
    }
  }
  return output;
}

function zeroPadIntegerText(value: i64, width: i32): string {
  let output = value.toString();
  while (output.length < width) {
    output = "0" + output;
  }
  return output;
}

function formatFixedText(value: f64, decimals: i32, includeThousands: bool): string | null {
  if (!isFinite(value)) {
    return null;
  }
  const rounded = roundToDigits(value, decimals);
  const sign = rounded < 0.0 ? "-" : "";
  const unsigned = Math.abs(rounded);
  const fixedDecimals = decimals >= 0 ? decimals : 0;
  const scale = Math.pow(10.0, <f64>fixedDecimals);
  let integerPartValue = <i64>Math.floor(unsigned);
  let scaledFraction =
    fixedDecimals > 0 ? <i64>Math.round((unsigned - <f64>integerPartValue) * scale) : 0;
  if (fixedDecimals > 0 && scaledFraction >= <i64>scale) {
    integerPartValue += 1;
    scaledFraction -= <i64>scale;
  }
  const integerPart = integerPartValue.toString();
  const fractionPart = fixedDecimals > 0 ? zeroPadIntegerText(scaledFraction, fixedDecimals) : "";
  const normalizedInteger = includeThousands ? formatThousandsText(integerPart) : integerPart;
  return fractionPart.length > 0
    ? `${sign}${normalizedInteger}.${fractionPart}`
    : `${sign}${normalizedInteger}`;
}

function splitFormatSectionsText(format: string): Array<string> {
  const sections = new Array<string>();
  let current = "";
  let inQuotes = false;
  let bracketDepth = 0;
  let escaped = false;
  for (let index = 0; index < format.length; index += 1) {
    const char = format.charAt(index);
    if (escaped) {
      current += char;
      escaped = false;
      continue;
    }
    if (char == "\\") {
      current += char;
      escaped = true;
      continue;
    }
    if (char == '"') {
      current += char;
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && char == "[") {
      bracketDepth += 1;
      current += char;
      continue;
    }
    if (!inQuotes && char == "]" && bracketDepth > 0) {
      bracketDepth -= 1;
      current += char;
      continue;
    }
    if (!inQuotes && bracketDepth == 0 && char == ";") {
      sections.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  sections.push(current);
  return sections;
}

function stripFormatDecorationsText(section: string): string {
  let output = "";
  let inQuotes = false;
  for (let index = 0; index < section.length; index += 1) {
    const char = section.charAt(index);
    if (inQuotes) {
      if (char == '"') {
        inQuotes = false;
      } else {
        output += char;
      }
      continue;
    }
    if (char == '"') {
      inQuotes = true;
      continue;
    }
    if (char == "\\") {
      if (index + 1 < section.length) {
        output += section.charAt(index + 1);
        index += 1;
      }
      continue;
    }
    if (char == "_") {
      output += " ";
      index += 1;
      continue;
    }
    if (char == "*") {
      index += 1;
      continue;
    }
    if (char == "[") {
      const end = section.indexOf("]", index + 1);
      if (end >= 0) {
        index = end;
      }
      continue;
    }
    output += char;
  }
  return output;
}

function isFormatPlaceholderCode(code: i32): bool {
  return code == 48 || code == 35 || code == 63;
}

function countOccurrences(text: string, needle: string): i32 {
  let count = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (text.charAt(index) == needle) {
      count += 1;
    }
  }
  return count;
}

function countZeroPlaceholders(text: string): i32 {
  let count = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) == 48) {
      count += 1;
    }
  }
  return count;
}

function countDigitPlaceholders(text: string): i32 {
  let count = 0;
  for (let index = 0; index < text.length; index += 1) {
    if (isFormatPlaceholderCode(text.charCodeAt(index))) {
      count += 1;
    }
  }
  return count;
}

function zeroPadPositiveText(value: i32, width: i32): string {
  let output = value.toString();
  while (output.length < width) {
    output = "0" + output;
  }
  return output;
}

function trimOptionalFractionText(fraction: string, minDigits: i32): string {
  let trimmed = fraction;
  while (trimmed.length > minDigits && trimmed.endsWith("0")) {
    trimmed = trimmed.slice(0, trimmed.length - 1);
  }
  return trimmed;
}

function monthNameText(month: i32, longName: bool): string {
  switch (month) {
    case 1:
      return longName ? "January" : "Jan";
    case 2:
      return longName ? "February" : "Feb";
    case 3:
      return longName ? "March" : "Mar";
    case 4:
      return longName ? "April" : "Apr";
    case 5:
      return "May";
    case 6:
      return longName ? "June" : "Jun";
    case 7:
      return longName ? "July" : "Jul";
    case 8:
      return longName ? "August" : "Aug";
    case 9:
      return longName ? "September" : "Sep";
    case 10:
      return longName ? "October" : "Oct";
    case 11:
      return longName ? "November" : "Nov";
    case 12:
      return longName ? "December" : "Dec";
    default:
      return "";
  }
}

function weekdayNameText(index: i32, longName: bool): string {
  switch (index) {
    case 0:
      return longName ? "Sunday" : "Sun";
    case 1:
      return longName ? "Monday" : "Mon";
    case 2:
      return longName ? "Tuesday" : "Tue";
    case 3:
      return longName ? "Wednesday" : "Wed";
    case 4:
      return longName ? "Thursday" : "Thu";
    case 5:
      return longName ? "Friday" : "Fri";
    case 6:
      return longName ? "Saturday" : "Sat";
    default:
      return "";
  }
}

function excelWeekdayIndexFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE || whole < 0) {
    return i32.MIN_VALUE;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return ((adjustedWhole % 7) + 7) % 7;
}

function containsDateTimeTokens(text: string): bool {
  const upper = text.toUpperCase();
  if (upper.indexOf("AM/PM") >= 0 || upper.indexOf("A/P") >= 0) {
    return true;
  }
  for (let index = 0; index < upper.length; index += 1) {
    const code = upper.charCodeAt(index);
    if (code == 89 || code == 68 || code == 72 || code == 83 || code == 77) {
      return true;
    }
  }
  return false;
}

function formatTextSectionText(value: string, section: string): string {
  const cleaned = stripFormatDecorationsText(section);
  let output = "";
  for (let index = 0; index < cleaned.length; index += 1) {
    if (cleaned.charAt(index) == "@") {
      output += value;
    } else {
      output += cleaned.charAt(index);
    }
  }
  return output;
}

const DATE_TOKEN_NONE: i32 = 0;
const DATE_TOKEN_YEAR: i32 = 1;
const DATE_TOKEN_MONTH: i32 = 2;
const DATE_TOKEN_MINUTE: i32 = 3;
const DATE_TOKEN_DAY: i32 = 4;
const DATE_TOKEN_HOUR: i32 = 5;
const DATE_TOKEN_SECOND: i32 = 6;

function nextDateTokenKind(text: string, start: i32): i32 {
  for (let index = start; index < text.length; index += 1) {
    const remainder = text.substring(index).toUpperCase();
    if (remainder.startsWith("AM/PM") || remainder.startsWith("A/P")) {
      return DATE_TOKEN_NONE;
    }
    const lower = text.charAt(index).toLowerCase();
    if (lower == "y") return DATE_TOKEN_YEAR;
    if (lower == "m") return DATE_TOKEN_MONTH;
    if (lower == "d") return DATE_TOKEN_DAY;
    if (lower == "h") return DATE_TOKEN_HOUR;
    if (lower == "s") return DATE_TOKEN_SECOND;
  }
  return DATE_TOKEN_NONE;
}

function formatAmPmText(token: string, hour: i32): string {
  const isPm = hour >= 12;
  const upper = token.toUpperCase();
  if (upper == "A/P") {
    const letter = isPm ? "P" : "A";
    return token == token.toLowerCase() ? letter.toLowerCase() : letter;
  }
  return token == token.toLowerCase() ? (isPm ? "pm" : "am") : isPm ? "PM" : "AM";
}

function formatDateTimePatternText(tag: u8, value: f64, section: string): string | null {
  const cleaned = stripFormatDecorationsText(section);
  const year = excelYearPartFromSerial(tag, value);
  const month = excelMonthPartFromSerial(tag, value);
  const day = excelDayPartFromSerial(tag, value);
  const secondOfDay = excelSecondOfDay(tag, value);
  const weekdayIndex = excelWeekdayIndexFromSerial(tag, value);
  if (
    year == i32.MIN_VALUE ||
    month == i32.MIN_VALUE ||
    day == i32.MIN_VALUE ||
    secondOfDay == i32.MIN_VALUE ||
    weekdayIndex == i32.MIN_VALUE
  ) {
    return null;
  }
  const hour24 = secondOfDay / 3600;
  const minute = (secondOfDay % 3600) / 60;
  const second = secondOfDay % 60;
  const hasAmPm =
    cleaned.toUpperCase().indexOf("AM/PM") >= 0 || cleaned.toUpperCase().indexOf("A/P") >= 0;
  let output = "";
  let previousKind = DATE_TOKEN_NONE;
  for (let index = 0; index < cleaned.length; ) {
    const remainderUpper = cleaned.substring(index).toUpperCase();
    if (remainderUpper.startsWith("AM/PM")) {
      output += formatAmPmText(cleaned.substring(index, index + 5), hour24);
      index += 5;
      previousKind = DATE_TOKEN_NONE;
      continue;
    }
    if (remainderUpper.startsWith("A/P")) {
      output += formatAmPmText(cleaned.substring(index, index + 3), hour24);
      index += 3;
      previousKind = DATE_TOKEN_NONE;
      continue;
    }
    const lower = cleaned.charAt(index).toLowerCase();
    if (lower == "y" || lower == "m" || lower == "d" || lower == "h" || lower == "s") {
      let end = index + 1;
      while (end < cleaned.length && cleaned.charAt(end).toLowerCase() == lower) {
        end += 1;
      }
      const token = cleaned.substring(index, end);
      if (lower == "y") {
        output += token.length == 2 ? zeroPadPositiveText(year % 100, 2) : year.toString();
        previousKind = DATE_TOKEN_YEAR;
      } else if (lower == "d") {
        output +=
          token.length == 1
            ? day.toString()
            : token.length == 2
              ? zeroPadPositiveText(day, 2)
              : token.length == 3
                ? weekdayNameText(weekdayIndex, false)
                : weekdayNameText(weekdayIndex, true);
        previousKind = DATE_TOKEN_DAY;
      } else if (lower == "h") {
        const hourValue = hasAmPm ? ((hour24 + 11) % 12) + 1 : hour24;
        output += token.length >= 2 ? zeroPadPositiveText(hourValue, 2) : hourValue.toString();
        previousKind = DATE_TOKEN_HOUR;
      } else if (lower == "s") {
        output += token.length >= 2 ? zeroPadPositiveText(second, 2) : second.toString();
        previousKind = DATE_TOKEN_SECOND;
      } else {
        const nextKind = nextDateTokenKind(cleaned, end);
        const minuteToken =
          previousKind == DATE_TOKEN_HOUR ||
          previousKind == DATE_TOKEN_MINUTE ||
          nextKind == DATE_TOKEN_SECOND;
        if (minuteToken) {
          output += token.length >= 2 ? zeroPadPositiveText(minute, 2) : minute.toString();
          previousKind = DATE_TOKEN_MINUTE;
        } else {
          output +=
            token.length == 1
              ? month.toString()
              : token.length == 2
                ? zeroPadPositiveText(month, 2)
                : token.length == 3
                  ? monthNameText(month, false)
                  : monthNameText(month, true);
          previousKind = DATE_TOKEN_MONTH;
        }
      }
      index = end;
      continue;
    }
    output += cleaned.charAt(index);
    index += 1;
  }
  return output;
}

function parseSignedIntegerText(text: string): i32 {
  if (text.length == 0) {
    return i32.MIN_VALUE;
  }
  let index = 0;
  let sign = 1;
  if (text.charAt(0) == "+") {
    index = 1;
  } else if (text.charAt(0) == "-") {
    sign = -1;
    index = 1;
  }
  if (index >= text.length) {
    return i32.MIN_VALUE;
  }
  let value = 0;
  for (; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if (code < 48 || code > 57) {
      return i32.MIN_VALUE;
    }
    value = value * 10 + (code - 48);
  }
  return sign * value;
}

function formatScientificPatternText(value: f64, core: string): string {
  let exponentIndex = -1;
  for (let index = 0; index + 1 < core.length; index += 1) {
    const upper = core.charAt(index).toUpperCase();
    const sign = core.charAt(index + 1);
    if (upper == "E" && (sign == "+" || sign == "-")) {
      exponentIndex = index;
      break;
    }
  }
  const mantissaPattern = core.slice(0, exponentIndex);
  const exponentPattern = core.slice(exponentIndex + 2);
  const dotIndex = mantissaPattern.indexOf(".");
  const fractionPattern = dotIndex >= 0 ? mantissaPattern.slice(dotIndex + 1) : "";
  const maxFractionDigits = countDigitPlaceholders(fractionPattern);
  const minFractionDigits = countZeroPlaceholders(fractionPattern);
  let exponentValue = value == 0.0 ? 0 : <i32>Math.floor(Math.log(value) / Math.log(10.0));
  let mantissaValue = value == 0.0 ? 0.0 : value / Math.pow(10.0, <f64>exponentValue);
  mantissaValue = roundToDigits(mantissaValue, maxFractionDigits);
  if (mantissaValue >= 10.0) {
    mantissaValue /= 10.0;
    exponentValue += 1;
  }
  let mantissa = formatFixedText(mantissaValue, maxFractionDigits, false);
  if (mantissa == null) {
    mantissa = "0";
  }
  const mantissaDot = mantissa.indexOf(".");
  if (mantissaDot >= 0) {
    const integerPart = mantissa.slice(0, mantissaDot);
    const trimmedFraction = trimOptionalFractionText(
      mantissa.slice(mantissaDot + 1),
      minFractionDigits,
    );
    mantissa = trimmedFraction.length > 0 ? `${integerPart}.${trimmedFraction}` : integerPart;
  }
  return `${mantissa}E${exponentValue < 0 ? "-" : "+"}${zeroPadPositiveText(<i32>Math.abs(<f64>exponentValue), exponentPattern.length)}`;
}

function formatNumericPatternText(value: f64, section: string, autoNegative: bool): string {
  const cleaned = stripFormatDecorationsText(section);
  let firstPlaceholder = -1;
  let lastPlaceholder = -1;
  for (let index = 0; index < cleaned.length; index += 1) {
    if (isFormatPlaceholderCode(cleaned.charCodeAt(index))) {
      if (firstPlaceholder == -1) {
        firstPlaceholder = index;
      }
      lastPlaceholder = index;
    }
  }
  if (firstPlaceholder == -1) {
    return autoNegative && !cleaned.startsWith("-") ? `-${cleaned}` : cleaned;
  }
  const prefix = cleaned.slice(0, firstPlaceholder);
  const core = cleaned.slice(firstPlaceholder, lastPlaceholder + 1);
  const suffix = cleaned.slice(lastPlaceholder + 1);
  const scaledValue = Math.abs(value) * Math.pow(100.0, <f64>countOccurrences(cleaned, "%"));
  let numericText = "";
  let exponentIndex = -1;
  for (let index = 0; index + 1 < core.length; index += 1) {
    const upper = core.charAt(index).toUpperCase();
    const sign = core.charAt(index + 1);
    if (upper == "E" && (sign == "+" || sign == "-")) {
      exponentIndex = index;
      break;
    }
  }
  if (exponentIndex >= 0) {
    numericText = formatScientificPatternText(scaledValue, core);
  } else {
    const decimalIndex = core.indexOf(".");
    const integerPatternRaw = decimalIndex >= 0 ? core.slice(0, decimalIndex) : core;
    const fractionPattern = decimalIndex >= 0 ? core.slice(decimalIndex + 1) : "";
    const integerPattern = integerPatternRaw.replace(",", "");
    const maxFractionDigits = countDigitPlaceholders(fractionPattern);
    const minFractionDigits = countZeroPlaceholders(fractionPattern);
    const minIntegerDigits = countZeroPlaceholders(integerPattern);
    const rounded = roundToDigits(scaledValue, maxFractionDigits);
    let fixed = formatFixedText(rounded, maxFractionDigits, false);
    if (fixed == null) {
      fixed = "0";
    }
    const fixedDot = fixed.indexOf(".");
    let integerPart: string = fixedDot >= 0 ? fixed.slice(0, fixedDot) : fixed;
    let fractionPart: string = fixedDot >= 0 ? fixed.slice(fixedDot + 1) : "";
    while (integerPart.length < minIntegerDigits) {
      integerPart = "0" + integerPart;
    }
    if (integerPatternRaw.indexOf(",") >= 0) {
      integerPart = formatThousandsText(integerPart);
    }
    fractionPart = trimOptionalFractionText(fractionPart, minFractionDigits);
    numericText = fractionPart.length > 0 ? `${integerPart}.${fractionPart}` : integerPart;
  }
  const combined = `${prefix}${numericText}${suffix}`;
  return autoNegative && !combined.startsWith("-") ? `-${combined}` : combined;
}

function digitCount(value: i32): i32 {
  if (value <= 0) {
    return 1;
  }
  let current = value;
  let count = 0;
  while (current > 0) {
    count += 1;
    current /= 10;
  }
  return count;
}

function isValidDollarFractionNative(fraction: i32): bool {
  if (fraction <= 0) {
    return false;
  }
  if (fraction == 1) {
    return true;
  }
  let current = fraction;
  while ((current & 1) == 0) {
    current >>= 1;
  }
  return current == 1;
}

function parsePositiveDigits(value: string): i32 {
  if (value.length == 0) {
    return 0;
  }
  let output = 0;
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code < 48 || code > 57) {
      return i32.MIN_VALUE;
    }
    output = output * 10 + (code - 48);
  }
  return output;
}

function dollarFractionalNumerator(value: f64): i32 {
  const absoluteText = Math.abs(value).toString();
  const dot = absoluteText.indexOf(".");
  if (dot < 0) {
    return 0;
  }
  return parsePositiveDigits(absoluteText.substring(dot + 1));
}

function coerceBitwiseUnsigned(tag: u8, value: f64): i64 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i64.MIN_VALUE;
  }
  const truncated = Math.trunc(numeric);
  if (Math.abs(truncated) > MAX_SAFE_INTEGER_F64) {
    return i64.MIN_VALUE;
  }
  return <i64>(<u32>(<i64>truncated));
}

function coerceNonNegativeShift(tag: u8, value: f64): i64 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i64.MIN_VALUE;
  }
  const truncated = Math.trunc(numeric);
  if (truncated < 0.0 || truncated > MAX_SAFE_INTEGER_F64) {
    return i64.MIN_VALUE;
  }
  return <i64>truncated;
}

function signedRadixInputText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Error) {
    return null;
  }
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
      return null;
    }
    return trimAsciiWhitespace(text).toUpperCase();
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return null;
  }
  return (<i64>numeric).toString().toUpperCase();
}

function writeStringResult(
  base: i32,
  text: string,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  const outputStringId = allocateOutputString(text.length);
  for (let index = 0; index < text.length; index++) {
    writeOutputStringData(outputStringId, index, <u16>text.charCodeAt(index));
  }
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.String,
    encodeOutputStringId(outputStringId),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

function coerceLength(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

function coercePositiveStart(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

const CONVERT_GROUP_INVALID = 0;
const CONVERT_GROUP_MASS = 1;
const CONVERT_GROUP_DISTANCE = 2;
const CONVERT_GROUP_TIME = 3;
const CONVERT_GROUP_PRESSURE = 4;
const CONVERT_GROUP_FORCE = 5;
const CONVERT_GROUP_ENERGY = 6;
const CONVERT_GROUP_POWER = 7;
const CONVERT_GROUP_MAGNETISM = 8;
const CONVERT_GROUP_TEMPERATURE = 9;
const CONVERT_GROUP_VOLUME = 10;
const CONVERT_GROUP_AREA = 11;
const CONVERT_GROUP_INFORMATION = 12;
const CONVERT_GROUP_SPEED = 13;

const CONVERT_TEMP_NONE = 0;
const CONVERT_TEMP_C = 1;
const CONVERT_TEMP_F = 2;
const CONVERT_TEMP_K = 3;
const CONVERT_TEMP_RANK = 4;
const CONVERT_TEMP_REAU = 5;

let resolvedConvertGroup = CONVERT_GROUP_INVALID;
let resolvedConvertFactor = NaN;
let resolvedConvertTemperature = CONVERT_TEMP_NONE;

function setResolvedConvertScaleUnit(group: i32, factor: f64): bool {
  resolvedConvertGroup = group;
  resolvedConvertFactor = factor;
  resolvedConvertTemperature = CONVERT_TEMP_NONE;
  return true;
}

function setResolvedConvertTemperatureUnit(code: i32): bool {
  resolvedConvertGroup = CONVERT_GROUP_TEMPERATURE;
  resolvedConvertFactor = NaN;
  resolvedConvertTemperature = code;
  return true;
}

function resolveConvertExactUnit(unitText: string): bool {
  if (unitText == "g") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1.0);
  if (unitText == "sg") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 14593.902937206363);
  if (unitText == "lbm") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 453.59237);
  if (unitText == "u") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1.660538782e-24);
  if (unitText == "ozm") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 28.349523125);
  if (unitText == "grain") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 0.06479891);
  if (unitText == "cwt" || unitText == "shweight") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 45359.237);
  }
  if (unitText == "uk_cwt" || unitText == "lcwt" || unitText == "hweight") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 50802.34544);
  }
  if (unitText == "stone") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 6350.29318);
  if (unitText == "ton") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 907184.74);
  if (unitText == "uk_ton" || unitText == "LTON" || unitText == "brton") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, 1016046.9088);
  }

  if (unitText == "m") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1.0);
  if (unitText == "mi") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1609.344);
  if (unitText == "Nmi") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1852.0);
  if (unitText == "in") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254);
  if (unitText == "ft") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.3048);
  if (unitText == "yd") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.9144);
  if (unitText == "ang") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1e-10);
  if (unitText == "ell") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1.143);
  if (unitText == "ly")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 9.4607304725808e15);
  if (unitText == "parsec" || unitText == "pc") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 3.085677581491367e16);
  }
  if (unitText == "Picapt" || unitText == "Pica") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254 / 72.0);
  }
  if (unitText == "pica") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 0.0254 / 6.0);
  if (unitText == "survey_mi") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, 1609.3472186944373);
  }

  if (unitText == "yr") return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 31557600.0);
  if (unitText == "day" || unitText == "d")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 86400.0);
  if (unitText == "hr") return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 3600.0);
  if (unitText == "mn" || unitText == "min")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 60.0);
  if (unitText == "sec" || unitText == "s")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, 1.0);

  if (unitText == "Pa" || unitText == "p")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 1.0);
  if (unitText == "atm" || unitText == "at")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 101325.0);
  if (unitText == "mmHg" || unitText == "Torr") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 101325.0 / 760.0);
  }
  if (unitText == "psi")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, 6894.757293168361);

  if (unitText == "N") return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 1.0);
  if (unitText == "dyn" || unitText == "dy")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 1e-5);
  if (unitText == "lbf") return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, 4.4482216152605);

  if (unitText == "J") return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.0);
  if (unitText == "e") return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1e-7);
  if (unitText == "c") return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 4.184);
  if (unitText == "cal") return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 4.1868);
  if (unitText == "eV" || unitText == "ev") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.602176487e-19);
  }
  if (unitText == "HPh" || unitText == "hh") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 2684519.537696173);
  }
  if (unitText == "Wh" || unitText == "wh")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 3600.0);
  if (unitText == "flb")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.3558179483314004);
  if (unitText == "BTU" || unitText == "btu") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1055.05585262);
  }

  if (unitText == "HP" || unitText == "h")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 745.6998715822701);
  if (unitText == "PS") return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 735.49875);
  if (unitText == "W" || unitText == "w")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, 1.0);

  if (unitText == "T") return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, 1.0);
  if (unitText == "ga") return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, 1e-4);

  if (unitText == "C" || unitText == "cel")
    return setResolvedConvertTemperatureUnit(CONVERT_TEMP_C);
  if (unitText == "F" || unitText == "fah")
    return setResolvedConvertTemperatureUnit(CONVERT_TEMP_F);
  if (unitText == "K" || unitText == "kel")
    return setResolvedConvertTemperatureUnit(CONVERT_TEMP_K);
  if (unitText == "Rank") return setResolvedConvertTemperatureUnit(CONVERT_TEMP_RANK);
  if (unitText == "Reau") return setResolvedConvertTemperatureUnit(CONVERT_TEMP_REAU);

  if (unitText == "tsp") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.92892159375e-6);
  if (unitText == "tspm") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 5e-6);
  if (unitText == "tbs")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.478676478125e-5);
  if (unitText == "oz") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 2.95735295625e-5);
  if (unitText == "cup") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.0002365882365);
  if (unitText == "pt" || unitText == "us_pt")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.000473176473);
  if (unitText == "uk_pt") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.00056826125);
  if (unitText == "qt") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.000946352946);
  if (unitText == "uk_qt") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.0011365225);
  if (unitText == "gal") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.003785411784);
  if (unitText == "uk_gal") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.00454609);
  if (unitText == "l" || unitText == "L" || unitText == "lt") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.001);
  }
  if (unitText == "ang3" || unitText == "ang^3")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1e-30);
  if (unitText == "barrel")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.158987294928);
  if (unitText == "bushel")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.03523907016688);
  if (unitText == "ft3" || unitText == "ft^3")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.028316846592);
  if (unitText == "in3" || unitText == "in^3")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.6387064e-5);
  if (unitText == "ly3" || unitText == "ly^3") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 8.46786664623715e47);
  }
  if (unitText == "m3" || unitText == "m^3")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.0);
  if (unitText == "mi3" || unitText == "mi^3") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.16818182544058e9);
  }
  if (unitText == "yd3" || unitText == "yd^3")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.764554857984);
  if (unitText == "Nmi3" || unitText == "Nmi^3") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 6.350012608e9);
  }
  if (
    unitText == "Picapt3" ||
    unitText == "Picapt^3" ||
    unitText == "Pica3" ||
    unitText == "Pica^3"
  ) {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 4.390243770290368e-11);
  }
  if (unitText == "GRT" || unitText == "regton")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 2.8316846592);
  if (unitText == "MTON") return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 1.13267386368);

  if (unitText == "uk_acre") return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 4046.8564224);
  if (unitText == "us_acre")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 4046.872609874252);
  if (unitText == "ang2" || unitText == "ang^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1e-20);
  if (unitText == "ar") return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 100.0);
  if (unitText == "ft2" || unitText == "ft^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.09290304);
  if (unitText == "ha") return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 10000.0);
  if (unitText == "in2" || unitText == "in^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.00064516);
  if (unitText == "ly2" || unitText == "ly^2") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 8.95054210748189e31);
  }
  if (unitText == "m2" || unitText == "m^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1.0);
  if (unitText == "Morgen") return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 2500.0);
  if (unitText == "mi2" || unitText == "mi^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 2589988.110336);
  if (unitText == "Nmi2" || unitText == "Nmi^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 3429904.0);
  if (
    unitText == "Picapt2" ||
    unitText == "Pica2" ||
    unitText == "Pica^2" ||
    unitText == "Picapt^2"
  ) {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 1.244707012345679e-7);
  }
  if (unitText == "yd2" || unitText == "yd^2")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, 0.83612736);

  if (unitText == "bit") return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 1.0);
  if (unitText == "byte") return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0);

  if (unitText == "admkn")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 0.5147733333333333);
  if (unitText == "kn") return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1852.0 / 3600.0);
  if (unitText == "m/h" || unitText == "m/hr")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1.0 / 3600.0);
  if (unitText == "m/s" || unitText == "m/sec")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1.0);
  if (unitText == "mph") return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, 1609.344 / 3600.0);

  return false;
}

function resolveConvertMetricSuffix(suffix: string, prefixFactor: f64): bool {
  if (suffix == "g") return setResolvedConvertScaleUnit(CONVERT_GROUP_MASS, prefixFactor);
  if (suffix == "m") return setResolvedConvertScaleUnit(CONVERT_GROUP_DISTANCE, prefixFactor);
  if (suffix == "sec" || suffix == "s")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_TIME, prefixFactor);
  if (suffix == "Pa" || suffix == "p")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_PRESSURE, prefixFactor);
  if (suffix == "N") return setResolvedConvertScaleUnit(CONVERT_GROUP_FORCE, prefixFactor);
  if (suffix == "J") return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, prefixFactor);
  if (suffix == "eV" || suffix == "ev") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_ENERGY, 1.602176487e-19 * prefixFactor);
  }
  if (suffix == "W" || suffix == "w")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_POWER, prefixFactor);
  if (suffix == "T") return setResolvedConvertScaleUnit(CONVERT_GROUP_MAGNETISM, prefixFactor);
  if (suffix == "l" || suffix == "L")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, 0.001 * prefixFactor);
  if (suffix == "m3" || suffix == "m^3") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_VOLUME, Math.pow(prefixFactor, 3.0));
  }
  if (suffix == "m2" || suffix == "m^2") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_AREA, Math.pow(prefixFactor, 2.0));
  }
  if (suffix == "bit") return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, prefixFactor);
  if (suffix == "byte")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0 * prefixFactor);
  if (suffix == "m/h" || suffix == "m/hr") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, prefixFactor / 3600.0);
  }
  if (suffix == "m/s" || suffix == "m/sec") {
    return setResolvedConvertScaleUnit(CONVERT_GROUP_SPEED, prefixFactor);
  }
  return false;
}

function resolveConvertBinarySuffix(suffix: string, prefixFactor: f64): bool {
  if (suffix == "bit") return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, prefixFactor);
  if (suffix == "byte")
    return setResolvedConvertScaleUnit(CONVERT_GROUP_INFORMATION, 8.0 * prefixFactor);
  return false;
}

function resolveConvertMetricPrefixedUnit(unitText: string): bool {
  if (unitText.length <= 1) return false;
  if (unitText.startsWith("da")) return resolveConvertMetricSuffix(unitText.slice(2), 1e1);
  if (unitText.startsWith("Y")) return resolveConvertMetricSuffix(unitText.slice(1), 1e24);
  if (unitText.startsWith("Z")) return resolveConvertMetricSuffix(unitText.slice(1), 1e21);
  if (unitText.startsWith("E")) return resolveConvertMetricSuffix(unitText.slice(1), 1e18);
  if (unitText.startsWith("P")) return resolveConvertMetricSuffix(unitText.slice(1), 1e15);
  if (unitText.startsWith("T")) return resolveConvertMetricSuffix(unitText.slice(1), 1e12);
  if (unitText.startsWith("G")) return resolveConvertMetricSuffix(unitText.slice(1), 1e9);
  if (unitText.startsWith("M")) return resolveConvertMetricSuffix(unitText.slice(1), 1e6);
  if (unitText.startsWith("k")) return resolveConvertMetricSuffix(unitText.slice(1), 1e3);
  if (unitText.startsWith("h")) return resolveConvertMetricSuffix(unitText.slice(1), 1e2);
  if (unitText.startsWith("e")) return resolveConvertMetricSuffix(unitText.slice(1), 1e1);
  if (unitText.startsWith("d")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-1);
  if (unitText.startsWith("c")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-2);
  if (unitText.startsWith("m")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-3);
  if (unitText.startsWith("u")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-6);
  if (unitText.startsWith("n")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-9);
  if (unitText.startsWith("p")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-12);
  if (unitText.startsWith("f")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-15);
  if (unitText.startsWith("a")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-18);
  if (unitText.startsWith("z")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-21);
  if (unitText.startsWith("y")) return resolveConvertMetricSuffix(unitText.slice(1), 1e-24);
  return false;
}

function resolveConvertBinaryPrefixedUnit(unitText: string): bool {
  if (unitText.length <= 2) return false;
  if (unitText.startsWith("Yi"))
    return resolveConvertBinarySuffix(unitText.slice(2), 1208925819614629174706176.0);
  if (unitText.startsWith("Zi"))
    return resolveConvertBinarySuffix(unitText.slice(2), 1180591620717411303424.0);
  if (unitText.startsWith("Ei"))
    return resolveConvertBinarySuffix(unitText.slice(2), 1152921504606846976.0);
  if (unitText.startsWith("Pi"))
    return resolveConvertBinarySuffix(unitText.slice(2), 1125899906842624.0);
  if (unitText.startsWith("Ti"))
    return resolveConvertBinarySuffix(unitText.slice(2), 1099511627776.0);
  if (unitText.startsWith("Gi")) return resolveConvertBinarySuffix(unitText.slice(2), 1073741824.0);
  if (unitText.startsWith("Mi")) return resolveConvertBinarySuffix(unitText.slice(2), 1048576.0);
  if (unitText.startsWith("ki")) return resolveConvertBinarySuffix(unitText.slice(2), 1024.0);
  return false;
}

function resolveConvertUnit(unitText: string): bool {
  return (
    resolveConvertExactUnit(unitText) ||
    resolveConvertBinaryPrefixedUnit(unitText) ||
    resolveConvertMetricPrefixedUnit(unitText)
  );
}

function convertTemperatureToKelvin(unitCode: i32, value: f64): f64 {
  if (unitCode == CONVERT_TEMP_C) return value + 273.15;
  if (unitCode == CONVERT_TEMP_F) return (value + 459.67) * (5.0 / 9.0);
  if (unitCode == CONVERT_TEMP_K) return value;
  if (unitCode == CONVERT_TEMP_RANK) return value * (5.0 / 9.0);
  if (unitCode == CONVERT_TEMP_REAU) return value * 1.25 + 273.15;
  return NaN;
}

function convertKelvinToTemperature(unitCode: i32, value: f64): f64 {
  if (unitCode == CONVERT_TEMP_C) return value - 273.15;
  if (unitCode == CONVERT_TEMP_F) return value * (9.0 / 5.0) - 459.67;
  if (unitCode == CONVERT_TEMP_K) return value;
  if (unitCode == CONVERT_TEMP_RANK) return value * (9.0 / 5.0);
  if (unitCode == CONVERT_TEMP_REAU) return (value - 273.15) * 0.8;
  return NaN;
}

function roundToPlacesNative(value: f64, places: i32): f64 {
  const scale = Math.pow(10.0, <f64>places);
  return Math.round(value * scale) / scale;
}

function roundToSignificantDigitsNative(value: f64, digits: i32): f64 {
  if (value == 0.0 || !isFinite(value)) {
    return value;
  }
  const exponent = <i32>Math.floor(Math.log(Math.abs(value)) / Math.log(10.0));
  const scale = Math.pow(10.0, <f64>(digits - exponent - 1));
  return Math.round(value * scale) / scale;
}

function euroRateNative(code: string): f64 {
  if (code == "BEF" || code == "LUF") return 40.3399;
  if (code == "DEM") return 1.95583;
  if (code == "ESP") return 166.386;
  if (code == "FRF") return 6.55957;
  if (code == "IEP") return 0.787564;
  if (code == "ITL") return 1936.27;
  if (code == "NLG") return 2.20371;
  if (code == "ATS") return 13.7603;
  if (code == "PTE") return 200.482;
  if (code == "FIM") return 5.94573;
  if (code == "GRD") return 340.75;
  if (code == "SIT") return 239.64;
  if (code == "EUR") return 1.0;
  return NaN;
}

function euroCalculationPrecisionNative(code: string): i32 {
  if (
    code == "BEF" ||
    code == "LUF" ||
    code == "ESP" ||
    code == "ITL" ||
    code == "PTE" ||
    code == "GRD"
  ) {
    return 0;
  }
  if (
    code == "DEM" ||
    code == "FRF" ||
    code == "IEP" ||
    code == "NLG" ||
    code == "ATS" ||
    code == "FIM" ||
    code == "SIT" ||
    code == "EUR"
  ) {
    return 2;
  }
  return i32.MIN_VALUE;
}

function excelTrim(input: string): string {
  let start = 0;
  let end = input.length;
  while (start < end && input.charCodeAt(start) == 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) == 32) {
    end -= 1;
  }
  let result = "";
  let previousSpace = false;
  for (let index = start; index < end; index++) {
    const char = input.charCodeAt(index);
    if (char == 32) {
      if (!previousSpace) {
        result += " ";
      }
      previousSpace = true;
      continue;
    }
    previousSpace = false;
    result += String.fromCharCode(char);
  }
  return result;
}

function hasSearchSyntax(pattern: string): bool {
  for (let index = 0; index < pattern.length; index++) {
    const char = pattern.charCodeAt(index);
    if (char == 126 || char == 42 || char == 63) {
      return true;
    }
  }
  return false;
}

function wildcardMatchAt(
  pattern: string,
  haystack: string,
  patternIndex: i32,
  haystackIndex: i32,
): bool {
  let p = patternIndex;
  let h = haystackIndex;
  while (p < pattern.length) {
    const char = pattern.charCodeAt(p);
    if (char == 126) {
      const nextIndex = p + 1;
      const nextChar = nextIndex < pattern.length ? pattern.charCodeAt(nextIndex) : 126;
      if (h >= haystack.length || haystack.charCodeAt(h) != nextChar) {
        return false;
      }
      p = nextIndex < pattern.length ? nextIndex + 1 : nextIndex;
      h += 1;
      continue;
    }
    if (char == 42) {
      let nextPatternIndex = p + 1;
      while (nextPatternIndex < pattern.length && pattern.charCodeAt(nextPatternIndex) == 42) {
        nextPatternIndex += 1;
      }
      if (nextPatternIndex >= pattern.length) {
        return true;
      }
      for (let scan = h; scan <= haystack.length; scan++) {
        if (wildcardMatchAt(pattern, haystack, nextPatternIndex, scan)) {
          return true;
        }
      }
      return false;
    }
    if (h >= haystack.length) {
      return false;
    }
    if (char == 63) {
      p += 1;
      h += 1;
      continue;
    }
    if (haystack.charCodeAt(h) != char) {
      return false;
    }
    p += 1;
    h += 1;
  }
  return true;
}

function findPosition(
  needle: string,
  haystack: string,
  start: i32,
  caseSensitive: bool,
  wildcardAware: bool,
): i32 {
  const startIndex = start - 1;
  if (needle.length == 0) {
    return start;
  }
  if (startIndex > haystack.length) {
    return i32.MIN_VALUE;
  }
  const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase();
  const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase();
  if (wildcardAware && hasSearchSyntax(normalizedNeedle)) {
    for (let index = startIndex; index <= normalizedHaystack.length; index++) {
      if (wildcardMatchAt(normalizedNeedle, normalizedHaystack, 0, index)) {
        return index + 1;
      }
    }
    return i32.MIN_VALUE;
  }
  const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex);
  return found < 0 ? i32.MIN_VALUE : found + 1;
}

const BAHTTEXT_MAX_SATANG: f64 = 9007199254740991.0;

function trimLeadingZerosText(text: string): string {
  let index = 0;
  while (index + 1 < text.length && text.charCodeAt(index) == 48) {
    index += 1;
  }
  return text.slice(index);
}

function isAllZeroText(text: string): bool {
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) != 48) {
      return false;
    }
  }
  return true;
}

function bahtDigitWord(digit: i32): string {
  switch (digit) {
    case 0:
      return "ศูนย์";
    case 1:
      return "หนึ่ง";
    case 2:
      return "สอง";
    case 3:
      return "สาม";
    case 4:
      return "สี่";
    case 5:
      return "ห้า";
    case 6:
      return "หก";
    case 7:
      return "เจ็ด";
    case 8:
      return "แปด";
    case 9:
      return "เก้า";
    default:
      return "";
  }
}

function bahtScaleWord(position: i32): string {
  switch (position) {
    case 1:
      return "สิบ";
    case 2:
      return "ร้อย";
    case 3:
      return "พัน";
    case 4:
      return "หมื่น";
    case 5:
      return "แสน";
    default:
      return "";
  }
}

function bahtSegmentText(text: string): string {
  const normalized = trimLeadingZerosText(text);
  if (normalized.length == 0 || isAllZeroText(normalized)) {
    return "";
  }

  let output = "";
  let hasHigherNonZero = false;
  const length = normalized.length;
  for (let index = 0; index < length; index += 1) {
    const digit = <i32>(normalized.charCodeAt(index) - 48);
    if (digit == 0) {
      continue;
    }
    const position = length - index - 1;
    if (position == 0) {
      output += digit == 1 && hasHigherNonZero ? "เอ็ด" : bahtDigitWord(digit);
    } else if (position == 1) {
      if (digit == 1) {
        output += "สิบ";
      } else if (digit == 2) {
        output += "ยี่สิบ";
      } else {
        output += bahtDigitWord(digit) + "สิบ";
      }
    } else {
      output += bahtDigitWord(digit) + bahtScaleWord(position);
    }
    hasHigherNonZero = true;
  }
  return output;
}

function bahtIntegerText(text: string): string {
  const normalized = trimLeadingZerosText(text);
  if (normalized.length == 0 || isAllZeroText(normalized)) {
    return "ศูนย์";
  }
  if (normalized.length > 6) {
    const split = normalized.length - 6;
    return (
      bahtIntegerText(normalized.slice(0, split)) + "ล้าน" + bahtSegmentText(normalized.slice(split))
    );
  }
  const segment = bahtSegmentText(normalized);
  return segment.length == 0 ? "ศูนย์" : segment;
}

function bahtTextFromNumber(value: f64): string | null {
  if (!isFinite(value)) {
    return null;
  }
  const absolute = Math.abs(value);
  const scaled = Math.round(absolute * 100.0);
  if (!isFinite(scaled) || scaled > BAHTTEXT_MAX_SATANG) {
    return null;
  }
  const scaledSatang = <i64>scaled;
  const baht = scaledSatang / 100;
  const satang = <i32>(scaledSatang % 100);
  const prefix = value < 0.0 ? "ลบ" : "";
  const bahtText = bahtIntegerText(baht.toString());
  if (satang == 0) {
    return prefix + bahtText + "บาทถ้วน";
  }
  return prefix + bahtText + "บาท" + bahtSegmentText(satang.toString()) + "สตางค์";
}

function removeAsciiWhitespace(input: string): string {
  let result = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input.charCodeAt(index);
    if (char <= 32) {
      continue;
    }
    result += String.fromCharCode(char);
  }
  return result;
}

function numberValueParseText(
  input: string,
  decimalSeparator: string,
  groupSeparator: string,
): f64 {
  const compact = removeAsciiWhitespace(input);
  if (compact.length == 0) {
    return 0.0;
  }

  let percentCount = 0;
  let coreEnd = compact.length;
  while (coreEnd > 0 && compact.charCodeAt(coreEnd - 1) == 37) {
    coreEnd -= 1;
    percentCount += 1;
  }
  const core = coreEnd == compact.length ? compact : compact.slice(0, coreEnd);
  if (core.indexOf("%") >= 0) {
    return NaN;
  }

  const decimal =
    decimalSeparator.length == 0 ? "." : String.fromCharCode(decimalSeparator.charCodeAt(0));
  const group = groupSeparator.length == 0 ? "" : String.fromCharCode(groupSeparator.charCodeAt(0));
  if (decimal.length > 0 && group.length > 0 && decimal == group) {
    return NaN;
  }

  const decimalIndex = decimal.length == 0 ? -1 : core.indexOf(decimal);
  if (decimalIndex >= 0 && core.indexOf(decimal, decimalIndex + decimal.length) >= 0) {
    return NaN;
  }

  let normalized = "";
  for (let index = 0; index < core.length; index += 1) {
    const char = String.fromCharCode(core.charCodeAt(index));
    if (group.length > 0 && char == group) {
      if (decimalIndex >= 0 && index > decimalIndex) {
        return NaN;
      }
      continue;
    }
    if (decimal.length > 0 && char == decimal) {
      normalized += ".";
      continue;
    }
    normalized += char;
  }
  if (normalized == "" || normalized == "." || normalized == "+" || normalized == "-") {
    return NaN;
  }

  const parsed = parseNumericText(normalized);
  if (!isFinite(parsed)) {
    return NaN;
  }
  return parsed / Math.pow(100.0, <f64>percentCount);
}

function errorLabel(code: i32): string {
  if (code == ErrorCode.Div0) return "#DIV/0!";
  if (code == ErrorCode.Ref) return "#REF!";
  if (code == ErrorCode.Value) return "#VALUE!";
  if (code == ErrorCode.Name) return "#NAME?";
  if (code == ErrorCode.NA) return "#N/A";
  if (code == ErrorCode.Cycle) return "#CYCLE!";
  if (code == ErrorCode.Spill) return "#SPILL!";
  if (code == ErrorCode.Blocked) return "#BLOCKED!";
  return "#ERROR!";
}

function hexDigit(value: i32): string {
  return String.fromCharCode(value < 10 ? 48 + value : 55 + value);
}

function jsonQuoteText(input: string): string {
  let result = '"';
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    if (code == 34) {
      result += '\\"';
    } else if (code == 92) {
      result += "\\\\";
    } else if (code == 8) {
      result += "\\b";
    } else if (code == 12) {
      result += "\\f";
    } else if (code == 10) {
      result += "\\n";
    } else if (code == 13) {
      result += "\\r";
    } else if (code == 9) {
      result += "\\t";
    } else if (code < 32) {
      result += "\\u00" + hexDigit((code >> 4) & 0xf) + hexDigit(code & 0xf);
    } else {
      result += String.fromCharCode(code);
    }
  }
  return result + '"';
}

function isWeekendSerial(whole: i32): bool {
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const dayOfWeek = ((adjustedWhole % 7) + 7) % 7;
  return dayOfWeek == 0 || dayOfWeek == 6;
}

function weekendSerialDay(whole: i32): i32 {
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return ((adjustedWhole % 7) + 7) % 7;
}

function weekendMaskFromCode(code: i32): i32 {
  if (code == 1) return (1 << 6) | (1 << 0);
  if (code == 2) return (1 << 0) | (1 << 1);
  if (code == 3) return (1 << 1) | (1 << 2);
  if (code == 4) return (1 << 2) | (1 << 3);
  if (code == 5) return (1 << 3) | (1 << 4);
  if (code == 6) return (1 << 4) | (1 << 5);
  if (code == 7) return (1 << 5) | (1 << 6);
  if (code >= 11 && code <= 17) {
    return 1 << (code == 17 ? 0 : code - 10);
  }
  return i32.MIN_VALUE;
}

function weekendMaskFromString(maskText: string): i32 {
  const trimmed = trimAsciiWhitespace(maskText);
  if (trimmed.length != 7) {
    return i32.MIN_VALUE;
  }
  let mask = 0;
  for (let index = 0; index < 7; index += 1) {
    const char = trimmed.charCodeAt(index);
    if (char != 48 && char != 49) {
      return i32.MIN_VALUE;
    }
    if (char != 49) {
      continue;
    }
    const day = index == 6 ? 0 : index + 1;
    mask |= 1 << day;
  }
  return mask != 0x7f ? mask : i32.MIN_VALUE;
}

function coerceWeekendMask(
  hasWeekendArg: bool,
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (!hasWeekendArg) {
    return (1 << 0) | (1 << 6);
  }
  if (tag == ValueTag.String) {
    const maskText = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    return maskText == null ? i32.MIN_VALUE : weekendMaskFromString(maskText);
  }
  const code = coerceInteger(tag, value);
  return code == i32.MIN_VALUE ? i32.MIN_VALUE : weekendMaskFromCode(code);
}

function isWeekendWithMask(serial: i32, weekendMask: i32): bool {
  return (weekendMask & (1 << weekendSerialDay(serial))) != 0;
}

function isHolidaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (kind == STACK_KIND_SCALAR) {
    if (tag == ValueTag.Error) {
      return -1;
    }
    const holiday = truncToInt(tag, value);
    return holiday == i32.MIN_VALUE ? -1 : holiday == serial ? 1 : 0;
  }
  if (kind != STACK_KIND_RANGE) {
    return 0;
  }
  const start = <i32>rangeOffsets[rangeIndex];
  const length = <i32>rangeLengths[rangeIndex];
  for (let index = 0; index < length; index += 1) {
    const memberIndex = rangeMembers[start + index];
    if (cellTags[memberIndex] == ValueTag.Error) {
      return -1;
    }
    const serialCandidate = truncToInt(
      cellTags[memberIndex],
      memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    );
    if (serialCandidate == i32.MIN_VALUE) {
      return -1;
    }
    if (serialCandidate == serial) {
      return 1;
    }
  }
  return 0;
}

function isWorkdaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (isWeekendSerial(serial)) {
    return 0;
  }
  if (kind == STACK_KIND_SCALAR && tag == ValueTag.Empty) {
    return 1;
  }
  if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE) {
    return 1;
  }
  const holiday = isHolidaySerial(
    serial,
    kind,
    tag,
    value,
    rangeIndex,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  );
  if (holiday < 0) {
    return -1;
  }
  return holiday == 1 ? 0 : 1;
}

function isWorkdaySerialWithWeekendMask(
  serial: i32,
  weekendMask: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (isWeekendWithMask(serial, weekendMask)) {
    return 0;
  }
  if (kind == STACK_KIND_SCALAR && tag == ValueTag.Empty) {
    return 1;
  }
  if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE) {
    return 1;
  }
  const holiday = isHolidaySerial(
    serial,
    kind,
    tag,
    value,
    rangeIndex,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  );
  if (holiday < 0) {
    return -1;
  }
  return holiday == 1 ? 0 : 1;
}

function coerceNonNegativeLength(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

function coerceInteger(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  return <i32>numeric;
}

function coerceTrimMode(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  const integer = <i32>numeric;
  return integer >= 0 && integer <= 3 ? integer : i32.MIN_VALUE;
}

function clipIndex(index: i32, length: i32): i32 {
  if (length <= 0) {
    return i32.MIN_VALUE;
  }
  if (index == 0) {
    return i32.MIN_VALUE;
  }
  return index < 0 ? max(index, -length) : min(index, length);
}

function inputRowsFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
): i32 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeRowCounts[rangeIndexStack[slot]];
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayRows(rangeIndexStack[slot]);
  }
  return i32.MIN_VALUE;
}

function inputColsFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeColCounts[rangeIndexStack[slot]];
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayCols(rangeIndexStack[slot]);
  }
  return i32.MIN_VALUE;
}

function inputCellTag(
  slot: i32,
  row: i32,
  col: i32,
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
): u8 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return <u8>ValueTag.Error;
  }
  if (kind == STACK_KIND_SCALAR) {
    if (row != 0 || col != 0) {
      return <u8>ValueTag.Error;
    }
    return tagStack[slot];
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return <u8>ValueTag.Error;
    }
    return cellTags[memberIndex];
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return <u8>ValueTag.Error;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return <u8>ValueTag.Error;
    }
    return readSpillArrayTag(arrayIndex, arrayOffset);
  }
  return <u8>ValueTag.Error;
}

function inputCellScalarValue(
  slot: i32,
  row: i32,
  col: i32,
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
): f64 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return NaN;
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? valueStack[slot] : NaN;
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return NaN;
    }
    return memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors);
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN;
    }
    return readSpillArrayNumber(arrayIndex, arrayOffset);
  }
  return NaN;
}

function inputCellNumeric(
  slot: i32,
  row: i32,
  col: i32,
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
): f64 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return NaN;
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? toNumberOrNaN(tagStack[slot], valueStack[slot]) : NaN;
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return NaN;
    }
    return toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN;
    }
    return toNumberOrNaN(
      readSpillArrayTag(arrayIndex, arrayOffset),
      readSpillArrayNumber(arrayIndex, arrayOffset),
    );
  }
  return NaN;
}

let pairedSampleCount: i32 = 0;
let pairedSumX: f64 = 0;
let pairedSumY: f64 = 0;
let pairedSumXX: f64 = 0;
let pairedSumYY: f64 = 0;
let pairedSumXY: f64 = 0;

function normalizeNearZero(value: f64): f64 {
  return Math.abs(value) < 1e-12 ? 0.0 : value;
}

function collectPairedNumericStats(
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

function pairedCenteredSumSquaresX(): f64 {
  return normalizeNearZero(pairedSumXX - (pairedSumX * pairedSumX) / <f64>pairedSampleCount);
}

function pairedCenteredSumSquaresY(): f64 {
  return normalizeNearZero(pairedSumYY - (pairedSumY * pairedSumY) / <f64>pairedSampleCount);
}

function pairedCenteredCrossProducts(): f64 {
  return normalizeNearZero(pairedSumXY - (pairedSumX * pairedSumY) / <f64>pairedSampleCount);
}

function copyInputCellToSpill(
  arrayIndex: u32,
  outputOffset: i32,
  slot: i32,
  row: i32,
  col: i32,
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
  const value = inputCellScalarValue(
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
  if (tag == ValueTag.Error && isNaN(value)) {
    return ErrorCode.Value;
  }
  writeSpillArrayValue(arrayIndex, outputOffset, tag, value);
  return ErrorCode.None;
}

function materializeSlotResult(
  base: i32,
  sourceSlot: i32,
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
  if (kindStack[sourceSlot] == STACK_KIND_SCALAR) {
    return copySlotResult(base, sourceSlot, rangeIndexStack, valueStack, tagStack, kindStack);
  }
  const rows = inputRowsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeRowCounts);
  const cols = inputColsFromSlot(sourceSlot, kindStack, rangeIndexStack, rangeColCounts);
  if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || rows <= 0 || cols <= 0) {
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
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const copyError = copyInputCellToSpill(
        arrayIndex,
        outputOffset,
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
      if (copyError != ErrorCode.None) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          copyError,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
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

function uniqueScalarKey(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Empty) {
    return "E:";
  }
  if (tag == ValueTag.Number) {
    return "N:" + value.toString();
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? "B:1" : "B:0";
  }
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    return text == null ? null : "S:" + text.toUpperCase();
  }
  return null;
}

function uniqueRowKey(
  slot: i32,
  row: i32,
  cols: i32,
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  let key = "";
  for (let col = 0; col < cols; col++) {
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
    const value = inputCellScalarValue(
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
      return null;
    }
    const cellKey = uniqueScalarKey(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (cellKey == null) {
      return null;
    }
    if (col > 0) {
      key += "\u0001";
    }
    key += cellKey;
  }
  return key;
}

function uniqueColKey(
  slot: i32,
  col: i32,
  rows: i32,
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  let key = "";
  for (let row = 0; row < rows; row++) {
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
    const value = inputCellScalarValue(
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
      return null;
    }
    const cellKey = uniqueScalarKey(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (cellKey == null) {
      return null;
    }
    if (row > 0) {
      key += "\u0001";
    }
    key += cellKey;
  }
  return key;
}

function replaceText(text: string, start: i32, count: i32, replacement: string): string {
  const startIndex = start - 1;
  if (startIndex >= text.length) {
    return text;
  }
  return text.slice(0, startIndex) + replacement + text.slice(startIndex + count);
}

function substituteText(text: string, oldText: string, newText: string): string {
  if (oldText.length == 0) {
    return text;
  }
  let result = "";
  let searchIndex = 0;
  let found = text.indexOf(oldText, searchIndex);
  if (found < 0) {
    return text;
  }
  while (found >= 0) {
    result += text.slice(searchIndex, found) + newText;
    searchIndex = found + oldText.length;
    found = text.indexOf(oldText, searchIndex);
  }
  return result + text.slice(searchIndex);
}

function substituteNthText(text: string, oldText: string, newText: string, instance: i32): string {
  let count = 0;
  let searchIndex = 0;
  while (searchIndex <= text.length) {
    const found = text.indexOf(oldText, searchIndex);
    if (found < 0) {
      return text;
    }
    count += 1;
    if (count == instance) {
      return text.slice(0, found) + newText + text.slice(found + oldText.length);
    }
    searchIndex = found + oldText.length;
  }
  return text;
}

function repeatText(text: string, count: i32): string {
  let result = "";
  for (let index = 0; index < count; index += 1) {
    result += text;
  }
  return result;
}

function indexOfTextWithMode(
  text: string,
  delimiter: string,
  searchFrom: i32,
  matchMode: i32,
): i32 {
  const normalizedText = matchMode == 1 ? text.toLowerCase() : text;
  const normalizedDelimiter = matchMode == 1 ? delimiter.toLowerCase() : delimiter;
  return normalizedText.indexOf(normalizedDelimiter, max<i32>(0, searchFrom));
}

function lastIndexOfTextWithMode(
  text: string,
  delimiter: string,
  searchFrom: i32,
  matchMode: i32,
): i32 {
  const normalizedText = matchMode == 1 ? text.toLowerCase() : text;
  const normalizedDelimiter = matchMode == 1 ? delimiter.toLowerCase() : delimiter;
  let start = min<i32>(searchFrom, normalizedText.length - normalizedDelimiter.length);
  if (start < 0) {
    return -1;
  }
  while (start >= 0) {
    if (normalizedText.slice(start, start + normalizedDelimiter.length) == normalizedDelimiter) {
      return start;
    }
    start -= 1;
  }
  return -1;
}

function splitTextByDelimiterWithMode(
  text: string,
  delimiter: string,
  matchMode: i32,
): Array<string> {
  const parts = new Array<string>();
  if (delimiter.length == 0) {
    parts.push(text);
    return parts;
  }
  let cursor = 0;
  while (cursor <= text.length) {
    const found = indexOfTextWithMode(text, delimiter, cursor, matchMode);
    if (found < 0) {
      parts.push(text.slice(cursor));
      break;
    }
    parts.push(text.slice(cursor, found));
    cursor = found + delimiter.length;
  }
  return parts;
}

function valueNumber(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag != ValueTag.String) {
    return NaN;
  }
  const text = scalarText(
    tag,
    value,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  return text == null ? NaN : parseNumericText(text);
}

const CRITERIA_OP_EQ: i32 = 0;
const CRITERIA_OP_NE: i32 = 1;
const CRITERIA_OP_GT: i32 = 2;
const CRITERIA_OP_GTE: i32 = 3;
const CRITERIA_OP_LT: i32 = 4;
const CRITERIA_OP_LTE: i32 = 5;
const AXIS_AGG_SUM: i32 = 1;
const AXIS_AGG_AVERAGE: i32 = 2;
const AXIS_AGG_MIN: i32 = 3;
const AXIS_AGG_MAX: i32 = 4;
const AXIS_AGG_COUNT: i32 = 5;
const AXIS_AGG_COUNTA: i32 = 6;

function compareScalarValues(
  leftTag: u8,
  leftValue: f64,
  rightTag: u8,
  rightValue: f64,
  rightText: string | null,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const leftTextlike = leftTag == ValueTag.String || leftTag == ValueTag.Empty;
  const rightTextlike = rightTag == ValueTag.String || rightTag == ValueTag.Empty;
  if (leftTextlike && rightTextlike) {
    const leftText = scalarText(
      leftTag,
      leftValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const resolvedRightText =
      rightText != null
        ? rightText
        : scalarText(
            rightTag,
            rightValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
    if (leftText == null || resolvedRightText == null) {
      return i32.MIN_VALUE;
    }
    const normalizedLeft = leftText.toUpperCase();
    const normalizedRight = resolvedRightText.toUpperCase();
    if (normalizedLeft == normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNumeric = toNumberOrNaN(leftTag, leftValue);
  const rightNumeric = toNumberOrNaN(rightTag, rightValue);
  if (isNaN(leftNumeric) || isNaN(rightNumeric)) {
    return i32.MIN_VALUE;
  }
  if (leftNumeric == rightNumeric) {
    return 0;
  }
  return leftNumeric < rightNumeric ? -1 : 1;
}

function matchesCriteriaValue(
  valueTag: u8,
  valueValue: f64,
  criteriaTag: u8,
  criteriaValue: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): bool {
  if (valueTag == ValueTag.Error) {
    return false;
  }

  let operator = CRITERIA_OP_EQ;
  let operandTag = criteriaTag;
  let operandValue = criteriaValue;
  let operandText: string | null = null;

  if (criteriaTag == ValueTag.String) {
    const criteriaText = scalarText(
      criteriaTag,
      criteriaValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (criteriaText == null) {
      return false;
    }

    let rawOperand = criteriaText;
    let parsedOperator = false;
    if (criteriaText.length >= 2) {
      const prefix = criteriaText.slice(0, 2);
      if (prefix == "<=") {
        operator = CRITERIA_OP_LTE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      } else if (prefix == ">=") {
        operator = CRITERIA_OP_GTE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      } else if (prefix == "<>") {
        operator = CRITERIA_OP_NE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      }
    }
    if (!parsedOperator && criteriaText.length >= 1) {
      const first = criteriaText.charCodeAt(0);
      if (first == 61) {
        operator = CRITERIA_OP_EQ;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      } else if (first == 62) {
        operator = CRITERIA_OP_GT;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      } else if (first == 60) {
        operator = CRITERIA_OP_LT;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      }
    }

    if (!parsedOperator) {
      operandText = criteriaText;
    } else {
      const trimmed = trimAsciiWhitespace(rawOperand);
      if (trimmed.length == 0) {
        operandTag = <u8>ValueTag.String;
        operandValue = 0;
        operandText = "";
      } else {
        const upper = trimmed.toUpperCase();
        if (upper == "TRUE" || upper == "FALSE") {
          operandTag = <u8>ValueTag.Boolean;
          operandValue = upper == "TRUE" ? 1 : 0;
        } else {
          const numeric = parseNumericText(trimmed);
          if (!isNaN(numeric)) {
            operandTag = <u8>ValueTag.Number;
            operandValue = numeric;
          } else {
            operandTag = <u8>ValueTag.String;
            operandValue = 0;
            operandText = trimmed;
          }
        }
      }
    }
  }

  const comparison = compareScalarValues(
    valueTag,
    valueValue,
    operandTag,
    operandValue,
    operandText,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (comparison == i32.MIN_VALUE) {
    return false;
  }
  if (operator == CRITERIA_OP_EQ) {
    return comparison == 0;
  }
  if (operator == CRITERIA_OP_NE) {
    return comparison != 0;
  }
  if (operator == CRITERIA_OP_GT) {
    return comparison > 0;
  }
  if (operator == CRITERIA_OP_GTE) {
    return comparison >= 0;
  }
  if (operator == CRITERIA_OP_LT) {
    return comparison < 0;
  }
  return comparison <= 0;
}

function memberScalarValue(
  memberIndex: u32,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  const tag = cellTags[memberIndex];
  if (tag == ValueTag.String) {
    return <f64>cellStringIds[memberIndex];
  }
  if (tag == ValueTag.Error) {
    return <f64>cellErrors[memberIndex];
  }
  return cellNumbers[memberIndex];
}

function rangeMemberAt(
  rangeIndex: u32,
  row: i32,
  col: i32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): u32 {
  if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
    return 0xffffffff;
  }
  const rowCount = <i32>rangeRowCounts[rangeIndex];
  const colCount = <i32>rangeColCounts[rangeIndex];
  const length = <i32>rangeLengths[rangeIndex];
  if (
    rowCount <= 0 ||
    colCount <= 0 ||
    row < 0 ||
    col < 0 ||
    row >= rowCount ||
    col >= colCount ||
    row * colCount + col >= length
  ) {
    return 0xffffffff;
  }
  return rangeMembers[rangeOffsets[rangeIndex] + row * colCount + col];
}

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

function unresolvedRangeOperandError(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_RANGE && rangeIndexStack[slot] == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
  }
  return -1;
}

function databaseBuiltinResult(
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

  const databaseRangeIndex = rangeIndexStack[databaseSlot];
  const criteriaRangeIndex = rangeIndexStack[criteriaSlot];
  const databaseRows = <i32>rangeRowCounts[databaseRangeIndex];
  const databaseCols = <i32>rangeColCounts[databaseRangeIndex];
  const criteriaRows = <i32>rangeRowCounts[criteriaRangeIndex];
  const criteriaCols = <i32>rangeColCounts[criteriaRangeIndex];
  if (databaseRows < 1 || databaseCols < 1 || criteriaRows < 2 || criteriaCols < 1) {
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      fieldValue,
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
  } else if (fieldTag == ValueTag.Number) {
    const position = truncToInt(<u8>fieldTag, fieldValue);
    if (position == i32.MIN_VALUE || position < 1 || position > databaseCols) {
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
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            headerValue,
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

  for (let databaseRow = 1; databaseRow < databaseRows; databaseRow++) {
    let matchesAnyCriteriaRow = false;
    for (let criteriaRow = 1; criteriaRow < criteriaRows; criteriaRow++) {
      let blocked = false;
      let matchesAll = true;
      for (let criteriaCol = 0; criteriaCol < criteriaCols; criteriaCol++) {
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
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            criteriaValue,
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
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            headerValue,
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
        for (let databaseCol = 0; databaseCol < databaseCols; databaseCol++) {
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
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              databaseHeaderValue,
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
        const databaseTag = cellTags[databaseMemberIndex];
        const databaseValue = memberScalarValue(
          databaseMemberIndex,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (
          !matchesCriteriaValue(
            databaseTag,
            databaseValue,
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

    const numeric = toNumberOrNaN(memberTag, memberValue);
    if (isNaN(numeric)) {
      continue;
    }
    numericCount += 1;
    sum += numeric;
    sumSquares += numeric * numeric;
    product *= numeric;
    minimum = min(minimum, numeric);
    maximum = max(maximum, numeric);
    hasNumeric = true;
  }

  if (builtinId == BuiltinId.Dcount) {
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
  if (builtinId == BuiltinId.Dcounta) {
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
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          sum / numericCount,
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
          : 0
        : builtinId == BuiltinId.Dmin
          ? hasNumeric
            ? minimum
            : 0
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
      numericCount == 0 ? 0 : product,
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      ErrorCode.Div0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const mean = sum / numericCount;
  let variance = sumSquares - numericCount * mean * mean;
  variance /=
    builtinId == BuiltinId.Dstdev || builtinId == BuiltinId.Dvar ? numericCount - 1 : numericCount;
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

function writeMemberResult(
  base: i32,
  memberIndex: u32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    cellTags[memberIndex],
    memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

function writeResult(
  base: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  rangeIndexStack[base] = 0;
  valueStack[base] = value;
  tagStack[base] = tag;
  kindStack[base] = kind;
  return base + 1;
}

function writeArrayResult(
  base: i32,
  arrayIndex: u32,
  rows: i32,
  cols: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  registerTrackedArrayShapeImpl(arrayIndex, rows, cols);
  rangeIndexStack[base] = arrayIndex;
  valueStack[base] = 0;
  tagStack[base] = ValueTag.Empty;
  kindStack[base] = STACK_KIND_ARRAY;
  return base + 1;
}

function copySlotResult(
  base: i32,
  sourceSlot: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  kindStack[base] = kindStack[sourceSlot];
  tagStack[base] = tagStack[sourceSlot];
  valueStack[base] = valueStack[sourceSlot];
  rangeIndexStack[base] = rangeIndexStack[sourceSlot];
  return base + 1;
}

function vectorSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (kindStack[slot] == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kindStack[slot] != STACK_KIND_RANGE) {
    return i32.MIN_VALUE;
  }
  const rangeIndex = rangeIndexStack[slot];
  const rowCount = <i32>rangeRowCounts[rangeIndex];
  const colCount = <i32>rangeColCounts[rangeIndex];
  if (rowCount <= 0 || colCount <= 0 || (rowCount != 1 && colCount != 1)) {
    return i32.MIN_VALUE;
  }
  return <i32>rangeLengths[rangeIndex];
}

function coercePositiveIntegerArg(tag: u8, value: f64, hasValue: bool, fallback: i32): i32 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

function coerceNumberArg(tag: u8, value: f64, hasValue: bool, fallback: f64): f64 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  return isFinite(numeric) ? numeric : NaN;
}

function scalarErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  valueStack: Float64Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR && tagStack[slot] == ValueTag.Error) {
      return valueStack[slot];
    }
  }
  return -1;
}

function rangeErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellErrors: Uint16Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] != STACK_KIND_RANGE) {
      continue;
    }
    const rangeIndex = rangeIndexStack[slot];
    if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    for (let cursor = 0; cursor < length; cursor++) {
      const memberIndex = rangeMembers[start + cursor];
      if (cellTags[memberIndex] == ValueTag.Error) {
        return cellErrors[memberIndex];
      }
    }
  }
  return -1;
}

function rangeSupportedScalarOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] == STACK_KIND_RANGE) {
      return false;
    }
  }
  return true;
}

function coerceLogical(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag == ValueTag.Error) {
    return -(<i32>value) - 1;
  }
  return -(<i32>ErrorCode.Value) - 1;
}

function scalarArgsOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] != STACK_KIND_SCALAR) {
      return false;
    }
  }
  return true;
}

function statScalarValue(tag: u8, value: f64, includeStringsAsZero: bool): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.String) {
    return includeStringsAsZero ? 0 : NaN;
  }
  return NaN;
}

function collectScalarStatValues(
  base: i32,
  argc: i32,
  tagStack: Uint8Array,
  valueStack: Float64Array,
  includeStringsAsZero: bool,
): Array<f64> {
  const values = new Array<f64>();
  for (let index = 0; index < argc; index++) {
    const numeric = statScalarValue(
      tagStack[base + index],
      valueStack[base + index],
      includeStringsAsZero,
    );
    if (!isNaN(numeric)) {
      values.push(numeric);
    }
  }
  return values;
}

let statCollectionErrorCode = 0;

function collectStatValuesFromArgs(
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
  includeStringsAsZero: bool,
): Array<f64> | null {
  statCollectionErrorCode = 0;
  const values = new Array<f64>();
  for (let index = 0; index < argc; index += 1) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR) {
      const tag = tagStack[slot];
      const raw = valueStack[slot];
      if (tag == ValueTag.Error) {
        statCollectionErrorCode = <i32>raw;
        return null;
      }
      if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
        values.push(raw);
      } else if (tag == ValueTag.String && includeStringsAsZero) {
        values.push(0.0);
      }
      continue;
    }

    const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
    const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
    if (rows < 1 || cols < 1) {
      statCollectionErrorCode = ErrorCode.Value;
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
          statCollectionErrorCode = <i32>raw;
          return null;
        }
        if (tag == ValueTag.Number) {
          values.push(raw);
        } else if (tag == ValueTag.Boolean && includeStringsAsZero) {
          values.push(raw);
        } else if (tag == ValueTag.String && includeStringsAsZero) {
          values.push(0.0);
        }
      }
    }
  }
  return values;
}

function paymentType(tag: u8, value: f64, hasValue: bool): i32 {
  if (!hasValue) {
    return 0;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return -1;
  }
  const truncated = <i32>numeric;
  return truncated == 0 || truncated == 1 ? truncated : -1;
}

function chiSquareTestPValue(
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

let sampleCollectionErrorCode = 0;

function collectSampleNumbersFromSlot(
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

function collectNumericSeriesFromSlot(
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

function collectDateSeriesFromSlot(
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

function collectNumericCellRangeSeriesFromSlot(
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

function collectDateCellRangeSeriesFromSlot(
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

let orderStatisticErrorCode = 0;

function collectNumericValuesFromSlot(
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

function collectNumericValuesFromArgs(
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

function fTestPValue(
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

function zTestPValue(
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

function tTestPValue(
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

function isNumericResult(value: f64): bool {
  return !isNaN(value);
}

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
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
  sp: i32,
): i32 {
  const base = sp - argc;
  const unresolvedRangeError = unresolvedRangeOperandError(base, argc, kindStack, rangeIndexStack);
  if (unresolvedRangeError >= 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      unresolvedRangeError,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Irr && (argc == 1 || argc == 2)) ||
    (builtinId == BuiltinId.Mirr && argc == 3)
  ) {
    const values = collectNumericCellRangeSeriesFromSlot(
      base,
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
      false,
    );
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (!hasPositiveAndNegativeSeries(values)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        builtinId == BuiltinId.Mirr ? ErrorCode.Div0 : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (builtinId == BuiltinId.Irr) {
      if (argc == 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
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
      const guess = argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 0.1;
      const result = isNaN(guess) ? NaN : solvePeriodicCashflowRateCalc(values, guess);
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
        isNaN(result) ? ErrorCode.Value : result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const financeRate =
      kindStack[base + 1] == STACK_KIND_SCALAR
        ? toNumberExact(tagStack[base + 1], valueStack[base + 1])
        : NaN;
    const reinvestRate =
      kindStack[base + 2] == STACK_KIND_SCALAR
        ? toNumberExact(tagStack[base + 2], valueStack[base + 2])
        : NaN;
    const result = mirrCalc(values, financeRate, reinvestRate);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(result)
        ? isNaN(financeRate) || isNaN(reinvestRate)
          ? ErrorCode.Value
          : ErrorCode.Div0
        : result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Xnpv && argc == 3) ||
    (builtinId == BuiltinId.Xirr && (argc == 2 || argc == 3))
  ) {
    if (
      (builtinId == BuiltinId.Xnpv && kindStack[base] != STACK_KIND_SCALAR) ||
      (builtinId == BuiltinId.Xirr && kindStack[base] == STACK_KIND_SCALAR)
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
    const firstNumeric =
      builtinId == BuiltinId.Xnpv ? toNumberExact(tagStack[base], valueStack[base]) : NaN;
    const guess =
      builtinId == BuiltinId.Xirr
        ? argc == 3
          ? kindStack[base + 2] == STACK_KIND_SCALAR
            ? toNumberExact(tagStack[base + 2], valueStack[base + 2])
            : NaN
          : 0.1
        : NaN;
    if (
      (builtinId == BuiltinId.Xnpv && isNaN(firstNumeric)) ||
      (builtinId == BuiltinId.Xirr && isNaN(guess))
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
    const valuesSlot = builtinId == BuiltinId.Xnpv ? base + 1 : base;
    const datesSlot = builtinId == BuiltinId.Xnpv ? base + 2 : base + 1;
    const values = collectNumericCellRangeSeriesFromSlot(
      valuesSlot,
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const dates = collectDateCellRangeSeriesFromSlot(
      datesSlot,
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
    );
    if (dates === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      values.length != dates.length ||
      values.length == 0 ||
      !hasPositiveAndNegativeSeries(values)
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
    const start = unchecked(dates[0]);
    for (let index = 0; index < dates.length; index += 1) {
      if (unchecked(dates[index]) < start) {
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
    }
    const result =
      builtinId == BuiltinId.Xnpv
        ? xnpvCalc(firstNumeric, values, dates)
        : solveXirrCalc(values, dates, guess);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(result) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(result) ? ErrorCode.Value : result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == 233 || builtinId == 234 || builtinId == 235) && argc == 2) {
    const result = chiSquareTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == 236 || builtinId == 237) && argc == 2) {
    const result = fTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == 238 || builtinId == 239) && (argc == 2 || argc == 3)) {
    const result = zTestPValue(
      base,
      argc,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.TTest || builtinId == BuiltinId.Ttest) && argc == 4) {
    const result = tTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (result < 0.0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Today) {
    if (argc != 0)
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
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(nowSerial),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Now) {
    if (argc != 0)
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
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      nowSerial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rand) {
    if (argc != 0)
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
    const next = nextVolatileRandomValue();
    if (!isFinite(next))
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
    const bounded = Math.min(Math.max(next, 0), 1 - f64.EPSILON);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      bounded,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

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

  if (builtinId == BuiltinId.Filter && (argc == 2 || argc == 3)) {
    const sourceKind = kindStack[base];
    const includeKind = kindStack[base + 1];
    if (
      (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) ||
      (includeKind != STACK_KIND_RANGE && includeKind != STACK_KIND_ARRAY)
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
    const includeRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
    const includeCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      includeRows <= 0 ||
      includeCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE ||
      includeRows == i32.MIN_VALUE ||
      includeCols == i32.MIN_VALUE
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

    if (includeRows == sourceRows && includeCols == 1) {
      const keptRows = new Array<i32>();
      for (let row = 0; row < sourceRows; row++) {
        const includeTag = inputCellTag(
          base + 1,
          row,
          0,
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
        const includeValue = inputCellScalarValue(
          base + 1,
          row,
          0,
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
        if (includeTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            includeValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const keep = coerceBoolean(includeTag, includeValue);
        if (keep < 0) {
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
        if (keep == 1) {
          keptRows.push(row);
        }
      }
      if (keptRows.length == 0) {
        if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
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
          tagStack[base + 2],
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const arrayIndex = allocateSpillArrayResult(keptRows.length, sourceCols);
      let outputOffset = 0;
      for (let index = 0; index < keptRows.length; index++) {
        const sourceRow = unchecked(keptRows[index]);
        for (let col = 0; col < sourceCols; col++) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
            sourceRow,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        keptRows.length,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (includeRows == 1 && includeCols == sourceCols) {
      const keptCols = new Array<i32>();
      for (let col = 0; col < sourceCols; col++) {
        const includeTag = inputCellTag(
          base + 1,
          0,
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
        const includeValue = inputCellScalarValue(
          base + 1,
          0,
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
        if (includeTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            includeValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const keep = coerceBoolean(includeTag, includeValue);
        if (keep < 0) {
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
        if (keep == 1) {
          keptCols.push(col);
        }
      }
      if (keptCols.length == 0) {
        if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
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
          tagStack[base + 2],
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, keptCols.length);
      let outputOffset = 0;
      for (let row = 0; row < sourceRows; row++) {
        for (let index = 0; index < keptCols.length; index++) {
          const sourceCol = unchecked(keptCols[index]);
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
            row,
            sourceCol,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        keptCols.length,
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

  if (builtinId == BuiltinId.Unique && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
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
    if (argc >= 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (argc >= 3 && kindStack[base + 2] != STACK_KIND_SCALAR) {
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
    if (argc >= 2 && tagStack[base + 1] == ValueTag.Error) {
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
    if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const byColFlag = argc >= 2 ? coerceBoolean(tagStack[base + 1], valueStack[base + 1]) : 0;
    const exactlyOnceFlag = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (byColFlag < 0 || exactlyOnceFlag < 0) {
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

    if (sourceRows == 1 || sourceCols == 1) {
      const vectorLength = sourceRows * sourceCols;
      const keys = new Array<string>();
      for (let index = 0; index < vectorLength; index++) {
        const row = sourceRows == 1 ? 0 : index;
        const col = sourceRows == 1 ? index : 0;
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
        const value = inputCellScalarValue(
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
            value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const key = uniqueScalarKey(
          tag,
          value,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (key == null) {
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
        keys.push(key);
      }

      const keptIndexes = new Array<i32>();
      for (let index = 0; index < vectorLength; index++) {
        const key = unchecked(keys[index]);
        let seenEarlier = false;
        for (let prior = 0; prior < index; prior++) {
          if (unchecked(keys[prior]) == key) {
            seenEarlier = true;
            break;
          }
        }
        if (seenEarlier) {
          continue;
        }
        if (exactlyOnceFlag == 1) {
          let count = 0;
          for (let cursor = 0; cursor < vectorLength; cursor++) {
            if (unchecked(keys[cursor]) == key) {
              count += 1;
            }
          }
          if (count != 1) {
            continue;
          }
        }
        keptIndexes.push(index);
      }

      const outputRows = sourceRows == 1 ? 1 : keptIndexes.length;
      const outputCols = sourceRows == 1 ? keptIndexes.length : 1;
      const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
      for (let outputOffset = 0; outputOffset < keptIndexes.length; outputOffset++) {
        const sourceIndex = unchecked(keptIndexes[outputOffset]);
        const sourceRow = sourceRows == 1 ? 0 : sourceIndex;
        const sourceCol = sourceRows == 1 ? sourceIndex : 0;
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          sourceRow,
          sourceCol,
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
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
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

    if (byColFlag == 1) {
      const keys = new Array<string>();
      for (let col = 0; col < sourceCols; col++) {
        const key = uniqueColKey(
          base,
          col,
          sourceRows,
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
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (key == null) {
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
        keys.push(key);
      }

      const keptCols = new Array<i32>();
      for (let col = 0; col < sourceCols; col++) {
        const key = unchecked(keys[col]);
        let seenEarlier = false;
        for (let prior = 0; prior < col; prior++) {
          if (unchecked(keys[prior]) == key) {
            seenEarlier = true;
            break;
          }
        }
        if (seenEarlier) {
          continue;
        }
        if (exactlyOnceFlag == 1) {
          let count = 0;
          for (let cursor = 0; cursor < sourceCols; cursor++) {
            if (unchecked(keys[cursor]) == key) {
              count += 1;
            }
          }
          if (count != 1) {
            continue;
          }
        }
        keptCols.push(col);
      }

      const arrayIndex = allocateSpillArrayResult(sourceRows, keptCols.length);
      let outputOffset = 0;
      for (let row = 0; row < sourceRows; row++) {
        for (let index = 0; index < keptCols.length; index++) {
          const sourceCol = unchecked(keptCols[index]);
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            base,
            row,
            sourceCol,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        keptCols.length,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const keys = new Array<string>();
    for (let row = 0; row < sourceRows; row++) {
      const key = uniqueRowKey(
        base,
        row,
        sourceCols,
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
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (key == null) {
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
      keys.push(key);
    }

    const keptRows = new Array<i32>();
    for (let row = 0; row < sourceRows; row++) {
      const key = unchecked(keys[row]);
      let seenEarlier = false;
      for (let prior = 0; prior < row; prior++) {
        if (unchecked(keys[prior]) == key) {
          seenEarlier = true;
          break;
        }
      }
      if (seenEarlier) {
        continue;
      }
      if (exactlyOnceFlag == 1) {
        let count = 0;
        for (let cursor = 0; cursor < sourceRows; cursor++) {
          if (unchecked(keys[cursor]) == key) {
            count += 1;
          }
        }
        if (count != 1) {
          continue;
        }
      }
      keptRows.push(row);
    }

    const arrayIndex = allocateSpillArrayResult(keptRows.length, sourceCols);
    let outputOffset = 0;
    for (let index = 0; index < keptRows.length; index++) {
      const sourceRow = unchecked(keptRows[index]);
      for (let col = 0; col < sourceCols; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          sourceRow,
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
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      keptRows.length,
      sourceCols,
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

  if (builtinId == BuiltinId.Offset && argc >= 3 && argc <= 5) {
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
    const rowOffset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const colOffset = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const height =
      argc >= 4
        ? coercePositiveIntegerArg(tagStack[base + 3], valueStack[base + 3], true, sourceRows)
        : sourceRows;
    const width =
      argc >= 5
        ? coercePositiveIntegerArg(tagStack[base + 4], valueStack[base + 4], true, sourceCols)
        : sourceCols;
    const areaNumber = argc >= 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1;
    if (
      rowOffset == i32.MIN_VALUE ||
      colOffset == i32.MIN_VALUE ||
      height == i32.MIN_VALUE ||
      width == i32.MIN_VALUE ||
      areaNumber == i32.MIN_VALUE ||
      areaNumber != 1 ||
      height < 1 ||
      width < 1
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

    const rowStart = rowOffset < 0 ? sourceRows + rowOffset : rowOffset;
    const colStart = colOffset < 0 ? sourceCols + colOffset : colOffset;
    if (rowStart < 0 || colStart < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (rowStart + height > sourceRows || colStart + width > sourceCols) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (height == 1 && width == 1) {
      const result = inputCellNumeric(
        base,
        rowStart,
        colStart,
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
      if (isNaN(result)) {
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
        result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(height, width);
    for (let row = 0; row < height; row++) {
      for (let col = 0; col < width; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowStart + row,
          colStart + col,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, row * width + col, sourceValue);
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      height,
      width,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Take && argc >= 1 && argc <= 3) {
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
    const requestedRows =
      argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : sourceRows;
    const requestedCols =
      argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
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
    const clippedRows = argc >= 2 ? clipIndex(requestedRows, sourceRows) : sourceRows;
    const clippedCols = argc >= 3 ? clipIndex(requestedCols, sourceCols) : sourceCols;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
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
    const rowCount =
      clippedRows > 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colCount =
      clippedCols > 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    if (rowCount == 0 || colCount == 0) {
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
    const rowOffset = clippedRows > 0 ? 0 : max<i32>(sourceRows - rowCount, 0);
    const colOffset = clippedCols > 0 ? 0 : max<i32>(sourceCols - colCount, 0);

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Drop && argc >= 1 && argc <= 3) {
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
    const requestedRows = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const requestedCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
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
    const clippedRows =
      argc >= 2 ? (requestedRows == 0 ? 0 : clipIndex(requestedRows, sourceRows)) : 0;
    const clippedCols =
      argc >= 3 ? (requestedCols == 0 ? 0 : clipIndex(requestedCols, sourceCols)) : 0;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
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
    const rowDrop =
      clippedRows >= 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colDrop =
      clippedCols >= 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    const rowCount = sourceRows - rowDrop;
    const colCount = sourceCols - colDrop;
    if (rowCount <= 0 || colCount <= 0) {
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
    const rowOffset = clippedRows > 0 ? rowDrop : 0;
    const colOffset = clippedCols > 0 ? colDrop : 0;

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Expand && argc >= 2 && argc <= 4) {
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
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg;
      if (kindStack[slot] != STACK_KIND_SCALAR) {
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
    }
    if (tagStack[base + 1] == ValueTag.Error) {
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
    if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const targetRows = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    const targetCols =
      argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols;
    if (
      targetRows == i32.MIN_VALUE ||
      targetCols == i32.MIN_VALUE ||
      targetRows < sourceRows ||
      targetCols < sourceCols ||
      targetRows < 1 ||
      targetCols < 1
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
    const padTag = argc >= 4 ? tagStack[base + 3] : <u8>ValueTag.Error;
    const padValue = argc >= 4 ? valueStack[base + 3] : ErrorCode.NA;
    const arrayIndex = allocateSpillArrayResult(targetRows, targetCols);
    let outputOffset = 0;
    for (let row = 0; row < targetRows; row++) {
      for (let col = 0; col < targetCols; col++) {
        if (row < sourceRows && col < sourceCols) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
        } else {
          writeSpillArrayValue(arrayIndex, outputOffset, padTag, padValue);
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      targetRows,
      targetCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Trimrange && argc >= 1 && argc <= 3) {
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
    for (let arg = 1; arg < argc; arg++) {
      const slot = base + arg;
      if (kindStack[slot] != STACK_KIND_SCALAR) {
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
      if (tagStack[slot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[slot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    const trimRows = argc >= 2 ? coerceTrimMode(tagStack[base + 1], valueStack[base + 1]) : 3;
    const trimCols = argc >= 3 ? coerceTrimMode(tagStack[base + 2], valueStack[base + 2]) : 3;
    if (trimRows == i32.MIN_VALUE || trimCols == i32.MIN_VALUE) {
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

    let startRow = 0;
    let endRow = sourceRows - 1;
    let startCol = 0;
    let endCol = sourceCols - 1;

    const trimLeadingRows = trimRows == 1 || trimRows == 3;
    const trimTrailingRows = trimRows == 2 || trimRows == 3;
    const trimLeadingCols = trimCols == 1 || trimCols == 3;
    const trimTrailingCols = trimCols == 2 || trimCols == 3;

    if (trimLeadingRows) {
      while (startRow <= endRow) {
        let hasNonEmpty = false;
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              startRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        startRow += 1;
      }
    }

    if (trimTrailingRows) {
      while (endRow >= startRow) {
        let hasNonEmpty = false;
        for (let col = 0; col < sourceCols; col++) {
          if (
            inputCellTag(
              base,
              endRow,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        endRow -= 1;
      }
    }

    if (startRow > endRow) {
      const arrayIndex = allocateSpillArrayResult(1, 1);
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0);
      return writeArrayResult(
        base,
        arrayIndex,
        1,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (trimLeadingCols) {
      while (startCol <= endCol) {
        let hasNonEmpty = false;
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              startCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        startCol += 1;
      }
    }

    if (trimTrailingCols) {
      while (endCol >= startCol) {
        let hasNonEmpty = false;
        for (let row = startRow; row <= endRow; row++) {
          if (
            inputCellTag(
              base,
              row,
              endCol,
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
            ) != ValueTag.Empty
          ) {
            hasNonEmpty = true;
            break;
          }
        }
        if (hasNonEmpty) {
          break;
        }
        endCol -= 1;
      }
    }

    if (startCol > endCol) {
      const arrayIndex = allocateSpillArrayResult(1, 1);
      writeSpillArrayValue(arrayIndex, 0, <u8>ValueTag.Empty, 0);
      return writeArrayResult(
        base,
        arrayIndex,
        1,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const outputRows = endRow - startRow + 1;
    const outputCols = endCol - startCol + 1;
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    let outputOffset = 0;
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
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
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
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

  if (builtinId == BuiltinId.Choosecols && argc >= 2) {
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
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
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
    const selectedCols = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
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
      const selectedCol = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedCol < 0 || selectedCol >= sourceCols || selectedCol == i32.MIN_VALUE - 1) {
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
      selectedCols.push(selectedCol);
    }
    if (selectedCols.length == 0) {
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
    const outputCols = <i32>selectedCols.length;
    const arrayIndex = allocateSpillArrayResult(sourceRows, outputCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let selectedColIndex = 0; selectedColIndex < outputCols; selectedColIndex++) {
        const selectedCol = selectedCols[selectedColIndex];
        const sourceValue = inputCellNumeric(
          base,
          row,
          selectedCol,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Chooserows && argc >= 2) {
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
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
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
    const selectedRows = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
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
      const selectedRow = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedRow < 0 || selectedRow >= sourceRows || selectedRow == i32.MIN_VALUE - 1) {
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
      selectedRows.push(selectedRow);
    }
    if (selectedRows.length == 0) {
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
    const outputRows = <i32>selectedRows.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, sourceCols);
    let outputOffset = 0;
    for (let selectedRowIndex = 0; selectedRowIndex < outputRows; selectedRowIndex++) {
      const selectedRow = selectedRows[selectedRowIndex];
      for (let col = 0; col < sourceCols; col++) {
        const sourceValue = inputCellNumeric(
          base,
          selectedRow,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sort && argc >= 1 && argc <= 4) {
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

    const sortIndex = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 1;
    const sortOrder = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 1;
    const sortByColBoolean =
      argc >= 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 0;
    if (
      sortIndex == i32.MIN_VALUE ||
      sortIndex < 1 ||
      (sortOrder != 1 && sortOrder != -1) ||
      sortByColBoolean < 0
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
    const sortByCol = sortByColBoolean != 0;

    if (sourceRows == 1 || sourceCols == 1) {
      const length = sourceRows * sourceCols;
      const order = new Array<i32>(length);
      for (let index = 0; index < length; index++) {
        order.push(index);
      }
      for (let index = 1; index < length; index++) {
        const current = order[index];
        const currentRow = current / sourceCols;
        const currentCol = current - currentRow * sourceCols;
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          currentCol,
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
        if (isNaN(currentValue)) {
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
        let cursor = index;
        while (cursor > 0) {
          const previous = order[cursor - 1];
          const previousRow = previous / sourceCols;
          const previousCol = previous - previousRow * sourceCols;
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            previousCol,
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
          if (isNaN(previousValue)) {
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
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            order[cursor] = previous;
            cursor -= 1;
            continue;
          }
          break;
        }
        order[cursor] = current;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      for (let index = 0; index < length; index++) {
        const sourceOffset = order[index];
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        const sourceValue = inputCellNumeric(
          base,
          sourceRow,
          sourceCol,
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
        if (isNaN(sourceValue)) {
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
        writeSpillArrayNumber(arrayIndex, index, sourceValue);
      }
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

    if (sortByCol) {
      if (sortIndex > sourceRows) {
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
      const rowSort = new Array<i32>(sourceRows);
      for (let row = 0; row < sourceRows; row++) {
        rowSort.push(row);
      }
      const sortCol = sortIndex - 1;
      for (let cursor = 1; cursor < sourceRows; cursor++) {
        const currentRow = rowSort[cursor];
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          sortCol,
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
        if (isNaN(currentValue)) {
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
        let position = cursor;
        while (position > 0) {
          const previousRow = rowSort[position - 1];
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            sortCol,
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
          if (isNaN(previousValue)) {
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
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            rowSort[position] = previousRow;
            position -= 1;
            continue;
          }
          break;
        }
        rowSort[position] = currentRow;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      let outputOffset = 0;
      for (let sortedRow = 0; sortedRow < sourceRows; sortedRow++) {
        const sourceRow = rowSort[sortedRow];
        for (let col = 0; col < sourceCols; col++) {
          const value = inputCellNumeric(
            base,
            sourceRow,
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
          if (isNaN(value)) {
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
          writeSpillArrayNumber(arrayIndex, outputOffset, value);
          outputOffset += 1;
        }
      }
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

    if (sortIndex > sourceCols) {
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
    const colSort = new Array<i32>(sourceCols);
    for (let col = 0; col < sourceCols; col++) {
      colSort.push(col);
    }
    const sortRow = sortIndex - 1;
    for (let cursor = 1; cursor < sourceCols; cursor++) {
      const currentCol = colSort[cursor];
      const currentValue = inputCellNumeric(
        base,
        currentCol,
        sortRow,
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
      if (isNaN(currentValue)) {
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
      let position = cursor;
      while (position > 0) {
        const previousCol = colSort[position - 1];
        const previousValue = inputCellNumeric(
          base,
          previousCol,
          sortRow,
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
        if (isNaN(previousValue)) {
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
        const comparison =
          currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
        if (comparison * sortOrder < 0) {
          colSort[position] = previousCol;
          position -= 1;
          continue;
        }
        break;
      }
      colSort[position] = currentCol;
    }
    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let sortedCol = 0; sortedCol < sourceCols; sortedCol++) {
        const sourceCol = colSort[sortedCol];
        const value = inputCellNumeric(
          base,
          row,
          sourceCol,
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
        if (isNaN(value)) {
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
        writeSpillArrayNumber(arrayIndex, outputOffset, value);
        outputOffset += 1;
      }
    }
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

  if (builtinId == BuiltinId.Sortby && argc >= 2) {
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
    const sourceLength = sourceRows * sourceCols;
    if (sourceRows > 1 && sourceCols > 1) {
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

    const sortBySlots = new Array<i32>();
    const sortByLengths = new Array<i32>();
    const sortByCols = new Array<i32>();
    const sortByOrders = new Array<i32>();
    let arg = 1;
    while (arg < argc) {
      const criterionSlot = base + arg;
      const criterionKind = kindStack[criterionSlot];
      if (
        criterionKind != STACK_KIND_SCALAR &&
        criterionKind != STACK_KIND_RANGE &&
        criterionKind != STACK_KIND_ARRAY
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

      const criterionRows = inputRowsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeRowCounts,
      );
      const criterionCols = inputColsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeColCounts,
      );
      if (
        criterionRows <= 0 ||
        criterionCols <= 0 ||
        criterionRows == i32.MIN_VALUE ||
        criterionCols == i32.MIN_VALUE
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
      const criterionLength = criterionRows * criterionCols;
      if (criterionLength != 1 && criterionLength != sourceLength) {
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

      sortBySlots.push(criterionSlot);
      sortByLengths.push(criterionLength);
      sortByCols.push(criterionCols);

      const nextSlot = criterionSlot + 1;
      if (nextSlot < base + argc) {
        if (kindStack[nextSlot] == STACK_KIND_RANGE || kindStack[nextSlot] == STACK_KIND_ARRAY) {
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
        if (kindStack[nextSlot] == STACK_KIND_SCALAR) {
          const requestedOrder = coerceInteger(tagStack[nextSlot], valueStack[nextSlot]);
          if (requestedOrder != 1 && requestedOrder != -1) {
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
          sortByOrders.push(requestedOrder);
          arg += 2;
          continue;
        }
      }

      sortByOrders.push(1);
      arg += 1;
    }

    if (sortBySlots.length == 0) {
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

    const sourceIndexes = new Array<i32>(sourceLength);
    for (let offset = 0; offset < sourceLength; offset++) {
      sourceIndexes.push(offset);
    }

    for (let cursor = 1; cursor < sourceLength; cursor++) {
      const currentOffset = sourceIndexes[cursor];
      const currentSourceRow = currentOffset / sourceCols;
      const currentSourceCol = currentOffset - currentSourceRow * sourceCols;
      const currentSourceValue = inputCellNumeric(
        base,
        currentSourceRow,
        currentSourceCol,
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
      if (isNaN(currentSourceValue)) {
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

      let position = cursor;
      while (position > 0) {
        const previousOffset = sourceIndexes[position - 1];
        const previousSourceRow = previousOffset / sourceCols;
        const previousSourceCol = previousOffset - previousSourceRow * sourceCols;
        const previousSourceValue = inputCellNumeric(
          base,
          previousSourceRow,
          previousSourceCol,
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
        if (isNaN(previousSourceValue)) {
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

        let comparison = 0;
        for (let sortByIndex = 0; sortByIndex < sortBySlots.length; sortByIndex++) {
          const slot = sortBySlots[sortByIndex];
          const slotLength = sortByLengths[sortByIndex];
          const slotCols = sortByCols[sortByIndex];
          const slotOrder = sortByOrders[sortByIndex];

          const currentCriterionOffset = slotLength == 1 ? 0 : currentOffset;
          const previousCriterionOffset = slotLength == 1 ? 0 : previousOffset;
          const currentCriterionRow = slotLength == 1 ? 0 : currentCriterionOffset / slotCols;
          const currentCriterionCol =
            slotLength == 1 ? 0 : currentCriterionOffset - currentCriterionRow * slotCols;
          const previousCriterionRow = slotLength == 1 ? 0 : previousCriterionOffset / slotCols;
          const previousCriterionCol =
            slotLength == 1 ? 0 : previousCriterionOffset - previousCriterionRow * slotCols;

          const currentCriterionTag = inputCellTag(
            slot,
            currentCriterionRow,
            currentCriterionCol,
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
          const currentCriterionValue = inputCellNumeric(
            slot,
            currentCriterionRow,
            currentCriterionCol,
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
          const previousCriterionTag = inputCellTag(
            slot,
            previousCriterionRow,
            previousCriterionCol,
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
          const previousCriterionValue = inputCellNumeric(
            slot,
            previousCriterionRow,
            previousCriterionCol,
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
          const criterionComparison = compareScalarValues(
            currentCriterionTag,
            currentCriterionValue,
            previousCriterionTag,
            previousCriterionValue,
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (criterionComparison == i32.MIN_VALUE) {
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
          if (criterionComparison != 0) {
            comparison = criterionComparison * slotOrder;
            break;
          }
        }
        if (comparison < 0) {
          sourceIndexes[position] = previousOffset;
          position -= 1;
          continue;
        }
        if (comparison > 0) {
          break;
        }
        break;
      }
      sourceIndexes[position] = currentOffset;
    }

    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let index = 0; index < sourceLength; index++) {
      const sortedOffset = sourceIndexes[index];
      const sourceRow = sortedOffset / sourceCols;
      const sourceCol = sortedOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
        base,
        sourceRow,
        sourceCol,
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
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
      outputOffset += 1;
    }
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

  if (builtinId == BuiltinId.Tocol && argc >= 1 && argc <= 3) {
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
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
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
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
      for (let row = 0; row < sourceRows; row++) {
        for (let col = 0; col < sourceCols; col++) {
          const sourceTag = inputCellTag(
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
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    } else {
      for (let col = 0; col < sourceCols; col++) {
        for (let row = 0; row < sourceRows; row++) {
          const sourceTag = inputCellTag(
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
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    }

    const outputRows = values.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, 1);
    for (let offset = 0; offset < outputRows; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Torow && argc >= 1 && argc <= 3) {
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
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
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
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
      for (let row = 0; row < sourceRows; row++) {
        for (let col = 0; col < sourceCols; col++) {
          const sourceTag = inputCellTag(
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
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    } else {
      for (let col = 0; col < sourceCols; col++) {
        for (let row = 0; row < sourceRows; row++) {
          const sourceTag = inputCellTag(
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
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
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
          if (isNaN(sourceValue)) {
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
          values.push(sourceValue);
        }
      }
    }

    const outputCols = values.length;
    const arrayIndex = allocateSpillArrayResult(1, outputCols);
    for (let offset = 0; offset < outputCols; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      1,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Wraprows && argc >= 2 && argc <= 4) {
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

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
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

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
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

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
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
    }

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
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

    const sourceLength = sourceRows * sourceCols;
    const outputCols = wrapCount;
    const outputRows = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < sourceLength; outputOffset++) {
      const sourceRow = outputOffset / sourceCols;
      const sourceCol = outputOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
        base,
        sourceRow,
        sourceCol,
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
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
    }
    const padValue = needsPad && argc >= 3 ? defaultPadValue : 0;
    for (let outputOffset = sourceLength; outputOffset < outputLength; outputOffset++) {
      writeSpillArrayNumber(arrayIndex, outputOffset, padValue);
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

  if (builtinId == BuiltinId.Wrapcols && argc >= 2 && argc <= 4) {
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

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
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

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
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
    }

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
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

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
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

    const sourceLength = sourceRows * sourceCols;
    const outputRows = wrapCount;
    const outputCols = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < outputLength; outputOffset++) {
      const sourceOffset = (outputOffset / outputRows) * outputRows + (outputOffset % outputRows);
      let sourceValue = defaultPadValue;
      if (sourceOffset < sourceLength) {
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        sourceValue = inputCellNumeric(
          base,
          sourceRow,
          sourceCol,
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
      }
      if (isNaN(sourceValue)) {
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
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
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

  if (builtinId == BuiltinId.Index && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_RANGE || kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (tagStack[base + 1] == ValueTag.Error) {
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
    if (
      argc == 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const rawRowNum = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const rawColNum = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      rawRowNum == i32.MIN_VALUE ||
      rawColNum == i32.MIN_VALUE
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

    let rowNum = rawRowNum;
    let colNum = rawColNum;
    if (rowCount == 1 && rawColNum == 1) {
      rowNum = 1;
      colNum = rawRowNum;
    }
    if (rowNum < 1 || colNum < 1 || rowNum > rowCount || colNum > colCount) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const memberIndex = rangeMemberAt(
      rangeIndex,
      rowNum - 1,
      colNum - 1,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeMemberResult(
      base,
      memberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Vlookup && (argc == 3 || argc == 4)) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_SCALAR
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
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const colIndex = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const rangeLookup = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      colIndex == i32.MIN_VALUE ||
      colIndex < 1 ||
      colIndex > colCount ||
      rangeLookup < 0
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

    let matchedRow = -1;
    for (let row = 0; row < rowCount; row++) {
      const memberIndex = rangeMemberAt(
        rangeIndex,
        row,
        0,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
      );
      if (memberIndex == 0xffffffff) {
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
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
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
      if (comparison == 0) {
        matchedRow = row;
        break;
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedRow = row;
        continue;
      }
      if (rangeLookup == 1 && comparison > 0) {
        break;
      }
    }

    if (matchedRow < 0) {
      return writeResult(
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
    const resultMemberIndex = rangeMemberAt(
      rangeIndex,
      matchedRow,
      colIndex - 1,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (resultMemberIndex == 0xffffffff) {
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
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Hlookup && (argc == 3 || argc == 4)) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_SCALAR
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
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const rowIndex = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const rangeLookup = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      rowIndex == i32.MIN_VALUE ||
      rowIndex < 1 ||
      rowIndex > rowCount ||
      rangeLookup < 0
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

    let matchedCol = -1;
    for (let col = 0; col < colCount; col++) {
      const memberIndex = rangeMemberAt(
        rangeIndex,
        0,
        col,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
      );
      if (memberIndex == 0xffffffff) {
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
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
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
      if (comparison == 0) {
        matchedCol = col;
        break;
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedCol = col;
        continue;
      }
      if (rangeLookup == 1 && comparison > 0) {
        break;
      }
    }

    if (matchedCol < 0) {
      return writeResult(
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
    const resultMemberIndex = rangeMemberAt(
      rangeIndex,
      rowIndex - 1,
      matchedCol,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (resultMemberIndex == 0xffffffff) {
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
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Sum) {
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
    const rangeError = rangeErrorAt(
      base,
      argc,
      kindStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellErrors,
    );
    if (rangeError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        rangeError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let sum = 0.0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric)) {
            sum += numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Avg) {
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
    const rangeError = rangeErrorAt(
      base,
      argc,
      kindStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellErrors,
    );
    if (rangeError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        rangeError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let sum = 0.0;
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric)) {
            sum += numeric;
            count += 1;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0 : sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Min) {
    let min = Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric < min) {
            min = numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric) && numeric < min) {
            min = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric < min) {
        min = numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      min,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Max) {
    let max = -Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric > max) {
            max = numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric) && numeric > max) {
            max = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric > max) {
        max = numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      max,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (!isNaN(toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]))) {
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot]);
        continue;
      }
      if (!isNaN(toNumberOrNaN(tagStack[slot], valueStack[slot]))) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (cellTags[memberIndex] != ValueTag.Empty) {
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot]);
        continue;
      }
      if (tagStack[slot] != ValueTag.Empty) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Countblank) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor];
          if (cellTags[memberIndex] == ValueTag.Empty) {
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor += 1) {
          if (readSpillArrayTag(arrayIndex, cursor) == ValueTag.Empty) {
            count += 1;
          }
        }
        continue;
      }
      if (tagStack[slot] == ValueTag.Empty) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    builtinId == BuiltinId.Gcd ||
    builtinId == BuiltinId.Lcm ||
    builtinId == BuiltinId.Product ||
    builtinId == BuiltinId.Geomean ||
    builtinId == BuiltinId.Harmean ||
    builtinId == BuiltinId.Sumsq
  ) {
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
    const rangeError = rangeErrorAt(
      base,
      argc,
      kindStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellErrors,
    );
    if (rangeError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        rangeError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let count = 0;
    let product = 1.0;
    let sumSquares = 0.0;
    let gcdValue = 0.0;
    let lcmValue = 0.0;
    let logSum = 0.0;
    let reciprocalSum = 0.0;

    for (let index = 0; index < argc; index += 1) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (isNaN(numeric)) {
            continue;
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
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
            if (numeric == 0.0) {
              return writeResult(
                base,
                STACK_KIND_SCALAR,
                <u8>ValueTag.Number,
                0.0,
                rangeIndexStack,
                valueStack,
                tagStack,
                kindStack,
              );
            }
            logSum += Math.log(numeric);
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
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
            reciprocalSum += 1.0 / numeric;
          } else if (builtinId == BuiltinId.Gcd) {
            gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric);
          } else if (builtinId == BuiltinId.Lcm) {
            lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric);
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric;
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric;
          }
          count += 1;
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (isNaN(numeric)) {
            continue;
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
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
            if (numeric == 0.0) {
              return writeResult(
                base,
                STACK_KIND_SCALAR,
                <u8>ValueTag.Number,
                0.0,
                rangeIndexStack,
                valueStack,
                tagStack,
                kindStack,
              );
            }
            logSum += Math.log(numeric);
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
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
            reciprocalSum += 1.0 / numeric;
          } else if (builtinId == BuiltinId.Gcd) {
            gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric);
          } else if (builtinId == BuiltinId.Lcm) {
            lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric);
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric;
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric;
          }
          count += 1;
        }
        continue;
      }

      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (isNaN(numeric)) {
        continue;
      }
      if (builtinId == BuiltinId.Geomean) {
        if (numeric < 0.0) {
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
        if (numeric == 0.0) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            0.0,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        logSum += Math.log(numeric);
      } else if (builtinId == BuiltinId.Harmean) {
        if (numeric <= 0.0) {
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
        reciprocalSum += 1.0 / numeric;
      } else if (builtinId == BuiltinId.Gcd) {
        gcdValue = count == 0 ? truncAbs(numeric) : gcdPairCalc(gcdValue, numeric);
      } else if (builtinId == BuiltinId.Lcm) {
        lcmValue = count == 0 ? truncAbs(numeric) : lcmPairCalc(lcmValue, numeric);
      } else if (builtinId == BuiltinId.Product) {
        product *= numeric;
      } else if (builtinId == BuiltinId.Sumsq) {
        sumSquares += numeric * numeric;
      }
      count += 1;
    }

    if (
      (builtinId == BuiltinId.Gcd ||
        builtinId == BuiltinId.Lcm ||
        builtinId == BuiltinId.Geomean ||
        builtinId == BuiltinId.Harmean) &&
      count == 0
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

    let result = 0.0;
    if (builtinId == BuiltinId.Gcd) {
      result = gcdValue;
    } else if (builtinId == BuiltinId.Lcm) {
      result = lcmValue;
    } else if (builtinId == BuiltinId.Product) {
      result = count == 0 ? 0.0 : product;
    } else if (builtinId == BuiltinId.Geomean) {
      result = Math.exp(logSum / <f64>count);
    } else if (builtinId == BuiltinId.Harmean) {
      result = <f64>count / reciprocalSum;
    } else {
      result = sumSquares;
    }

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

  if (builtinId == BuiltinId.Countif && argc == 2) {
    if (kindStack[base] != STACK_KIND_RANGE || kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (tagStack[base + 1] == ValueTag.Error) {
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

    const rangeIndex = rangeIndexStack[base];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    let count = 0;
    for (let cursor = 0; cursor < length; cursor++) {
      const memberIndex = rangeMembers[start + cursor];
      const memberTag = cellTags[memberIndex];
      const memberValue =
        memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
      if (
        matchesCriteriaValue(
          memberTag,
          memberValue,
          tagStack[base + 1],
          valueStack[base + 1],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Daverage ||
      builtinId == BuiltinId.Dcount ||
      builtinId == BuiltinId.Dcounta ||
      builtinId == BuiltinId.Dget ||
      builtinId == BuiltinId.Dmax ||
      builtinId == BuiltinId.Dmin ||
      builtinId == BuiltinId.Dproduct ||
      builtinId == BuiltinId.Dstdev ||
      builtinId == BuiltinId.Dstdevp ||
      builtinId == BuiltinId.Dsum ||
      builtinId == BuiltinId.Dvar ||
      builtinId == BuiltinId.Dvarp) &&
    argc == 3
  ) {
    return databaseBuiltinResult(
      <u16>builtinId,
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
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
    );
  }

  if (builtinId == BuiltinId.Countifs) {
    if (argc == 0 || argc % 2 != 0) {
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

    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
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
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
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
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
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
    }

    let count = 0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 0; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
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
      if (matchesAll) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumif && (argc == 2 || argc == 3)) {
    const rangeSlot = base;
    const criteriaSlot = base + 1;
    const sumRangeSlot = argc == 3 ? base + 2 : base;
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[sumRangeSlot] != STACK_KIND_RANGE
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
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[criteriaSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[rangeSlot];
    const sumRangeIndex = rangeIndexStack[sumRangeSlot];
    const length = <i32>rangeLengths[rangeIndex];
    if (<i32>rangeLengths[sumRangeIndex] != length) {
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

    let sum = 0.0;
    for (let cursor = 0; cursor < length; cursor++) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor];
      const criteriaTag = cellTags[criteriaMemberIndex];
      const criteriaValue =
        criteriaTag == ValueTag.String
          ? <f64>cellStringIds[criteriaMemberIndex]
          : cellNumbers[criteriaMemberIndex];
      if (
        !matchesCriteriaValue(
          criteriaTag,
          criteriaValue,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue;
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + cursor];
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex]);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
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

    const sumRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[sumRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
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
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
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
    }

    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
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
      if (!matchesAll) {
        continue;
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + row];
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex]);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Averageif && (argc == 2 || argc == 3)) {
    const rangeSlot = base;
    const criteriaSlot = base + 1;
    const averageRangeSlot = argc == 3 ? base + 2 : base;
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[averageRangeSlot] != STACK_KIND_RANGE
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
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[criteriaSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[rangeSlot];
    const averageRangeIndex = rangeIndexStack[averageRangeSlot];
    const length = <i32>rangeLengths[rangeIndex];
    if (<i32>rangeLengths[averageRangeIndex] != length) {
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

    let count = 0;
    let sum = 0.0;
    for (let cursor = 0; cursor < length; cursor++) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor];
      const criteriaTag = cellTags[criteriaMemberIndex];
      const criteriaValue =
        criteriaTag == ValueTag.String
          ? <f64>cellStringIds[criteriaMemberIndex]
          : cellNumbers[criteriaMemberIndex];
      if (
        !matchesCriteriaValue(
          criteriaTag,
          criteriaValue,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue;
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + cursor];
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex]);
      if (isNaN(numeric)) {
        continue;
      }
      count += 1;
      sum += numeric;
    }
    if (count == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
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
      sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Averageifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
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

    const averageRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[averageRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
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
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
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
    }

    let count = 0;
    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
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
      if (!matchesAll) {
        continue;
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + row];
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex]);
      if (isNaN(numeric)) {
        continue;
      }
      count += 1;
      sum += numeric;
    }
    if (count == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
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
      sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Minifs || builtinId == BuiltinId.Maxifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
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

    const targetRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[targetRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
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
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
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
    }

    let found = false;
    let result = builtinId == BuiltinId.Minifs ? Infinity : -Infinity;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
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
      if (!matchesAll) {
        continue;
      }
      const targetMemberIndex = rangeMembers[rangeOffsets[targetRangeIndex] + row];
      if (cellTags[targetMemberIndex] != ValueTag.Number) {
        continue;
      }
      const numeric = cellNumbers[targetMemberIndex];
      result = builtinId == BuiltinId.Minifs ? min(result, numeric) : max(result, numeric);
      found = true;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      found ? result : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumproduct) {
    if (argc == 0) {
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

    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
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
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (
        kindStack[slot] != STACK_KIND_RANGE ||
        <i32>rangeLengths[rangeIndexStack[slot]] != expectedLength
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
    }

    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let product = 1.0;
      for (let index = 0; index < argc; index++) {
        const rangeIndex = rangeIndexStack[base + index];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        product *= toNumberOrZero(cellTags[memberIndex], cellNumbers[memberIndex]);
      }
      sum += product;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Correl ||
      builtinId == BuiltinId.Covar ||
      builtinId == BuiltinId.Pearson ||
      builtinId == BuiltinId.CovarianceP ||
      builtinId == BuiltinId.CovarianceS ||
      builtinId == BuiltinId.Intercept ||
      builtinId == BuiltinId.Rsq ||
      builtinId == BuiltinId.Slope ||
      builtinId == BuiltinId.Steyx) &&
    argc == 2
  ) {
    const statsError = collectPairedNumericStats(
      base,
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
    if (statsError != 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        <f64>statsError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const centeredSumSquaresX = pairedCenteredSumSquaresX();
    const centeredSumSquaresY = pairedCenteredSumSquaresY();
    const centeredCrossProducts = pairedCenteredCrossProducts();

    if (builtinId == BuiltinId.Covar || builtinId == BuiltinId.CovarianceP) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        centeredCrossProducts / <f64>pairedSampleCount,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (builtinId == BuiltinId.CovarianceS) {
      if (pairedSampleCount < 2) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
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
        centeredCrossProducts / <f64>(pairedSampleCount - 1),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (
      builtinId == BuiltinId.Correl ||
      builtinId == BuiltinId.Pearson ||
      builtinId == BuiltinId.Rsq
    ) {
      const denominator = Math.sqrt(centeredSumSquaresX * centeredSumSquaresY);
      if (pairedSampleCount < 2 || denominator <= 0 || !isFinite(denominator)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const correlation = centeredCrossProducts / denominator;
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        builtinId == BuiltinId.Rsq ? correlation * correlation : correlation,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (centeredSumSquaresX == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const slope = centeredCrossProducts / centeredSumSquaresX;
    if (builtinId == BuiltinId.Slope) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        slope,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const intercept =
      pairedSampleCount == 0 ? NaN : (pairedSumY - slope * pairedSumX) / <f64>pairedSampleCount;
    if (builtinId == BuiltinId.Intercept) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        intercept,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (builtinId == BuiltinId.Steyx) {
      if (pairedSampleCount <= 2) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const residualSumSquares = max<f64>(0, centeredSumSquaresY - slope * centeredCrossProducts);
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        Math.sqrt(residualSumSquares / <f64>(pairedSampleCount - 2)),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
  }

  if (builtinId == BuiltinId.Forecast && argc == 3) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
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
    const targetX = toNumberOrNaN(tagStack[base], valueStack[base]);
    if (!isFinite(targetX)) {
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

    const statsError = collectPairedNumericStats(
      base + 1,
      base + 2,
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
    if (statsError != 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        <f64>statsError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const centeredSumSquaresX = pairedCenteredSumSquaresX();
    if (centeredSumSquaresX == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const slope = pairedCenteredCrossProducts() / centeredSumSquaresX;
    const intercept =
      pairedSampleCount == 0 ? NaN : (pairedSumY - slope * pairedSumX) / <f64>pairedSampleCount;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      intercept + slope * targetX,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Trend || builtinId == BuiltinId.Growth) && argc >= 1 && argc <= 4) {
    let includeIntercept = true;
    if (argc == 4) {
      if (kindStack[base + 3] != STACK_KIND_SCALAR) {
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
      if (tagStack[base + 3] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 3],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[base + 3] == ValueTag.Boolean || tagStack[base + 3] == ValueTag.Number) {
        includeIntercept = valueStack[base + 3] != 0;
      } else if (tagStack[base + 3] == ValueTag.Empty) {
        includeIntercept = false;
      } else {
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
    }

    const knownYRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const knownYCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    const sampleCount = knownYRows * knownYCols;
    if (knownYRows < 1 || knownYCols < 1 || sampleCount < 1) {
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

    const knownYValues = new Array<f64>();
    for (let row = 0; row < knownYRows; row += 1) {
      for (let col = 0; col < knownYCols; col += 1) {
        const yTag = inputCellTag(
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
        const yRaw = inputCellScalarValue(
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
        if (yTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            yRaw,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const numeric = toNumberOrNaN(yTag, yRaw);
        if (!isFinite(numeric) || (builtinId == BuiltinId.Growth && numeric <= 0.0)) {
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
        knownYValues.push(builtinId == BuiltinId.Growth ? Math.log(numeric) : numeric);
      }
    }

    let knownXRows = knownYRows;
    let knownXCols = knownYCols;
    const knownXValues = new Array<f64>();
    if (argc >= 2) {
      knownXRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
      knownXCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
      if (knownXRows < 1 || knownXCols < 1 || knownXRows * knownXCols != sampleCount) {
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
      for (let row = 0; row < knownXRows; row += 1) {
        for (let col = 0; col < knownXCols; col += 1) {
          const numeric = inputCellNumeric(
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
          knownXValues.push(numeric);
        }
      }
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        knownXValues.push(<f64>(index + 1));
      }
    }

    let slope = 0.0;
    let intercept = 0.0;
    if (includeIntercept) {
      let sumX = 0.0;
      let sumY = 0.0;
      for (let index = 0; index < sampleCount; index += 1) {
        sumX += unchecked(knownXValues[index]);
        sumY += unchecked(knownYValues[index]);
      }
      const meanX = sumX / <f64>sampleCount;
      const meanY = sumY / <f64>sampleCount;
      let sumSquaresX = 0.0;
      let sumCrossProducts = 0.0;
      for (let index = 0; index < sampleCount; index += 1) {
        const xDeviation = unchecked(knownXValues[index]) - meanX;
        const yDeviation = unchecked(knownYValues[index]) - meanY;
        sumSquaresX += xDeviation * xDeviation;
        sumCrossProducts += xDeviation * yDeviation;
      }
      if (sumSquaresX == 0.0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      slope = sumCrossProducts / sumSquaresX;
      intercept = meanY - slope * meanX;
    } else {
      let sumSquaresX = 0.0;
      let sumCrossProducts = 0.0;
      for (let index = 0; index < sampleCount; index += 1) {
        const xValue = unchecked(knownXValues[index]);
        const yValue = unchecked(knownYValues[index]);
        sumSquaresX += xValue * xValue;
        sumCrossProducts += xValue * yValue;
      }
      if (sumSquaresX == 0.0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      slope = sumCrossProducts / sumSquaresX;
    }

    let predictionRows = knownYRows;
    let predictionCols = knownYCols;
    if (argc >= 3) {
      predictionRows = inputRowsFromSlot(base + 2, kindStack, rangeIndexStack, rangeRowCounts);
      predictionCols = inputColsFromSlot(base + 2, kindStack, rangeIndexStack, rangeColCounts);
      if (predictionRows < 1 || predictionCols < 1) {
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
    } else if (argc >= 2) {
      predictionRows = knownXRows;
      predictionCols = knownXCols;
    }

    const predictionCount = predictionRows * predictionCols;
    if (predictionCount == 1) {
      let predictionX = argc >= 2 ? unchecked(knownXValues[0]) : 1.0;
      if (argc >= 3) {
        predictionX = inputCellNumeric(
          base + 2,
          0,
          0,
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
        if (!isFinite(predictionX)) {
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
      }
      let result = intercept + slope * predictionX;
      if (builtinId == BuiltinId.Growth) {
        result = Math.exp(result);
      }
      if (!isFinite(result)) {
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
        result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(predictionRows, predictionCols);
    if (argc >= 3) {
      let cursor = 0;
      for (let row = 0; row < predictionRows; row += 1) {
        for (let col = 0; col < predictionCols; col += 1) {
          const predictionX = inputCellNumeric(
            base + 2,
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
          if (!isFinite(predictionX)) {
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
          let result = intercept + slope * predictionX;
          if (builtinId == BuiltinId.Growth) {
            result = Math.exp(result);
          }
          if (!isFinite(result)) {
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
          writeSpillArrayNumber(arrayIndex, cursor, result);
          cursor += 1;
        }
      }
    } else {
      for (let index = 0; index < predictionCount; index += 1) {
        let result = intercept + slope * unchecked(knownXValues[index]);
        if (builtinId == BuiltinId.Growth) {
          result = Math.exp(result);
        }
        if (!isFinite(result)) {
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
        writeSpillArrayNumber(arrayIndex, index, result);
      }
    }

    return writeArrayResult(
      base,
      arrayIndex,
      predictionRows,
      predictionCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Linest || builtinId == BuiltinId.Logest) && argc >= 1 && argc <= 4) {
    let includeIntercept = true;
    let includeStats = false;
    if (argc >= 3) {
      if (kindStack[base + 2] != STACK_KIND_SCALAR) {
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
      if (tagStack[base + 2] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[base + 2] == ValueTag.Boolean || tagStack[base + 2] == ValueTag.Number) {
        includeIntercept = valueStack[base + 2] != 0;
      } else if (tagStack[base + 2] == ValueTag.Empty) {
        includeIntercept = false;
      } else {
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
    }
    if (argc == 4) {
      if (kindStack[base + 3] != STACK_KIND_SCALAR) {
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
      if (tagStack[base + 3] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 3],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[base + 3] == ValueTag.Boolean || tagStack[base + 3] == ValueTag.Number) {
        includeStats = valueStack[base + 3] != 0;
      } else if (tagStack[base + 3] == ValueTag.Empty) {
        includeStats = false;
      } else {
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
    }

    const knownYRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const knownYCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    const sampleCount = knownYRows * knownYCols;
    if (sampleCount < 1) {
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

    const knownYValues = new Array<f64>();
    for (let row = 0; row < knownYRows; row += 1) {
      for (let col = 0; col < knownYCols; col += 1) {
        const yTag = inputCellTag(
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
        const yRaw = inputCellScalarValue(
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
        if (yTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            yRaw,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const numeric = toNumberOrNaN(yTag, yRaw);
        if (!isFinite(numeric) || (builtinId == BuiltinId.Logest && numeric <= 0.0)) {
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
        knownYValues.push(builtinId == BuiltinId.Logest ? Math.log(numeric) : numeric);
      }
    }

    const knownXValues = new Array<f64>();
    if (argc >= 2) {
      const knownXRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
      const knownXCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
      if (knownXRows * knownXCols != sampleCount) {
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
      for (let row = 0; row < knownXRows; row += 1) {
        for (let col = 0; col < knownXCols; col += 1) {
          const numeric = inputCellNumeric(
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
          knownXValues.push(numeric);
        }
      }
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        knownXValues.push(<f64>(index + 1));
      }
    }

    let sumX = 0.0;
    let sumY = 0.0;
    for (let index = 0; index < sampleCount; index += 1) {
      sumX += unchecked(knownXValues[index]);
      sumY += unchecked(knownYValues[index]);
    }

    let slope = 0.0;
    let intercept = 0.0;
    let totalSumSquares = 0.0;
    let sumSquaresX = 0.0;
    let sumCrossProducts = 0.0;
    if (includeIntercept) {
      const meanX = sumX / <f64>sampleCount;
      const meanY = sumY / <f64>sampleCount;
      for (let index = 0; index < sampleCount; index += 1) {
        const xDeviation = unchecked(knownXValues[index]) - meanX;
        const yDeviation = unchecked(knownYValues[index]) - meanY;
        sumSquaresX += xDeviation * xDeviation;
        sumCrossProducts += xDeviation * yDeviation;
        totalSumSquares += yDeviation * yDeviation;
      }
      if (sumSquaresX == 0.0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      slope = sumCrossProducts / sumSquaresX;
      intercept = meanY - slope * meanX;
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        const xValue = unchecked(knownXValues[index]);
        const yValue = unchecked(knownYValues[index]);
        sumSquaresX += xValue * xValue;
        sumCrossProducts += xValue * yValue;
        totalSumSquares += yValue * yValue;
      }
      if (sumSquaresX == 0.0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      slope = sumCrossProducts / sumSquaresX;
    }

    let residualSumSquares = 0.0;
    for (let index = 0; index < sampleCount; index += 1) {
      const residual =
        unchecked(knownYValues[index]) - (intercept + slope * unchecked(knownXValues[index]));
      residualSumSquares += residual * residual;
    }
    residualSumSquares = max<f64>(0.0, residualSumSquares);
    const regressionSumSquares = max<f64>(0.0, totalSumSquares - residualSumSquares);

    let leading = builtinId == BuiltinId.Logest ? Math.exp(slope) : slope;
    let trailing =
      builtinId == BuiltinId.Logest
        ? includeIntercept
          ? Math.exp(intercept)
          : 1.0
        : includeIntercept
          ? intercept
          : 0.0;
    if (!isFinite(leading) || !isFinite(trailing)) {
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

    const resultRows = includeStats ? 5 : 1;
    const arrayIndex = allocateSpillArrayResult(resultRows, 2);
    writeSpillArrayNumber(arrayIndex, 0, leading);
    writeSpillArrayNumber(arrayIndex, 1, trailing);

    if (includeStats) {
      const degreesFreedom = sampleCount - (includeIntercept ? 2 : 1);
      let slopeStandardError = NaN;
      let interceptStandardError = NaN;
      let rSquared = NaN;
      let standardErrorY = NaN;
      let fStatistic = NaN;

      if (degreesFreedom > 0) {
        const meanSquaredError = residualSumSquares / <f64>degreesFreedom;
        standardErrorY = Math.sqrt(meanSquaredError);
        if (includeIntercept) {
          const meanX = sumX / <f64>sampleCount;
          if (sumSquaresX > 0.0) {
            slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX);
            interceptStandardError = Math.sqrt(
              meanSquaredError * (1.0 / <f64>sampleCount + (meanX * meanX) / sumSquaresX),
            );
          }
        } else if (sumSquaresX > 0.0) {
          slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX);
          interceptStandardError = 0.0;
        }
        if (residualSumSquares == 0.0) {
          fStatistic = Infinity;
        } else {
          fStatistic = regressionSumSquares / (residualSumSquares / <f64>degreesFreedom);
        }
      }

      if (totalSumSquares == 0.0) {
        rSquared = residualSumSquares == 0.0 ? 1.0 : NaN;
      } else {
        rSquared = 1.0 - residualSumSquares / totalSumSquares;
      }

      if (isFinite(slopeStandardError)) {
        writeSpillArrayNumber(arrayIndex, 2, slopeStandardError);
      } else {
        writeSpillArrayValue(arrayIndex, 2, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      if (isFinite(interceptStandardError)) {
        writeSpillArrayNumber(arrayIndex, 3, interceptStandardError);
      } else {
        writeSpillArrayValue(arrayIndex, 3, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      if (isFinite(rSquared)) {
        writeSpillArrayNumber(arrayIndex, 4, rSquared);
      } else {
        writeSpillArrayValue(arrayIndex, 4, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      if (isFinite(standardErrorY)) {
        writeSpillArrayNumber(arrayIndex, 5, standardErrorY);
      } else {
        writeSpillArrayValue(arrayIndex, 5, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      if (isFinite(fStatistic)) {
        writeSpillArrayNumber(arrayIndex, 6, fStatistic);
      } else {
        writeSpillArrayValue(arrayIndex, 6, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      if (degreesFreedom > 0) {
        writeSpillArrayNumber(arrayIndex, 7, <f64>degreesFreedom);
      } else {
        writeSpillArrayValue(arrayIndex, 7, <u8>ValueTag.Error, ErrorCode.Div0);
      }
      writeSpillArrayNumber(arrayIndex, 8, regressionSumSquares);
      writeSpillArrayNumber(arrayIndex, 9, residualSumSquares);
    }

    return writeArrayResult(
      base,
      arrayIndex,
      resultRows,
      2,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Rank ||
      builtinId == BuiltinId.RankEq ||
      builtinId == BuiltinId.RankAvg) &&
    (argc == 2 || argc == 3)
  ) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
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

    const target = toNumberOrNaN(tagStack[base], valueStack[base]);
    if (!isFinite(target)) {
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

    let order = 0;
    if (argc == 3) {
      if (kindStack[base + 2] != STACK_KIND_SCALAR) {
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
      if (tagStack[base + 2] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const requestedOrder = coerceInteger(tagStack[base + 2], valueStack[base + 2]);
      if (requestedOrder == i32.MIN_VALUE || (requestedOrder != 0 && requestedOrder != 1)) {
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
      order = requestedOrder;
    }

    const arraySlot = base + 1;
    const rows = inputRowsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeRowCounts);
    const cols = inputColsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeColCounts);
    if (rows < 1 || cols < 1) {
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

    let preceding = 0;
    let ties = 0;
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const valueTag = inputCellTag(
          arraySlot,
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
        if (valueTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            inputCellScalarValue(
              arraySlot,
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
            ),
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }

        const numeric = inputCellNumeric(
          arraySlot,
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

        if (numeric == target) {
          ties += 1;
          continue;
        }
        if ((order == 0 && numeric > target) || (order == 1 && numeric < target)) {
          preceding += 1;
        }
      }
    }

    if (ties == 0) {
      return writeResult(
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

    const rank =
      builtinId == BuiltinId.RankAvg
        ? <f64>preceding + (<f64>ties + 1.0) / 2.0
        : <f64>(preceding + 1);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      rank,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Stdev ||
      builtinId == BuiltinId.StdevP ||
      builtinId == BuiltinId.StdevS ||
      builtinId == BuiltinId.Stdeva ||
      builtinId == BuiltinId.Stdevp ||
      builtinId == BuiltinId.Stdevpa ||
      builtinId == BuiltinId.Var ||
      builtinId == BuiltinId.VarP ||
      builtinId == BuiltinId.VarS ||
      builtinId == BuiltinId.Vara ||
      builtinId == BuiltinId.Varp ||
      builtinId == BuiltinId.Varpa ||
      builtinId == BuiltinId.Skew ||
      builtinId == BuiltinId.SkewP ||
      builtinId == BuiltinId.Kurt) &&
    argc >= 1
  ) {
    const values = collectStatValuesFromArgs(
      base,
      argc,
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
      builtinId == BuiltinId.Stdeva ||
        builtinId == BuiltinId.Stdevpa ||
        builtinId == BuiltinId.Vara ||
        builtinId == BuiltinId.Varpa,
    );
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        statCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let result = NaN;
    if (
      builtinId == BuiltinId.Stdev ||
      builtinId == BuiltinId.StdevS ||
      builtinId == BuiltinId.Stdeva
    ) {
      result = Math.sqrt(sampleVarianceOf(values));
    } else if (
      builtinId == BuiltinId.StdevP ||
      builtinId == BuiltinId.Stdevp ||
      builtinId == BuiltinId.Stdevpa
    ) {
      result = Math.sqrt(populationVarianceOf(values));
    } else if (
      builtinId == BuiltinId.Var ||
      builtinId == BuiltinId.VarS ||
      builtinId == BuiltinId.Vara
    ) {
      result = sampleVarianceOf(values);
    } else if (
      builtinId == BuiltinId.VarP ||
      builtinId == BuiltinId.Varp ||
      builtinId == BuiltinId.Varpa
    ) {
      result = populationVarianceOf(values);
    } else if (builtinId == BuiltinId.Skew) {
      result = skewSampleOf(values);
    } else if (builtinId == BuiltinId.SkewP) {
      result = skewPopulationOf(values);
    } else {
      result = kurtosisOf(values);
    }

    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Median && argc >= 1) {
    const values = collectNumericValuesFromArgs(
      base,
      argc,
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length == 0) {
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
    sortNumericValues(values);
    const middle = values.length >>> 1;
    const median =
      values.length % 2 == 0
        ? (unchecked(values[middle - 1]) + unchecked(values[middle])) / 2.0
        : unchecked(values[middle]);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      median,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Small || builtinId == BuiltinId.Large) && argc == 2) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const position = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    const values = collectNumericValuesFromSlot(
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (position < 1 || position > values.length) {
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
    sortNumericValues(values);
    const index = builtinId == BuiltinId.Small ? position - 1 : values.length - position;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      unchecked(values[index]),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Percentile ||
      builtinId == BuiltinId.PercentileInc ||
      builtinId == BuiltinId.PercentileExc ||
      builtinId == BuiltinId.Quartile ||
      builtinId == BuiltinId.QuartileInc ||
      builtinId == BuiltinId.QuartileExc) &&
    argc == 2
  ) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const values = collectNumericValuesFromSlot(
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length == 0) {
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
    sortNumericValues(values);

    if (
      builtinId == BuiltinId.Quartile ||
      builtinId == BuiltinId.QuartileInc ||
      builtinId == BuiltinId.QuartileExc
    ) {
      const quartile = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
      const exclusive = builtinId == BuiltinId.QuartileExc;
      if (quartile < 0 || quartile > 4 || (exclusive && (quartile <= 0 || quartile >= 4))) {
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
      if (!exclusive && quartile == 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          unchecked(values[0]),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (!exclusive && quartile == 4) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          unchecked(values[values.length - 1]),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const quartileValue = interpolateSortedPercentile(values, <f64>quartile / 4.0, exclusive);
      if (isNaN(quartileValue)) {
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
        quartileValue,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const percentile = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1]);
    const percentileValue = interpolateSortedPercentile(
      values,
      percentile,
      builtinId == BuiltinId.PercentileExc,
    );
    if (isNaN(percentileValue)) {
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
      percentileValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Percentrank ||
      builtinId == BuiltinId.PercentrankInc ||
      builtinId == BuiltinId.PercentrankExc) &&
    (argc == 2 || argc == 3)
  ) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const values = collectNumericValuesFromSlot(
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length < 2) {
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
    const target = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1]);
    const significance = argc == 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 3;
    if (!isFinite(target) || significance < 1) {
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
    sortNumericValues(values);
    const percentRank = interpolateSortedPercentRank(
      values,
      target,
      builtinId == BuiltinId.PercentrankExc,
    );
    if (isNaN(percentRank)) {
      return writeResult(
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      truncateToSignificance(percentRank, significance),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Mode || builtinId == BuiltinId.ModeSngl) && argc >= 1) {
    const values = new Array<f64>();
    for (let index = 0; index < argc; index += 1) {
      const collected = collectSampleNumbersFromSlot(
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
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          sampleCollectionErrorCode,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      for (let cursor = 0; cursor < collected.length; cursor += 1) {
        values.push(unchecked(collected[cursor]));
      }
    }
    const mode = modeSingleOf(values);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(mode) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(mode) ? ErrorCode.NA : mode,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ModeMult && argc >= 1) {
    const values = new Array<f64>();
    for (let index = 0; index < argc; index += 1) {
      const collected = collectSampleNumbersFromSlot(
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
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          sampleCollectionErrorCode,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      for (let cursor = 0; cursor < collected.length; cursor += 1) {
        values.push(unchecked(collected[cursor]));
      }
    }
    if (values.length == 0) {
      return writeResult(
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
    sortNumericValues(values);
    let maxCount = 0;
    let modeCount = 0;
    for (let start = 0; start < values.length; ) {
      let end = start + 1;
      while (end < values.length && unchecked(values[end]) == unchecked(values[start])) {
        end += 1;
      }
      const count = end - start;
      if (count > maxCount) {
        maxCount = count;
        modeCount = 1;
      } else if (count == maxCount) {
        modeCount += 1;
      }
      start = end;
    }
    if (maxCount < 2) {
      return writeResult(
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
    const arrayIndex = allocateSpillArrayResult(modeCount, 1);
    let output = 0;
    for (let start = 0; start < values.length; ) {
      let end = start + 1;
      while (end < values.length && unchecked(values[end]) == unchecked(values[start])) {
        end += 1;
      }
      if (end - start == maxCount) {
        writeSpillArrayNumber(arrayIndex, output, unchecked(values[start]));
        output += 1;
      }
      start = end;
    }
    return writeArrayResult(
      base,
      arrayIndex,
      modeCount,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Frequency && argc == 2) {
    const dataValues = collectSampleNumbersFromSlot(
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
    if (dataValues === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const binValues = collectSampleNumbersFromSlot(
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
    if (binValues === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    sortNumericValues(binValues);
    const bucketCount = binValues.length + 1;
    const arrayIndex = allocateSpillArrayResult(bucketCount, 1);
    for (let index = 0; index < bucketCount; index += 1) {
      writeSpillArrayNumber(arrayIndex, index, 0.0);
    }
    for (let dataIndex = 0; dataIndex < dataValues.length; dataIndex += 1) {
      const value = unchecked(dataValues[dataIndex]);
      let bucket = binValues.length;
      for (let binIndex = 0; binIndex < binValues.length; binIndex += 1) {
        if (value <= unchecked(binValues[binIndex])) {
          bucket = binIndex;
          break;
        }
      }
      writeSpillArrayNumber(arrayIndex, bucket, readSpillArrayNumber(arrayIndex, bucket) + 1.0);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      bucketCount,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Prob && (argc == 3 || argc == 4)) {
    if (
      kindStack[base + 2] != STACK_KIND_SCALAR ||
      (argc == 4 && kindStack[base + 3] != STACK_KIND_SCALAR)
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
    if (
      tagStack[base + 2] == ValueTag.Error ||
      (argc == 4 && tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : valueStack[base + 3],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lower = toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]);
    const upper = argc == 4 ? toNumberOrNaN(tagStack[base + 3], valueStack[base + 3]) : lower;
    if (!isFinite(lower) || !isFinite(upper) || upper < lower) {
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

    const xRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const xCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    const probabilityRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts);
    const probabilityCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts);
    if (
      xRows < 1 ||
      xCols < 1 ||
      probabilityRows < 1 ||
      probabilityCols < 1 ||
      xRows != probabilityRows ||
      xCols != probabilityCols
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

    let probabilitySum = 0.0;
    let total = 0.0;
    for (let row = 0; row < xRows; row += 1) {
      for (let col = 0; col < xCols; col += 1) {
        const xTag = inputCellTag(
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
        const xValue = inputCellScalarValue(
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
        if (xTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            xValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const probabilityTag = inputCellTag(
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
        const probabilityValue = inputCellScalarValue(
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
        if (probabilityTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            probabilityValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        if (xTag != ValueTag.Number || probabilityTag != ValueTag.Number) {
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
        if (!isFinite(probabilityValue) || probabilityValue < 0 || probabilityValue > 1) {
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
        probabilitySum += probabilityValue;
        if (xValue >= lower && xValue <= upper) {
          total += probabilityValue;
        }
      }
    }

    if (Math.abs(probabilitySum - 1.0) > 1e-9) {
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
      total,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Trimmean && argc == 2) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (tagStack[base + 1] == ValueTag.Error) {
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

    const percent = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(percent) || percent < 0.0 || percent >= 1.0) {
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

    const values = collectNumericCellRangeSeriesFromSlot(
      base,
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
      false,
    );
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        <f64>sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length == 0) {
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

    let excluded = <i32>Math.floor(<f64>values.length * percent);
    if ((excluded & 1) == 1) {
      excluded -= 1;
    }
    const retainedCount = values.length - excluded;
    if (retainedCount <= 0) {
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

    sortNumericValues(values);
    const trimEachSide = excluded >>> 1;
    let sum = 0.0;
    for (let index = trimEachSide; index < values.length - trimEachSide; index += 1) {
      sum += unchecked(values[index]);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum / <f64>retainedCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Match && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
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
    if (
      argc == 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const matchType = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (!(matchType == -1 || matchType == 0 || matchType == 1)) {
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

    const rangeIndex = rangeIndexStack[base + 1];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    let best = -1;
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index];
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (matchType == 0) {
        if (comparison == 0) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            index + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        continue;
      }
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
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
      if (matchType == 1) {
        if (comparison <= 0) {
          best = index + 1;
        } else {
          break;
        }
      } else if (comparison >= 0) {
        best = index + 1;
      } else {
        break;
      }
    }
    return best < 0
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          best,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Xmatch && argc >= 2 && argc <= 4) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
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
    if (
      argc >= 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const matchMode = argc >= 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const searchMode = argc == 4 ? truncToInt(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      !(matchMode == -1 || matchMode == 0 || matchMode == 1) ||
      !(searchMode == -1 || searchMode == 1)
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

    const rangeIndex = rangeIndexStack[base + 1];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    if (searchMode == -1) {
      if (matchMode == 0) {
        for (let index = length - 1; index >= 0; index--) {
          const memberIndex = rangeMembers[start + index];
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (comparison == 0) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Number,
              index + 1,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
        }
        return writeResult(
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

      let bestReversed = -1;
      let reversedPosition = 0;
      for (let index = length - 1; index >= 0; index--) {
        reversedPosition += 1;
        const memberIndex = rangeMembers[start + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == i32.MIN_VALUE) {
          return writeResult(
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
        if (matchMode == 1) {
          if (comparison <= 0) {
            bestReversed = reversedPosition;
          } else {
            break;
          }
        } else if (comparison >= 0) {
          bestReversed = reversedPosition;
        } else {
          break;
        }
      }
      return bestReversed < 0
        ? writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.NA,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          )
        : writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            length - bestReversed + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
    }

    if (matchMode == 0) {
      for (let index = 0; index < length; index++) {
        const memberIndex = rangeMembers[start + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            index + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
      }
      return writeResult(
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

    let best = -1;
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index];
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
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
      if (matchMode == 1) {
        if (comparison <= 0) {
          best = index + 1;
        } else {
          break;
        }
      } else if (comparison >= 0) {
        best = index + 1;
      } else {
        break;
      }
    }
    return best < 0
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          best,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Xlookup && argc >= 3 && argc <= 6) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_RANGE
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
    if (
      argc >= 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc >= 5 &&
      (kindStack[base + 4] != STACK_KIND_SCALAR || tagStack[base + 4] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 5 && tagStack[base + 4] == ValueTag.Error ? valueStack[base + 4] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 6 &&
      (kindStack[base + 5] != STACK_KIND_SCALAR || tagStack[base + 5] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 6 && tagStack[base + 5] == ValueTag.Error ? valueStack[base + 5] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lookupRangeIndex = rangeIndexStack[base + 1];
    const returnRangeIndex = rangeIndexStack[base + 2];
    const length = <i32>rangeLengths[lookupRangeIndex];
    if (<i32>rangeLengths[returnRangeIndex] != length) {
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

    const matchMode = argc >= 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const searchMode = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1;
    if (matchMode != 0 || !(searchMode == -1 || searchMode == 1)) {
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

    const lookupStart = rangeOffsets[lookupRangeIndex];
    const returnStart = rangeOffsets[returnRangeIndex];
    if (searchMode == -1) {
      for (let index = length - 1; index >= 0; index--) {
        const lookupMemberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index];
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
        }
      }
    } else {
      for (let index = 0; index < length; index++) {
        const lookupMemberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index];
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
        }
      }
    }

    if (argc >= 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[base + 3],
        valueStack[base + 3],
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
      ErrorCode.NA,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Lookup && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
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
    if (
      (kindStack[base + 1] != STACK_KIND_SCALAR && kindStack[base + 1] != STACK_KIND_RANGE) ||
      (argc == 3 &&
        kindStack[base + 2] != STACK_KIND_SCALAR &&
        kindStack[base + 2] != STACK_KIND_RANGE)
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
    if (kindStack[base + 1] == STACK_KIND_SCALAR && tagStack[base + 1] == ValueTag.Error) {
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
    if (
      argc == 3 &&
      kindStack[base + 2] == STACK_KIND_SCALAR &&
      tagStack[base + 2] == ValueTag.Error
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lookupLength = vectorSlotLength(
      base + 1,
      kindStack,
      rangeIndexStack,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
    );
    const resultSlot = argc == 3 ? base + 2 : base + 1;
    const resultLength = vectorSlotLength(
      resultSlot,
      kindStack,
      rangeIndexStack,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
    );
    if (
      lookupLength == i32.MIN_VALUE ||
      resultLength == i32.MIN_VALUE ||
      lookupLength != resultLength
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

    let position = -1;
    if (kindStack[base + 1] == STACK_KIND_SCALAR) {
      const exactComparison = compareScalarValues(
        tagStack[base + 1],
        valueStack[base + 1],
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (exactComparison == 0) {
        position = 1;
      } else if (
        position < 0 &&
        tagStack[base] == ValueTag.Number &&
        exactComparison != i32.MIN_VALUE &&
        exactComparison <= 0
      ) {
        position = 1;
      }
    } else {
      const lookupRangeIndex = rangeIndexStack[base + 1];
      const lookupStart = rangeOffsets[lookupRangeIndex];
      for (let index = 0; index < lookupLength; index++) {
        const memberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          position = index + 1;
          break;
        }
      }
      if (position < 0 && tagStack[base] == ValueTag.Number) {
        let best = -1;
        for (let index = 0; index < lookupLength; index++) {
          const memberIndex = rangeMembers[lookupStart + index];
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (comparison == i32.MIN_VALUE) {
            best = -1;
            break;
          }
          if (comparison <= 0) {
            best = index + 1;
            continue;
          }
          break;
        }
        position = best;
      }
    }

    if (position < 0) {
      return writeResult(
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

    if (kindStack[resultSlot] == STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[resultSlot],
        valueStack[resultSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const resultRangeIndex = rangeIndexStack[resultSlot];
    const resultMemberIndex = rangeMembers[rangeOffsets[resultRangeIndex] + position - 1];
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Areas && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
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
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Arraytotext && (argc == 1 || argc == 2)) {
    if (kindStack[base] != STACK_KIND_SCALAR && kindStack[base] != STACK_KIND_RANGE) {
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
    let format = 0;
    if (argc == 2) {
      if (kindStack[base + 1] != STACK_KIND_SCALAR) {
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
      format = truncToInt(tagStack[base + 1], valueStack[base + 1]);
      if (!(format == 0 || format == 1)) {
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
    }

    const strict = format == 1;
    let text = strict ? "{" : "";
    if (kindStack[base] == STACK_KIND_SCALAR) {
      const cellText = arrayToTextCell(
        tagStack[base],
        valueStack[base],
        strict,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (cellText == null) {
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
      text += cellText;
    } else {
      const rangeIndex = rangeIndexStack[base];
      const rowCount = <i32>rangeRowCounts[rangeIndex];
      const colCount = <i32>rangeColCounts[rangeIndex];
      for (let row = 0; row < rowCount; row++) {
        if (row > 0) {
          text += ";";
        }
        for (let col = 0; col < colCount; col++) {
          if (col > 0) {
            text += strict ? ", " : "\t";
          }
          const memberIndex = rangeMemberAt(
            rangeIndex,
            row,
            col,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
          );
          if (memberIndex == 0xffffffff) {
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
          const cellText = arrayToTextCell(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            strict,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (cellText == null) {
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
          text += cellText;
        }
      }
    }
    if (strict) {
      text += "}";
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Columns && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
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
      rangeColCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rows && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
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
      rangeRowCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Transpose && argc == 1) {
    if (
      kindStack[base] != STACK_KIND_SCALAR &&
      kindStack[base] != STACK_KIND_RANGE &&
      kindStack[base] != STACK_KIND_ARRAY
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
    if (sourceRows == 1 && sourceCols == 1) {
      const tag = inputCellTag(
        base,
        0,
        0,
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
      const value = inputCellScalarValue(
        base,
        0,
        0,
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
      if (tag == ValueTag.Error && isNaN(value)) {
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
        tag,
        value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(sourceCols, sourceRows);
    let outputOffset = 0;
    for (let col = 0; col < sourceCols; col++) {
      for (let row = 0; row < sourceRows; row++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
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
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceCols,
      sourceRows,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Hstack && argc >= 1) {
    let rowCount = 0;
    let totalCols = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const kind = kindStack[slot];
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
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
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
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
      rowCount = max(rowCount, rows);
      totalCols += cols;
    }
    for (let index = 0; index < argc; index++) {
      const rows = inputRowsFromSlot(base + index, kindStack, rangeIndexStack, rangeRowCounts);
      if (rows != 1 && rows != rowCount) {
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
    }

    const arrayIndex = allocateSpillArrayResult(rowCount, totalCols);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let index = 0; index < argc; index++) {
        const slot = base + index;
        const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
        const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
        const sourceRow = rows == 1 ? 0 : row;
        for (let col = 0; col < cols; col++) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            sourceRow,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      totalCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Vstack && argc >= 1) {
    let totalRows = 0;
    let colCount = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const kind = kindStack[slot];
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
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
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
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
      totalRows += rows;
      colCount = max(colCount, cols);
    }
    for (let index = 0; index < argc; index++) {
      const cols = inputColsFromSlot(base + index, kindStack, rangeIndexStack, rangeColCounts);
      if (cols != 1 && cols != colCount) {
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
    }

    const arrayIndex = allocateSpillArrayResult(totalRows, colCount);
    let outputOffset = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      for (let row = 0; row < rows; row++) {
        for (let col = 0; col < colCount; col++) {
          const sourceCol = cols == 1 ? 0 : col;
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            row,
            sourceCol,
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
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      totalRows,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Choose && argc >= 2) {
    const choice = truncToInt(tagStack[base], valueStack[base]);
    if (choice == i32.MIN_VALUE || choice < 1 || choice >= argc) {
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
    return materializeSlotResult(
      base,
      base + choice,
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
  }

  if (builtinId == BuiltinId.Textjoin && argc >= 3) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_SCALAR) {
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
    if (tagStack[base + 1] == ValueTag.Error) {
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

    const delimiter = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const ignoreEmptyNumeric = valueNumber(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (delimiter == null || !isFinite(ignoreEmptyNumeric)) {
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

    const ignoreEmpty = ignoreEmptyNumeric != 0;
    let joined = "";
    let hasJoinedValue = false;
    for (let slot = base + 2; slot < base + argc; slot++) {
      const kind = kindStack[slot];
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
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
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || rows <= 0 || cols <= 0) {
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
      for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
          const memberTag = inputCellTag(
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
          const memberValue = inputCellScalarValue(
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
          if (memberTag == ValueTag.Error) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              memberValue,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          const part = scalarText(
            memberTag,
            memberValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (part == null) {
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
          if (ignoreEmpty && part.length == 0) {
            continue;
          }
          if (hasJoinedValue) {
            joined += delimiter;
          }
          joined += part;
          hasJoinedValue = true;
        }
      }
    }
    return writeStringResult(base, joined, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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

  if (builtinId == BuiltinId.Na && argc == 0) {
    return writeResult(
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

  if (builtinId == BuiltinId.Iferror && argc == 2) {
    if (tagStack[base] == ValueTag.Error) {
      return copySlotResult(base, base + 1, rangeIndexStack, valueStack, tagStack, kindStack);
    }
    return base + 1;
  }

  if (builtinId == BuiltinId.Ifna && argc == 2) {
    if (tagStack[base] == ValueTag.Error && <i32>valueStack[base] == ErrorCode.NA) {
      return copySlotResult(base, base + 1, rangeIndexStack, valueStack, tagStack, kindStack);
    }
    return base + 1;
  }

  if (builtinId == BuiltinId.T && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Empty,
        0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
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
    if (tagStack[base] == ValueTag.Error) {
      return base + 1;
    }
    if (tagStack[base] == ValueTag.String) {
      return base + 1;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Empty,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.N && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
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
    if (tagStack[base] == ValueTag.Error) {
      return base + 1;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      toNumberOrZero(tagStack[base], valueStack[base]),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Type && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const typeCode =
      kindStack[base] == STACK_KIND_ARRAY || kindStack[base] == STACK_KIND_RANGE
        ? 64
        : tagStack[base] == ValueTag.Number || tagStack[base] == ValueTag.Empty
          ? 1
          : tagStack[base] == ValueTag.String
            ? 2
            : tagStack[base] == ValueTag.Boolean
              ? 4
              : 16;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>typeCode,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Delta && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const left = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const right =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0;
    if (isNaN(left) || isNaN(right)) {
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
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Gestep && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const numberValue = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const stepValue =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0;
    if (isNaN(numberValue) || isNaN(stepValue)) {
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
      numberValue >= stepValue ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Gauss || builtinId == BuiltinId.Phi) && argc == 1) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
      builtinId == BuiltinId.Gauss ? standardNormalCdf(numeric) - 0.5 : standardNormalPdf(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Erf && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const lower = toNumberExact(tagStack[base], valueStack[base]);
    const upper = argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 0.0;
    if (isNaN(lower) || (argc == 2 && isNaN(upper))) {
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
      argc == 2 ? erfApprox(upper) - erfApprox(lower) : erfApprox(lower),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.ErfPrecise ||
      builtinId == BuiltinId.Erfc ||
      builtinId == BuiltinId.ErfcPrecise ||
      builtinId == BuiltinId.Fisher ||
      builtinId == BuiltinId.Fisherinv ||
      builtinId == BuiltinId.Gammaln ||
      builtinId == BuiltinId.GammalnPrecise ||
      builtinId == BuiltinId.Gamma) &&
    argc == 1
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const value = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(value)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.ErfPrecise) {
      result = erfApprox(value);
    } else if (builtinId == BuiltinId.Erfc || builtinId == BuiltinId.ErfcPrecise) {
      result = 1.0 - erfApprox(value);
    } else if (builtinId == BuiltinId.Fisher) {
      result = value <= -1.0 || value >= 1.0 ? NaN : 0.5 * Math.log((1.0 + value) / (1.0 - value));
    } else if (builtinId == BuiltinId.Fisherinv) {
      const exponent = Math.exp(2.0 * value);
      result = (exponent - 1.0) / (exponent + 1.0);
    } else if (builtinId == BuiltinId.Gammaln || builtinId == BuiltinId.GammalnPrecise) {
      result = logGamma(value);
    } else if (builtinId == BuiltinId.Gamma) {
      result = gammaFunction(value);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.ConfidenceNorm ||
      builtinId == BuiltinId.Confidence ||
      builtinId == BuiltinId.ConfidenceT) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const alpha = toNumberExact(tagStack[base], valueStack[base]);
    const standardDeviation = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const size = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const useNormal = builtinId == BuiltinId.ConfidenceNorm || builtinId == BuiltinId.Confidence;
    const critical = useNormal
      ? inverseStandardNormal(1.0 - alpha / 2.0)
      : inverseStudentT(1.0 - alpha / 2.0, size - 1.0);
    const result =
      isNaN(alpha) ||
      isNaN(standardDeviation) ||
      isNaN(size) ||
      !(alpha > 0.0 && alpha < 1.0) ||
      standardDeviation <= 0.0 ||
      (useNormal ? size < 1.0 : size < 2.0) ||
      isNaN(critical)
        ? NaN
        : (critical * standardDeviation) / Math.sqrt(size);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Standardize && argc == 3) ||
    ((builtinId == BuiltinId.Normdist || builtinId == BuiltinId.NormDist) && argc == 4) ||
    ((builtinId == BuiltinId.Norminv || builtinId == BuiltinId.NormInv) && argc == 3) ||
    (builtinId == BuiltinId.Normsdist && argc == 1) ||
    (builtinId == BuiltinId.NormSDist && (argc == 1 || argc == 2)) ||
    (builtinId == BuiltinId.Normsinv && argc == 1) ||
    (builtinId == BuiltinId.NormSInv && argc == 1) ||
    ((builtinId == BuiltinId.Loginv || builtinId == BuiltinId.LognormInv) && argc == 3) ||
    ((builtinId == BuiltinId.Lognormdist || builtinId == BuiltinId.LognormDist) &&
      (argc == 3 || argc == 4))
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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

    let result = NaN;
    if (builtinId == BuiltinId.Standardize) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      result =
        isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0)
          ? NaN
          : (x - mean) / standardDeviation;
    } else if (builtinId == BuiltinId.Normdist || builtinId == BuiltinId.NormDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      result =
        isNaN(x) ||
        isNaN(mean) ||
        isNaN(standardDeviation) ||
        cumulative < 0 ||
        !(standardDeviation > 0.0)
          ? NaN
          : cumulative == 1
            ? standardNormalCdf((x - mean) / standardDeviation)
            : standardNormalPdf((x - mean) / standardDeviation) / standardDeviation;
    } else if (builtinId == BuiltinId.Norminv || builtinId == BuiltinId.NormInv) {
      const probability = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const inverse = inverseStandardNormal(probability);
      result =
        isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0) || isNaN(inverse)
          ? NaN
          : mean + standardDeviation * inverse;
    } else if (builtinId == BuiltinId.Normsdist || builtinId == BuiltinId.NormSDist) {
      const value = toNumberExact(tagStack[base], valueStack[base]);
      const cumulative =
        builtinId == BuiltinId.NormSDist && argc == 2
          ? coerceBoolean(tagStack[base + 1], valueStack[base + 1])
          : 1;
      result =
        isNaN(value) || cumulative < 0
          ? NaN
          : cumulative == 1
            ? standardNormalCdf(value)
            : standardNormalPdf(value);
    } else if (builtinId == BuiltinId.Normsinv || builtinId == BuiltinId.NormSInv) {
      result = inverseStandardNormal(toNumberExact(tagStack[base], valueStack[base]));
    } else if (builtinId == BuiltinId.Loginv || builtinId == BuiltinId.LognormInv) {
      const probability = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const inverse = inverseStandardNormal(probability);
      result =
        isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0) || isNaN(inverse)
          ? NaN
          : Math.exp(mean + standardDeviation * inverse);
    } else {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
      const z =
        isNaN(x) ||
        isNaN(mean) ||
        isNaN(standardDeviation) ||
        x <= 0.0 ||
        !(standardDeviation > 0.0)
          ? NaN
          : (Math.log(x) - mean) / standardDeviation;
      result =
        isNaN(z) || cumulative < 0
          ? NaN
          : cumulative == 1
            ? standardNormalCdf(z)
            : standardNormalPdf(z) / (x * standardDeviation);
    }

    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.GammaInv || builtinId == BuiltinId.Gammainv) && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const result = inverseGammaDistribution(probability, alpha, beta);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Expondist ||
      builtinId == BuiltinId.ExponDist ||
      builtinId == BuiltinId.Poisson ||
      builtinId == BuiltinId.PoissonDist ||
      builtinId == BuiltinId.Negbinomdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.Expondist || builtinId == BuiltinId.ExponDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const lambda = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      result =
        isNaN(x) || isNaN(lambda) || cumulative < 0 || x < 0.0 || lambda <= 0.0
          ? NaN
          : cumulative == 1
            ? 1.0 - Math.exp(-lambda * x)
            : lambda * Math.exp(-lambda * x);
    } else if (builtinId == BuiltinId.Poisson || builtinId == BuiltinId.PoissonDist) {
      const eventsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
      const events = <i32>eventsRaw;
      if (!isNaN(eventsRaw) && mean >= 0.0 && cumulative >= 0 && events >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= events; index += 1) {
            result += poissonProbability(index, mean);
          }
        } else {
          result = poissonProbability(events, mean);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw)) {
        result = negativeBinomialProbability(failures, successes, probability);
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Weibull ||
      builtinId == BuiltinId.WeibullDist ||
      builtinId == BuiltinId.Gammadist ||
      builtinId == BuiltinId.GammaDist ||
      builtinId == BuiltinId.Binomdist ||
      builtinId == BuiltinId.BinomDist ||
      builtinId == BuiltinId.NegbinomDist) &&
    argc == 4
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.Weibull || builtinId == BuiltinId.WeibullDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        if (cumulative == 1) {
          result = 1.0 - Math.exp(-Math.pow(x / beta, alpha));
        } else if (x == 0.0) {
          result = alpha == 1.0 ? 1.0 / beta : alpha < 1.0 ? Infinity : 0.0;
        } else {
          result =
            (alpha / Math.pow(beta, alpha)) *
            Math.pow(x, alpha - 1.0) *
            Math.exp(-Math.pow(x / beta, alpha));
        }
      }
    } else if (builtinId == BuiltinId.Gammadist || builtinId == BuiltinId.GammaDist) {
      const x = toNumberExact(tagStack[base], valueStack[base]);
      const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(x) &&
        !isNaN(alpha) &&
        !isNaN(beta) &&
        cumulative >= 0 &&
        x >= 0.0 &&
        alpha > 0.0 &&
        beta > 0.0
      ) {
        result =
          cumulative == 1
            ? gammaDistributionCdf(x, alpha, beta)
            : gammaDistributionDensity(x, alpha, beta);
      }
    } else if (builtinId == BuiltinId.Binomdist || builtinId == BuiltinId.BinomDist) {
      const successesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const trialsRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const successes = <i32>successesRaw;
      const trials = <i32>trialsRaw;
      if (!isNaN(successesRaw) && !isNaN(trialsRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= successes; index += 1) {
            result += binomialProbability(index, trials, probability);
          }
        } else {
          result = binomialProbability(successes, trials, probability);
        }
      }
    } else {
      const failuresRaw = toNumberExact(tagStack[base], valueStack[base]);
      const successesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const probability = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      const failures = <i32>failuresRaw;
      const successes = <i32>successesRaw;
      if (!isNaN(failuresRaw) && !isNaN(successesRaw) && cumulative >= 0) {
        if (cumulative == 1) {
          result = 0.0;
          for (let index = 0; index <= failures; index += 1) {
            result += negativeBinomialProbability(index, successes, probability);
          }
        } else {
          result = negativeBinomialProbability(failures, successes, probability);
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chidist ||
      builtinId == BuiltinId.LegacyChidist ||
      builtinId == BuiltinId.ChisqDistRt ||
      builtinId == BuiltinId.Chisqdist) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const result =
      !isNaN(x) && !isNaN(degrees) && x >= 0.0 && degrees >= 1.0
        ? regularizedUpperGamma(degrees / 2.0, x / 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const result =
      !isNaN(x) && !isNaN(degrees) && cumulative >= 0 && x >= 0.0 && degrees >= 1.0
        ? cumulative == 1
          ? chiSquareCdf(x, degrees)
          : chiSquareDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Chiinv ||
      builtinId == BuiltinId.ChisqInvRt ||
      builtinId == BuiltinId.Chisqinv ||
      builtinId == BuiltinId.LegacyChiinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(1.0 - probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.ChisqInv && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const degrees = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const valid =
      !isNaN(probability) &&
      !isNaN(degrees) &&
      probability > 0.0 &&
      probability < 1.0 &&
      degrees >= 1.0;
    let result = NaN;
    if (valid) {
      result = inverseChiSquare(probability, degrees);
    }
    const resultTag = isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error;
    const resultValue = isNumericResult(result) ? result : valid ? ErrorCode.NA : ErrorCode.Value;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      resultTag,
      resultValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaDist && argc >= 4 && argc <= 6) ||
    (builtinId == BuiltinId.Betadist && argc >= 3 && argc <= 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const modern = builtinId == BuiltinId.BetaDist;
    const cumulative = modern ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    const lowerBound = modern
      ? argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 0.0
      : argc >= 4
        ? toNumberExact(tagStack[base + 3], valueStack[base + 3])
        : 0.0;
    const upperBound = modern
      ? argc >= 6
        ? toNumberExact(tagStack[base + 5], valueStack[base + 5])
        : 1.0
      : argc >= 5
        ? toNumberExact(tagStack[base + 4], valueStack[base + 4])
        : 1.0;
    const result =
      !isNaN(x) &&
      !isNaN(alpha) &&
      !isNaN(beta) &&
      cumulative >= 0 &&
      !isNaN(lowerBound) &&
      !isNaN(upperBound)
        ? cumulative == 1 || !modern
          ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
          : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BetaInv || builtinId == BuiltinId.Betainv) &&
    argc >= 3 &&
    argc <= 5
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probability = toNumberExact(tagStack[base], valueStack[base]);
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const lowerBound = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const upperBound = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 1.0;
    const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.FDist && argc == 4) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && cumulative >= 0
        ? cumulative == 1
          ? fDistributionCdf(x, degrees1, degrees2)
          : fDistributionDensity(x, degrees1, degrees2)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FDistRt ||
      builtinId == BuiltinId.Fdist ||
      builtinId == BuiltinId.LegacyFdist) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const cdf = fDistributionCdf(x, degrees1, degrees2);
    const result =
      !isNaN(x) && !isNaN(degrees1Raw) && !isNaN(degrees2Raw) && isFinite(cdf) ? 1.0 - cdf : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.FInv ||
      builtinId == BuiltinId.FInvRt ||
      builtinId == BuiltinId.Finv ||
      builtinId == BuiltinId.LegacyFinv) &&
    argc == 3
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degrees1Raw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees2Raw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees1 = Math.floor(degrees1Raw);
    const degrees2 = Math.floor(degrees2Raw);
    const probability = builtinId == BuiltinId.FInv ? probabilityRaw : 1.0 - probabilityRaw;
    const result = inverseFDistribution(probability, degrees1, degrees2);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.TDist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const cumulative = coerceBoolean(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const result =
      !isNaN(x) && !isNaN(degreesRaw) && cumulative >= 0
        ? cumulative == 1
          ? studentTCdf(x, degrees)
          : studentTDensity(x, degrees)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.TDistRt || builtinId == BuiltinId.TDist2T) && argc == 2) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      (builtinId != BuiltinId.TDist2T || x >= 0.0) &&
      isFinite(upperTail)
        ? builtinId == BuiltinId.TDistRt
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tdist && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const tailsRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const degrees = Math.floor(degreesRaw);
    const tails = <i32>tailsRaw;
    const upperTail = 1.0 - studentTCdf(x, degrees);
    const result =
      !isNaN(x) &&
      !isNaN(degreesRaw) &&
      !isNaN(tailsRaw) &&
      tailsRaw == <f64>tails &&
      x >= 0.0 &&
      (tails == 1 || tails == 2) &&
      isFinite(upperTail)
        ? tails == 1
          ? upperTail
          : min(1.0, upperTail * 2.0)
        : NaN;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.TInv || builtinId == BuiltinId.TInv2T || builtinId == BuiltinId.Tinv) &&
    argc == 2
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    const probabilityRaw = toNumberExact(tagStack[base], valueStack[base]);
    const degreesRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const degrees = Math.floor(degreesRaw);
    const probability = builtinId == BuiltinId.TInv ? probabilityRaw : 1.0 - probabilityRaw / 2.0;
    const result = inverseStudentT(probability, degrees);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.BinomDistRange && (argc == 3 || argc == 4)) ||
    (builtinId == BuiltinId.Critbinom && argc == 3) ||
    (builtinId == BuiltinId.BinomInv && argc == 3) ||
    (builtinId == BuiltinId.Hypgeomdist && argc == 4) ||
    (builtinId == BuiltinId.HypgeomDist && argc == 5)
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
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
    let result = NaN;
    if (builtinId == BuiltinId.BinomDistRange) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const lowerRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const upperRaw =
        argc == 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : lowerRaw;
      const trials = <i32>trialsRaw;
      const lower = <i32>lowerRaw;
      const upper = <i32>upperRaw;
      if (!isNaN(trialsRaw) && !isNaN(lowerRaw) && !isNaN(upperRaw) && lower <= upper) {
        result = 0.0;
        for (let index = lower; index <= upper; index += 1) {
          result += binomialProbability(index, trials, probability);
        }
      }
    } else if (builtinId == BuiltinId.Critbinom || builtinId == BuiltinId.BinomInv) {
      const trialsRaw = toNumberExact(tagStack[base], valueStack[base]);
      const probability = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const alpha = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const trials = <i32>trialsRaw;
      if (
        !isNaN(trialsRaw) &&
        !isNaN(probability) &&
        !isNaN(alpha) &&
        trials >= 0 &&
        probability >= 0.0 &&
        probability <= 1.0 &&
        alpha > 0.0 &&
        alpha < 1.0
      ) {
        let cumulative = 0.0;
        for (let index = 0; index <= trials; index += 1) {
          cumulative += binomialProbability(index, trials, probability);
          if (cumulative >= alpha) {
            result = <f64>index;
            break;
          }
        }
        if (isNaN(result)) {
          result = <f64>trials;
        }
      }
    } else if (builtinId == BuiltinId.Hypgeomdist) {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw)
      ) {
        result = hypergeometricProbability(
          <i32>sampleSuccessesRaw,
          <i32>sampleSizeRaw,
          <i32>populationSuccessesRaw,
          <i32>populationSizeRaw,
        );
      }
    } else {
      const sampleSuccessesRaw = toNumberExact(tagStack[base], valueStack[base]);
      const sampleSizeRaw = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
      const populationSuccessesRaw = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      const populationSizeRaw = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
      const cumulative = coerceBoolean(tagStack[base + 4], valueStack[base + 4]);
      const sampleSuccesses = <i32>sampleSuccessesRaw;
      const sampleSize = <i32>sampleSizeRaw;
      const populationSuccesses = <i32>populationSuccessesRaw;
      const populationSize = <i32>populationSizeRaw;
      if (
        !isNaN(sampleSuccessesRaw) &&
        !isNaN(sampleSizeRaw) &&
        !isNaN(populationSuccessesRaw) &&
        !isNaN(populationSizeRaw) &&
        cumulative >= 0
      ) {
        if (cumulative == 1) {
          result = 0.0;
          const minimum = max<i32>(0, sampleSize - (populationSize - populationSuccesses));
          for (let index = minimum; index <= sampleSuccesses; index += 1) {
            result += hypergeometricProbability(
              index,
              sampleSize,
              populationSuccesses,
              populationSize,
            );
          }
        } else {
          result = hypergeometricProbability(
            sampleSuccesses,
            sampleSize,
            populationSuccesses,
            populationSize,
          );
        }
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    builtinId != BuiltinId.Valuetotext &&
    builtinId != BuiltinId.Address &&
    builtinId != BuiltinId.Dollar &&
    builtinId != BuiltinId.Dollarde &&
    builtinId != BuiltinId.Dollarfr
  ) {
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
  }

  if (builtinId == BuiltinId.Address && argc >= 2 && argc <= 5) {
    const row = coercePositiveIntegerArg(tagStack[base], valueStack[base], true, 1);
    const column = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], true, 1);
    if (row == i32.MIN_VALUE || column == i32.MIN_VALUE) {
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
    const absNumeric = argc >= 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 1.0;
    const refStyleNumeric =
      argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 1.0;
    if (!isFinite(absNumeric) || !isFinite(refStyleNumeric)) {
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
    const absNum = <i32>absNumeric;
    const refStyle = <i32>refStyleNumeric;
    if (absNum < 1 || absNum > 4 || (refStyle != 1 && refStyle != 2)) {
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
    let sheetPrefix = "";
    if (argc == 5) {
      if (tagStack[base + 4] == ValueTag.Empty || tagStack[base + 4] != ValueTag.String) {
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
      const sheetText = scalarText(
        tagStack[base + 4],
        valueStack[base + 4],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (sheetText == null) {
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
      sheetPrefix = `'${escapeSheetNameText(sheetText)}'!`;
    }
    const columnLabel = columnLabelText(column);
    if (columnLabel == null) {
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
    if (refStyle == 2) {
      const rowLabel = absNum == 1 || absNum == 2 ? row.toString() : `[${row}]`;
      const colLabel = absNum == 1 || absNum == 3 ? column.toString() : `[${column}]`;
      return writeStringResult(
        base,
        `${sheetPrefix}R${rowLabel}C${colLabel}`,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowLabel = absNum == 1 || absNum == 2 ? `$${row.toString()}` : row.toString();
    const colLabel = absNum == 1 || absNum == 3 ? `$${columnLabel}` : columnLabel;
    return writeStringResult(
      base,
      `${sheetPrefix}${colLabel}${rowLabel}`,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Dollar && argc >= 1 && argc <= 3) {
    const value = toNumberExact(tagStack[base], valueStack[base]);
    const decimalsNumeric =
      argc >= 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 2.0;
    let noCommasValue = 0.0;
    if (argc >= 3) {
      const numeric = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
      noCommasValue = isNaN(numeric) ? 0.0 : numeric;
    }
    if (!isFinite(value) || !isFinite(decimalsNumeric)) {
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
    const text = formatFixedText(value, <i32>decimalsNumeric, noCommasValue == 0.0);
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
    const normalizedText = text.startsWith("-") ? text.slice(1) : text;
    return writeStringResult(
      base,
      value < 0.0 ? `-$${normalizedText}` : `$${text}`,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Dollarde || builtinId == BuiltinId.Dollarfr) && argc == 2) {
    const value = toNumberExact(tagStack[base], valueStack[base]);
    const fractionNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(value) || !isFinite(fractionNumeric)) {
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
    const fraction = <i32>fractionNumeric;
    if (!isValidDollarFractionNative(fraction)) {
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
    if (builtinId == BuiltinId.Dollarde) {
      const integerPart = <i32>Math.floor(Math.abs(value));
      const fractionalNumerator = dollarFractionalNumerator(value);
      if (fractionalNumerator == i32.MIN_VALUE || fractionalNumerator >= fraction) {
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
      const sign = value < 0.0 ? -1.0 : 1.0;
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        sign * (<f64>integerPart + <f64>fractionalNumerator / <f64>fraction),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sign = value < 0.0 ? -1.0 : 1.0;
    const absolute = Math.abs(value);
    const integerPart = <i32>Math.floor(absolute);
    const fractional = absolute - <f64>integerPart;
    const width = digitCount(fraction);
    const scaledNumerator = <i32>Math.round(fractional * <f64>fraction);
    const carry = scaledNumerator / fraction;
    const numerator = scaledNumerator - carry * fraction;
    const outputValue = <f64>(integerPart + carry) + <f64>numerator / Math.pow(10.0, <f64>width);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sign * outputValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Base && (argc == 2 || argc == 3)) {
    const numberNumeric = toNumberExact(tagStack[base], valueStack[base]);
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const minLengthNumeric =
      argc == 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 0.0;
    if (!isFinite(numberNumeric) || !isFinite(radixNumeric) || !isFinite(minLengthNumeric)) {
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
    const numberValue = <i64>numberNumeric;
    const radixValue = <i32>radixNumeric;
    const minLength = <i32>minLengthNumeric;
    if (numberValue < 0 || radixValue < 2 || radixValue > 36 || minLength < 0) {
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
      toBaseText(numberValue, radixValue, minLength),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Decimal && argc == 2) {
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
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(radixNumeric)) {
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
    const radixValue = <i32>radixNumeric;
    if (radixValue < 2 || radixValue > 36) {
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
    let raw = "";
    if (tagStack[base] == ValueTag.String) {
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
      raw = trimAsciiWhitespace(text);
    } else {
      const numeric = toNumberExact(tagStack[base], valueStack[base]);
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
      raw = (<i64>numeric).toString();
    }
    if (raw.length == 0 || !isValidBaseText(raw, radixValue)) {
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
      parseBaseText(raw, radixValue),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Bin2dec ||
      builtinId == BuiltinId.Hex2dec ||
      builtinId == BuiltinId.Oct2dec) &&
    argc == 1
  ) {
    const radix = builtinId == BuiltinId.Bin2dec ? 2 : builtinId == BuiltinId.Hex2dec ? 16 : 8;
    const raw = signedRadixInputText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (raw == null) {
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
    const numeric = parseSignedRadixText(raw, radix, 10);
    if (numeric == i64.MIN_VALUE) {
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
      <f64>numeric,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Bin2hex ||
      builtinId == BuiltinId.Bin2oct ||
      builtinId == BuiltinId.Dec2bin ||
      builtinId == BuiltinId.Dec2hex ||
      builtinId == BuiltinId.Dec2oct ||
      builtinId == BuiltinId.Hex2bin ||
      builtinId == BuiltinId.Hex2oct ||
      builtinId == BuiltinId.Oct2bin ||
      builtinId == BuiltinId.Oct2hex) &&
    (argc == 1 || argc == 2)
  ) {
    let numeric: i64 = i64.MIN_VALUE;
    let radix = 10;
    let negativeWidth = 10;
    let minValue: i64 = -549755813888;
    let maxValue: i64 = 549755813887;

    if (builtinId == BuiltinId.Bin2hex || builtinId == BuiltinId.Bin2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 2, 10);
      }
      if (builtinId == BuiltinId.Bin2oct) {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      } else {
        radix = 16;
      }
    } else if (builtinId == BuiltinId.Hex2bin || builtinId == BuiltinId.Hex2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 16, 10);
      }
      if (builtinId == BuiltinId.Hex2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      }
    } else if (builtinId == BuiltinId.Oct2bin || builtinId == BuiltinId.Oct2hex) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 8, 10);
      }
      if (builtinId == BuiltinId.Oct2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else {
        radix = 16;
      }
    } else {
      const inputNumeric = toNumberExact(tagStack[base], valueStack[base]);
      if (isFinite(inputNumeric) && Math.abs(inputNumeric) <= MAX_SAFE_INTEGER_F64) {
        numeric = <i64>inputNumeric;
      }
      if (builtinId == BuiltinId.Dec2bin) {
        radix = 2;
        minValue = -512;
        maxValue = 511;
      } else if (builtinId == BuiltinId.Dec2oct) {
        radix = 8;
        minValue = -536870912;
        maxValue = 536870911;
      } else {
        radix = 16;
      }
    }

    const places = argc == 2 ? coerceNonNegativeShift(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (numeric == i64.MIN_VALUE || places == i64.MIN_VALUE) {
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
    const text = formatSignedRadixText(
      numeric,
      radix,
      <i32>places,
      negativeWidth,
      minValue,
      maxValue,
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

  if (builtinId == BuiltinId.Convert && argc == 3) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (
      !isFinite(numeric) ||
      tagStack[base + 1] != ValueTag.String ||
      tagStack[base + 2] != ValueTag.String
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
    const fromText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const toText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (fromText == null || toText == null || !resolveConvertUnit(fromText)) {
      return writeResult(
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
    const fromGroup = resolvedConvertGroup;
    const fromFactor = resolvedConvertFactor;
    const fromTemperature = resolvedConvertTemperature;
    if (!resolveConvertUnit(toText)) {
      return writeResult(
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
    const toGroup = resolvedConvertGroup;
    const toFactor = resolvedConvertFactor;
    const toTemperature = resolvedConvertTemperature;
    let result = NaN;
    if (fromGroup == CONVERT_GROUP_TEMPERATURE || toGroup == CONVERT_GROUP_TEMPERATURE) {
      if (fromGroup != CONVERT_GROUP_TEMPERATURE || toGroup != CONVERT_GROUP_TEMPERATURE) {
        return writeResult(
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
      result = convertKelvinToTemperature(
        toTemperature,
        convertTemperatureToKelvin(fromTemperature, numeric),
      );
    } else {
      if (fromGroup != toGroup) {
        return writeResult(
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
      result = (numeric * fromFactor) / toFactor;
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Euroconvert && argc >= 3 && argc <= 5) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (
      !isFinite(numeric) ||
      tagStack[base + 1] != ValueTag.String ||
      tagStack[base + 2] != ValueTag.String
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
    const sourceText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const targetText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const fullPrecisionNumeric =
      argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const triangulationNumeric =
      argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : NaN;
    if (
      sourceText == null ||
      targetText == null ||
      !isFinite(fullPrecisionNumeric) ||
      (argc == 5 && (!isFinite(triangulationNumeric) || triangulationNumeric < 3.0))
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
    const sourceRate = euroRateNative(sourceText);
    const targetRate = euroRateNative(targetText);
    const targetPrecision = euroCalculationPrecisionNative(targetText);
    if (!isFinite(sourceRate) || !isFinite(targetRate) || targetPrecision == i32.MIN_VALUE) {
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
    if (sourceText == targetText) {
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
    let euroValue = sourceText == "EUR" ? numeric : numeric / sourceRate;
    if (sourceText != "EUR" && argc == 5) {
      euroValue = roundToSignificantDigitsNative(euroValue, <i32>triangulationNumeric);
    }
    let result = targetText == "EUR" ? euroValue : euroValue * targetRate;
    if (fullPrecisionNumeric == 0.0) {
      result = roundToPlacesNative(result, targetPrecision);
    }
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

  if (
    (builtinId == BuiltinId.Bitand ||
      builtinId == BuiltinId.Bitor ||
      builtinId == BuiltinId.Bitxor) &&
    argc >= 2
  ) {
    let accumulatorValue = coerceBitwiseUnsigned(tagStack[base], valueStack[base]);
    if (accumulatorValue == i64.MIN_VALUE) {
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
    let accumulator = <u32>accumulatorValue;
    for (let index = 1; index < argc; index += 1) {
      const currentValue = coerceBitwiseUnsigned(tagStack[base + index], valueStack[base + index]);
      if (currentValue == i64.MIN_VALUE) {
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
      const current = <u32>currentValue;
      if (builtinId == BuiltinId.Bitand) {
        accumulator &= current;
      } else if (builtinId == BuiltinId.Bitor) {
        accumulator |= current;
      } else {
        accumulator ^= current;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>accumulator,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Bitlshift || builtinId == BuiltinId.Bitrshift) && argc == 2) {
    const value = coerceBitwiseUnsigned(tagStack[base], valueStack[base]);
    const shift = coerceNonNegativeShift(tagStack[base + 1], valueStack[base + 1]);
    if (value == i64.MIN_VALUE || shift == i64.MIN_VALUE) {
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
    const shiftAmount = <i32>(shift & 31);
    const numeric = <u32>value;
    const result =
      builtinId == BuiltinId.Bitlshift
        ? <u32>(numeric << shiftAmount)
        : <u32>(numeric >>> shiftAmount);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Besseli ||
      builtinId == BuiltinId.Besselj ||
      builtinId == BuiltinId.Besselk ||
      builtinId == BuiltinId.Bessely) &&
    argc == 2
  ) {
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const orderNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(x) || !isFinite(orderNumeric)) {
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
    const order = <i32>orderNumeric;
    if (order < 0) {
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
    if ((builtinId == BuiltinId.Besselk || builtinId == BuiltinId.Bessely) && x <= 0.0) {
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
    let result = NaN;
    if (builtinId == BuiltinId.Besseli) {
      result = besselIValue(x, order);
    } else if (builtinId == BuiltinId.Besselj) {
      result = besselJValue(x, order);
    } else if (builtinId == BuiltinId.Besselk) {
      result = besselKValue(x, order);
    } else {
      result = besselYValue(x, order);
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Abs && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.abs(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
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
      roundToDigits(numeric, 0),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Round && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(digits)) {
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
      roundToDigits(numeric, <i32>digits),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
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
      Math.floor(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Floor && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significance = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(significance)) {
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
    if (significance == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
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
      Math.floor(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
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
      Math.ceil(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significance = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(significance)) {
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
    if (significance == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
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
      Math.ceil(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.FloorMath && (argc == 1 || argc == 2 || argc == 3)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significanceRaw =
      argc >= 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 1.0;
    const significance = isNaN(significanceRaw) ? 1.0 : Math.abs(significanceRaw);
    const modeRaw = argc == 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 0.0;
    const mode = isNaN(modeRaw) ? 0.0 : modeRaw;
    if (isNaN(numeric) || significance == 0.0) {
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
    const result =
      numeric >= 0.0
        ? Math.floor(numeric / significance) * significance
        : -(mode == 0.0
            ? Math.ceil(Math.abs(numeric) / significance)
            : Math.floor(Math.abs(numeric) / significance)) * significance;
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

  if (builtinId == BuiltinId.FloorPrecise && (argc == 1 || argc == 2)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significanceRaw =
      argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 1.0;
    const significance = isNaN(significanceRaw) ? 1.0 : Math.abs(significanceRaw);
    if (isNaN(numeric) || significance == 0.0) {
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
      Math.floor(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.CeilingMath && (argc == 1 || argc == 2 || argc == 3)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significanceRaw =
      argc >= 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 1.0;
    const significance = isNaN(significanceRaw) ? 1.0 : Math.abs(significanceRaw);
    const modeRaw = argc == 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 0.0;
    const mode = isNaN(modeRaw) ? 0.0 : modeRaw;
    if (isNaN(numeric) || significance == 0.0) {
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
    const result =
      numeric >= 0.0
        ? Math.ceil(numeric / significance) * significance
        : -(mode == 0.0
            ? Math.floor(Math.abs(numeric) / significance)
            : Math.ceil(Math.abs(numeric) / significance)) * significance;
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

  if (
    (builtinId == BuiltinId.CeilingPrecise || builtinId == BuiltinId.IsoCeiling) &&
    (argc == 1 || argc == 2)
  ) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significanceRaw =
      argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 1.0;
    const significance = isNaN(significanceRaw) ? 1.0 : Math.abs(significanceRaw);
    if (isNaN(numeric) || significance == 0.0) {
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
      Math.ceil(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = toNumberOrZero(tagStack[base + 1], valueStack[base + 1]);
    if (divisor == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
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
      toNumberOrZero(tagStack[base], valueStack[base]) % divisor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.And) {
    if (argc == 0) {
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
    for (let index = 0; index < argc; index++) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (coerced == 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Boolean,
          0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Or) {
    if (argc == 0) {
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
    for (let index = 0; index < argc; index++) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (coerced != 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Boolean,
          1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Xor) {
    if (argc == 0) {
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
    let parity = 0;
    for (let index = 0; index < argc; index++) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      parity = parity ^ (coerced != 0 ? 1 : 0);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      parity != 0 ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    const coerced = coerceLogical(tagStack[base], valueStack[base]);
    if (coerced < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -coerced - 1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      coerced == 0 ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ifs) {
    if (argc < 2 || argc % 2 != 0) {
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
    for (let index = 0; index < argc; index += 2) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (coerced != 0) {
        return copySlotResult(
          base,
          base + index + 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    return writeResult(
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

  if (builtinId == BuiltinId.Switch) {
    if (argc < 3) {
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
    const hasDefault = (argc - 1) % 2 == 1;
    const pairLimit = hasDefault ? argc - 1 : argc;
    for (let index = 1; index < pairLimit; index += 2) {
      if (tagStack[base + index] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + index],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const comparison = compareScalarValues(
        tagStack[base],
        valueStack[base],
        tagStack[base + index],
        valueStack[base + index],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == 0) {
        return copySlotResult(
          base,
          base + index + 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    if (hasDefault) {
      return copySlotResult(
        base,
        base + argc - 1,
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
      ErrorCode.NA,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Concat) {
    let scalarError = -1;
    for (let index = 0; index < argc; index++) {
      if (tagStack[base + index] == ValueTag.Error) {
        scalarError = <i32>valueStack[base + index];
        break;
      }
    }
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

    for (let index = 0; index < argc; index++) {
      const len = textLength(
        tagStack[base + index],
        valueStack[base + index],
        stringLengths,
        outputStringLengths,
      );
      if (len < 0) {
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
    }

    let text = "";
    for (let index = 0; index < argc; index++) {
      const part = scalarText(
        tagStack[base + index],
        valueStack[base + index],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (part != null) {
        text += part;
      }
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Len && argc == 1) {
    const length = textLength(tagStack[base], valueStack[base], stringLengths, outputStringLengths);
    if (length < 0) {
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
      <f64>length,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Lenb && argc == 1) {
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>utf8ByteLength(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Exact && argc == 2) {
    const left = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const right = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (left === null || right === null) {
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
      <u8>ValueTag.Boolean,
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Left || builtinId == BuiltinId.Right) && (argc == 1 || argc == 2)) {
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
    const count = argc == 2 ? coerceLength(tagStack[base + 1], valueStack[base + 1], 1) : 1;
    if (text == null || count == i32.MIN_VALUE) {
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
    const result =
      builtinId == BuiltinId.Left
        ? text.slice(0, count)
        : count == 0
          ? ""
          : count >= text.length
            ? text
            : text.slice(text.length - count);
    return writeStringResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if ((builtinId == BuiltinId.Leftb || builtinId == BuiltinId.Rightb) && (argc == 1 || argc == 2)) {
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
    const count = argc == 2 ? coerceLength(tagStack[base + 1], valueStack[base + 1], 1) : 1;
    if (text == null || count == i32.MIN_VALUE) {
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
    const result =
      builtinId == BuiltinId.Leftb ? leftBytesText(text, count) : rightBytesText(text, count);
    return writeStringResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Mid && argc == 3) {
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
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE) {
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
      text.slice(start - 1, start - 1 + count),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Midb && argc == 3) {
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
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE) {
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
      midBytesText(text, start, count),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Trim && argc == 1) {
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
      excelTrim(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Upper || builtinId == BuiltinId.Lower) && argc == 1) {
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
      builtinId == BuiltinId.Upper ? text.toUpperCase() : text.toLowerCase(),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Find && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1;
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
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
    const found = findPosition(needle, haystack, start, true, false);
    if (found == i32.MIN_VALUE) {
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
      <f64>found,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Findb || builtinId == BuiltinId.Searchb) &&
    (argc == 2 || argc == 3)
  ) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (needle == null || haystack == null) {
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
    let start = 1;
    if (argc == 3) {
      const startByte = coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1);
      if (startByte == i32.MIN_VALUE || startByte > utf8ByteLength(haystack) + 1) {
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
      start = bytePositionToCharPositionUtf8(haystack, startByte);
    }
    const found = findPosition(
      needle,
      haystack,
      start,
      builtinId == BuiltinId.Findb,
      builtinId == BuiltinId.Searchb,
    );
    if (found == i32.MIN_VALUE) {
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
      <f64>charPositionToBytePositionUtf8(haystack, found),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Search && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1;
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
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
    const found = findPosition(needle, haystack, start, false, true);
    if (found == i32.MIN_VALUE) {
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
      <f64>found,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

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

  if (builtinId == BuiltinId.Replaceb && argc == 4) {
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
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    const replacement = scalarText(
      tagStack[base + 3],
      valueStack[base + 3],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE || replacement == null) {
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
      replaceBytesText(text, start, count, replacement),
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
        if (isFormatPlaceholderCode(cleaned.charCodeAt(index))) {
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

  if (builtinId == BuiltinId.IsBlank && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsBlank && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.Empty ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsNumber && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsNumber && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.Number ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsText && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsText && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.String ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Date && argc == 3) {
    const serial = excelDateSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      tagStack[base + 2],
      valueStack[base + 2],
    );
    if (isNaN(serial)) {
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
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Year && argc == 1) {
    const year = excelYearPartFromSerial(tagStack[base], valueStack[base]);
    if (year == i32.MIN_VALUE) {
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
      <f64>year,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Month && argc == 1) {
    const month = excelMonthPartFromSerial(tagStack[base], valueStack[base]);
    if (month == i32.MIN_VALUE) {
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
      <f64>month,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Day && argc == 1) {
    const day = excelDayPartFromSerial(tagStack[base], valueStack[base]);
    if (day == i32.MIN_VALUE) {
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
      <f64>day,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Time && argc == 3) {
    const serial = excelTimeSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      tagStack[base + 2],
      valueStack[base + 2],
    );
    if (isNaN(serial)) {
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
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Timevalue && argc == 1) {
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
    const serial = text == null ? NaN : parseTimeValueText(text);
    if (isNaN(serial)) {
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
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Hour && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
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
      <f64>(second / 3600),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Minute && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
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
      <f64>((second % 3600) / 60),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Second && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
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
      <f64>(second % 60),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Weekday && (argc == 1 || argc == 2)) {
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1;
    const weekday =
      returnType == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : excelWeekdayFromSerial(tagStack[base], valueStack[base], returnType);
    if (weekday == i32.MIN_VALUE) {
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
      <f64>weekday,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Isoweeknum && argc == 1) {
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
    const whole = excelSerialWhole(tagStack[base], valueStack[base]);
    const week = whole == i32.MIN_VALUE ? i32.MIN_VALUE : excelIsoWeeknumValue(whole);
    if (week == i32.MIN_VALUE) {
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
      <f64>week,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Edate && argc == 2) {
    const serial = addMonthsExcelSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      false,
    );
    if (isNaN(serial)) {
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
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Eomonth && argc == 2) {
    const serial = addMonthsExcelSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      true,
    );
    if (isNaN(serial)) {
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
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Datedif && argc == 3) {
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
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base]);
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const unitText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || unitText == null) {
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
    const unit = trimAsciiWhitespace(unitText).toUpperCase();
    const value = unit.length == 0 ? NaN : excelDatedifValue(startWhole, endWhole, unit);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Int && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
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
      Math.floor(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.RoundUp || builtinId == BuiltinId.RoundDown) &&
    (argc == 1 || argc == 2)
  ) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (isNaN(numeric) || digits == i32.MIN_VALUE) {
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
    let result = 0.0;
    if (digits >= 0) {
      const factor = Math.pow(10.0, <f64>digits);
      const scaled = numeric * factor;
      result =
        (builtinId == BuiltinId.RoundUp
          ? numeric >= 0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) / factor;
    } else {
      const factor = Math.pow(10.0, <f64>-digits);
      const scaled = numeric / factor;
      result =
        (builtinId == BuiltinId.RoundUp
          ? numeric >= 0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) * factor;
    }
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

  if (builtinId == BuiltinId.Trunc && (argc == 1 || argc == 2)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (isNaN(numeric) || digits == i32.MIN_VALUE) {
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
      roundTowardZeroDigits(numeric, digits),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sin && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.sin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Cos && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.cos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Tan && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.tan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Asin && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.asin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Acos && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.acos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.atan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan2 && argc == 2) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.atan2(
        toNumberOrZero(tagStack[base], valueStack[base]),
        toNumberOrZero(tagStack[base + 1], valueStack[base + 1]),
      ),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Degrees && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (toNumberOrZero(tagStack[base], valueStack[base]) * 180.0) / Math.PI,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Radians && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (toNumberOrZero(tagStack[base], valueStack[base]) * Math.PI) / 180.0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Exp && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.exp(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Ln && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Log10 && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log10(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Log && (argc == 1 || argc == 2)) {
    const num = toNumberOrZero(tagStack[base], valueStack[base]);
    const baseVal = argc == 2 ? toNumberOrZero(tagStack[base + 1], valueStack[base + 1]) : 10.0;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log(num) / Math.log(baseVal),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Power && argc == 2) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.pow(
        toNumberOrZero(tagStack[base], valueStack[base]),
        toNumberOrZero(tagStack[base + 1], valueStack[base + 1]),
      ),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Sqrt && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.sqrt(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Seriessum && argc >= 3) {
    const x = toNumberExact(tagStack[base], valueStack[base]);
    const n = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const m = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    if (isNaN(x) || n == i32.MIN_VALUE || m == i32.MIN_VALUE) {
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
    let sum = 0.0;
    for (let index = 0; index < argc - 3; index += 1) {
      const coefficient = toNumberOrZero(tagStack[base + 3 + index], valueStack[base + 3 + index]);
      sum += coefficient * Math.pow(x, <f64>(n + index * m));
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Sqrtpi && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const result = isNaN(numeric) ? NaN : Math.sqrt(numeric * Math.PI);
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Pi && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.PI,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Sinh ||
      builtinId == BuiltinId.Cosh ||
      builtinId == BuiltinId.Tanh ||
      builtinId == BuiltinId.Asinh ||
      builtinId == BuiltinId.Acosh ||
      builtinId == BuiltinId.Atanh ||
      builtinId == BuiltinId.Acot ||
      builtinId == BuiltinId.Acoth ||
      builtinId == BuiltinId.Cot ||
      builtinId == BuiltinId.Coth ||
      builtinId == BuiltinId.Csc ||
      builtinId == BuiltinId.Csch ||
      builtinId == BuiltinId.Sec ||
      builtinId == BuiltinId.Sech ||
      builtinId == BuiltinId.Sign ||
      builtinId == BuiltinId.Even ||
      builtinId == BuiltinId.Odd ||
      builtinId == BuiltinId.Fact ||
      builtinId == BuiltinId.Factdouble) &&
    argc == 1
  ) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
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

    let result = 0.0;
    let errorCode = ErrorCode.None;
    if (builtinId == BuiltinId.Sinh) {
      result = Math.sinh(numeric);
    } else if (builtinId == BuiltinId.Cosh) {
      result = Math.cosh(numeric);
    } else if (builtinId == BuiltinId.Tanh) {
      result = Math.tanh(numeric);
    } else if (builtinId == BuiltinId.Asinh) {
      result = Math.asinh(numeric);
    } else if (builtinId == BuiltinId.Acosh) {
      result = Math.acosh(numeric);
    } else if (builtinId == BuiltinId.Atanh) {
      result = Math.atanh(numeric);
    } else if (builtinId == BuiltinId.Acot) {
      result = numeric == 0.0 ? Math.PI / 2.0 : Math.atan(1.0 / numeric);
    } else if (builtinId == BuiltinId.Acoth) {
      result = 0.5 * Math.log((numeric + 1.0) / (numeric - 1.0));
    } else if (builtinId == BuiltinId.Cot) {
      const tangent = Math.tan(numeric);
      if (tangent == 0.0) {
        errorCode = ErrorCode.Div0;
      } else {
        result = 1.0 / tangent;
      }
    } else if (builtinId == BuiltinId.Coth) {
      const hyperbolic = Math.tanh(numeric);
      if (hyperbolic == 0.0) {
        errorCode = ErrorCode.Div0;
      } else {
        result = 1.0 / hyperbolic;
      }
    } else if (builtinId == BuiltinId.Csc) {
      const sine = Math.sin(numeric);
      if (sine == 0.0) {
        errorCode = ErrorCode.Div0;
      } else {
        result = 1.0 / sine;
      }
    } else if (builtinId == BuiltinId.Csch) {
      const hyperbolic = Math.sinh(numeric);
      if (hyperbolic == 0.0) {
        errorCode = ErrorCode.Div0;
      } else {
        result = 1.0 / hyperbolic;
      }
    } else if (builtinId == BuiltinId.Sec) {
      const cosine = Math.cos(numeric);
      if (cosine == 0.0) {
        errorCode = ErrorCode.Div0;
      } else {
        result = 1.0 / cosine;
      }
    } else if (builtinId == BuiltinId.Sech) {
      result = 1.0 / Math.cosh(numeric);
    } else if (builtinId == BuiltinId.Sign) {
      result = numeric == 0.0 ? 0.0 : numeric > 0.0 ? 1.0 : -1.0;
    } else if (builtinId == BuiltinId.Even) {
      result = evenCalc(numeric);
    } else if (builtinId == BuiltinId.Odd) {
      result = oddCalc(numeric);
    } else if (builtinId == BuiltinId.Fact) {
      result = factorialCalc(numeric);
    } else {
      result = doubleFactorialCalc(numeric);
    }

    if (errorCode != ErrorCode.None) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        errorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Combin ||
      builtinId == BuiltinId.Combina ||
      builtinId == BuiltinId.Quotient) &&
    argc == 2
  ) {
    const left = toNumberExact(tagStack[base], valueStack[base]);
    const right = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (!isFinite(left) || !isFinite(right)) {
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

    if (builtinId == BuiltinId.Quotient) {
      if (right == 0.0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Div0,
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
        Math.trunc(left / right),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const numberValue = left < 0.0 ? NaN : Math.floor(left);
    const chosenValue = right < 0.0 ? NaN : Math.floor(right);
    if (!isFinite(numberValue) || !isFinite(chosenValue)) {
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

    if (builtinId == BuiltinId.Combin && chosenValue > numberValue) {
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

    let result = 0.0;
    if (builtinId == BuiltinId.Combin) {
      result =
        factorialCalc(numberValue) /
        (factorialCalc(chosenValue) * factorialCalc(numberValue - chosenValue));
    } else {
      if (chosenValue == 0.0) {
        result = 1.0;
      } else if (numberValue == 0.0) {
        result = 0.0;
      } else {
        const combined = numberValue + chosenValue - 1.0;
        result =
          factorialCalc(combined) / (factorialCalc(chosenValue) * factorialCalc(numberValue - 1.0));
      }
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Days && argc == 2) {
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
    const end = excelSerialWhole(tagStack[base], valueStack[base]);
    const start = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (end == i32.MIN_VALUE || start == i32.MIN_VALUE) {
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
      <f64>(end - start),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Permut || builtinId == BuiltinId.Permutationa) && argc == 2) {
    const left = toNumberExact(tagStack[base], valueStack[base]);
    const right = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const numberValue = !isFinite(left) || left < 0.0 ? NaN : Math.floor(left);
    const chosenValue = !isFinite(right) || right < 0.0 ? NaN : Math.floor(right);
    if (!isFinite(numberValue) || !isFinite(chosenValue)) {
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
    if (builtinId == BuiltinId.Permut && chosenValue > numberValue) {
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
    let result = 0.0;
    if (builtinId == BuiltinId.Permut) {
      result = 1.0;
      for (let index = 0; index < <i32>chosenValue; index += 1) {
        result *= numberValue - <f64>index;
      }
    } else {
      result = Math.pow(numberValue, chosenValue);
    }
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Mround && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const multiple = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(multiple) || multiple == 0.0) {
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
    if (numeric != 0.0 && Math.sign(numeric) != Math.sign(multiple)) {
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
      Math.round(numeric / multiple) * multiple,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Days360 && (argc == 2 || argc == 3)) {
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
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base]);
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const method = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const value =
      startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || (method != 0 && method != 1)
        ? NaN
        : excelDays360Value(startWhole, endWhole, method);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Yearfrac && (argc == 2 || argc == 3)) {
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
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base]);
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const basis = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const value =
      startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || basis < 0 || basis > 4
        ? NaN
        : excelYearfracValue(startWhole, endWhole, basis);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Weeknum && (argc == 1 || argc == 2)) {
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
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1;
    const weeknum =
      returnType == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : excelWeeknumFromSerial(tagStack[base], valueStack[base], returnType);
    if (weeknum == i32.MIN_VALUE) {
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
      <f64>weeknum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Workday && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
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
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE) {
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

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    let cursor = start;
    const direction = offset >= 0 ? 1 : -1;
    while (true) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        break;
      }
      cursor += direction;
    }

    let remaining = offset >= 0 ? offset : -offset;
    while (remaining > 0) {
      cursor += direction;
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        remaining -= 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>cursor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Networkdays && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
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
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE) {
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

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        count += step;
      }
      if (cursor == end) {
        break;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.WorkdayIntl && argc >= 2 && argc <= 4) {
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
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
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

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR;
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty;
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0;
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0;

    let cursor = start;
    const direction = offset >= 0 ? 1 : -1;
    while (true) {
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        break;
      }
      cursor += direction;
    }

    let remaining = offset >= 0 ? offset : -offset;
    while (remaining > 0) {
      cursor += direction;
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        remaining -= 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>cursor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.NetworkdaysIntl && argc >= 2 && argc <= 4) {
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
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
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

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR;
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty;
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0;
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0;

    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
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
      if (workday == 1) {
        count += step;
      }
      if (cursor == end) {
        break;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Pv && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const payment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const present =
      paymentTypeValue < 0
        ? NaN
        : presentValueCalc(rate, periods, payment, future, paymentTypeValue);
    if (isNaN(present)) {
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
      present,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Pmt && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const payment =
      paymentTypeValue < 0
        ? NaN
        : periodicPaymentCalc(rate, periods, present, future, paymentTypeValue);
    if (isNaN(payment)) {
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
      payment,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Nper && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const payment = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const periods =
      paymentTypeValue < 0
        ? NaN
        : totalPeriodsCalc(rate, payment, present, future, paymentTypeValue);
    if (isNaN(periods)) {
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
      periods,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Npv && argc >= 2) {
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(rate)) {
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
    const values = collectNumericValuesFromArgs(
      base + 1,
      argc - 1,
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (values.length == 0) {
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
    let total = 0.0;
    for (let index = 0; index < values.length; index += 1) {
      total += unchecked(values[index]) / Math.pow(1.0 + rate, <f64>(index + 1));
    }
    if (!isFinite(total)) {
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
      total,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rate && argc >= 3 && argc <= 6) {
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
    const periods = toNumberExact(tagStack[base], valueStack[base]);
    const payment = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const future = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const guess = argc >= 6 ? toNumberExact(tagStack[base + 5], valueStack[base + 5]) : 0.1;
    const rate =
      paymentTypeValue < 0
        ? NaN
        : solveRateCalc(periods, payment, present, future, paymentTypeValue, guess);
    if (isNaN(rate)) {
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
      rate,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Ipmt || builtinId == BuiltinId.Ppmt) && argc >= 4 && argc <= 6) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const period = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const future = argc >= 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 0.0;
    const paymentTypeValue =
      argc >= 6 ? paymentType(tagStack[base + 5], valueStack[base + 5], true) : 0;
    const result =
      paymentTypeValue < 0
        ? NaN
        : builtinId == BuiltinId.Ipmt
          ? interestPaymentCalc(rate, period, periods, present, future, paymentTypeValue)
          : principalPaymentCalc(rate, period, periods, present, future, paymentTypeValue);
    if (isNaN(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ispmt && argc == 4) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const period = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    if (
      isNaN(rate) ||
      isNaN(period) ||
      isNaN(periods) ||
      isNaN(present) ||
      periods <= 0.0 ||
      period < 1.0 ||
      period > periods
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
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      present * rate * (period / periods - 1.0),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Cumipmt || builtinId == BuiltinId.Cumprinc) && argc == 6) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const present = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const startPeriod = truncToInt(tagStack[base + 3], valueStack[base + 3]);
    const endPeriod = truncToInt(tagStack[base + 4], valueStack[base + 4]);
    const paymentTypeValue = paymentType(tagStack[base + 5], valueStack[base + 5], true);
    const total =
      startPeriod == i32.MIN_VALUE || endPeriod == i32.MIN_VALUE || paymentTypeValue < 0
        ? NaN
        : cumulativePeriodicPaymentCalc(
            rate,
            periods,
            present,
            startPeriod,
            endPeriod,
            paymentTypeValue,
            builtinId == BuiltinId.Cumprinc,
          );
    if (isNaN(total)) {
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
      total,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Fv && argc >= 3 && argc <= 5) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periods = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const payment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const present = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0;
    const paymentTypeValue =
      argc >= 5 ? paymentType(tagStack[base + 4], valueStack[base + 4], true) : 0;
    const future =
      paymentTypeValue < 0
        ? NaN
        : futureValueCalc(rate, periods, payment, present, paymentTypeValue);
    if (isNaN(future)) {
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
      future,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Fvschedule && argc >= 2) {
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
    const principal = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(principal)) {
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
    let result = principal;
    for (let index = 1; index < argc; index++) {
      const rate = toNumberExact(tagStack[base + index], valueStack[base + index]);
      if (isNaN(rate)) {
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
      result *= 1.0 + rate;
    }
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

  if ((builtinId == BuiltinId.Effect || builtinId == BuiltinId.Nominal) && argc == 2) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const periodsNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const periods = Math.trunc(periodsNumeric);
    if (
      isNaN(rate) ||
      isNaN(periodsNumeric) ||
      !isFinite(periods) ||
      periods < 1.0 ||
      (builtinId == BuiltinId.Nominal && rate <= -1.0)
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
    const result =
      builtinId == BuiltinId.Effect
        ? Math.pow(1.0 + rate / periods, periods) - 1.0
        : periods * (Math.pow(1.0 + rate, 1.0 / periods) - 1.0);
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Pduration && argc == 3) {
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
    const rate = toNumberExact(tagStack[base], valueStack[base]);
    const present = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const future = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const result =
      isNaN(rate) ||
      isNaN(present) ||
      isNaN(future) ||
      rate <= 0.0 ||
      present <= 0.0 ||
      future <= 0.0
        ? NaN
        : Math.log(future / present) / Math.log(1.0 + rate);
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rri && argc == 3) {
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
    const periods = toNumberExact(tagStack[base], valueStack[base]);
    const present = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const future = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const result =
      isNaN(periods) || isNaN(present) || isNaN(future) || periods <= 0.0 || present == 0.0
        ? NaN
        : Math.pow(future / present, 1.0 / periods) - 1.0;
    if (!isFinite(result)) {
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
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Db && (argc == 4 || argc == 5)) {
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
    const cost = toNumberExact(tagStack[base], valueStack[base]);
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const month = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 12.0;
    const depreciation = dbDepreciation(cost, salvage, life, period, month);
    if (isNaN(depreciation)) {
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
      depreciation,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ddb && (argc == 4 || argc == 5)) {
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
    const cost = toNumberExact(tagStack[base], valueStack[base]);
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const factor = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 2.0;
    const depreciation = ddbDepreciation(cost, salvage, life, period, factor);
    if (isNaN(depreciation)) {
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
      depreciation,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Vdb && argc >= 5 && argc <= 7) {
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
    const cost = toNumberExact(tagStack[base], valueStack[base]);
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const startPeriod = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const endPeriod = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const factor = argc >= 6 ? toNumberExact(tagStack[base + 5], valueStack[base + 5]) : 2.0;
    const noSwitch = argc >= 7 ? coerceBoolean(tagStack[base + 6], valueStack[base + 6]) : 0;
    const depreciation =
      noSwitch < 0
        ? NaN
        : vdbDepreciation(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch != 0);
    if (isNaN(depreciation)) {
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
      depreciation,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sln && argc == 3) {
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
    const cost = toNumberExact(tagStack[base], valueStack[base]);
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    if (isNaN(cost) || isNaN(salvage) || isNaN(life) || life <= 0) {
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
      (cost - salvage) / life,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Syd && argc == 4) {
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
    const cost = toNumberExact(tagStack[base], valueStack[base]);
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    if (
      isNaN(cost) ||
      isNaN(salvage) ||
      isNaN(life) ||
      isNaN(period) ||
      life <= 0 ||
      period <= 0 ||
      period > life
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
    const denominator = (life * (life + 1.0)) / 2.0;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      ((cost - salvage) * (life - period + 1.0)) / denominator,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Disc && (argc == 4 || argc == 5)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const value =
      isNaN(price) || isNaN(redemption) || redemption <= 0.0 || price <= 0.0 || isNaN(years)
        ? NaN
        : (redemption - price) / redemption / years;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Coupdaybs ||
      builtinId == BuiltinId.Coupdays ||
      builtinId == BuiltinId.Coupdaysnc ||
      builtinId == BuiltinId.Coupncd ||
      builtinId == BuiltinId.Coupnum ||
      builtinId == BuiltinId.Couppcd) &&
    (argc == 3 || argc == 4)
  ) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const frequency = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const basis = argc == 4 ? truncToInt(tagStack[base + 3], valueStack[base + 3]) : 0;
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency);
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining, frequency);
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency);
    const accruedDays =
      previousCoupon == i32.MIN_VALUE || settlement == i32.MIN_VALUE
        ? NaN
        : couponDaysByBasisValue(previousCoupon, settlement, basis);
    const daysToNextCoupon =
      nextCoupon == i32.MIN_VALUE || settlement == i32.MIN_VALUE
        ? NaN
        : couponDaysByBasisValue(settlement, nextCoupon, basis);
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency);
    let value = NaN;
    if (builtinId == BuiltinId.Coupdaybs) {
      value = accruedDays;
    } else if (builtinId == BuiltinId.Coupdays) {
      value = daysInPeriod;
    } else if (builtinId == BuiltinId.Coupdaysnc) {
      value = daysToNextCoupon;
    } else if (builtinId == BuiltinId.Coupncd) {
      value = nextCoupon == i32.MIN_VALUE ? NaN : <f64>nextCoupon;
    } else if (builtinId == BuiltinId.Coupnum) {
      value = periodsRemaining == i32.MIN_VALUE ? NaN : <f64>periodsRemaining;
    } else {
      value = previousCoupon == i32.MIN_VALUE ? NaN : <f64>previousCoupon;
    }
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Intrate && (argc == 4 || argc == 5)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const investment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const value =
      isNaN(investment) ||
      isNaN(redemption) ||
      investment <= 0.0 ||
      redemption <= 0.0 ||
      isNaN(years)
        ? NaN
        : (redemption - investment) / investment / years;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Received && (argc == 4 || argc == 5)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const investment = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const discount = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const denominator = isNaN(years) ? NaN : 1.0 - discount * years;
    const value =
      isNaN(investment) ||
      isNaN(discount) ||
      investment <= 0.0 ||
      discount <= 0.0 ||
      !isFinite(denominator) ||
      denominator <= 0.0
        ? NaN
        : investment / denominator;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Pricedisc && (argc == 4 || argc == 5)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const value =
      isNaN(discount) || isNaN(redemption) || discount <= 0.0 || redemption <= 0.0 || isNaN(years)
        ? NaN
        : redemption * (1.0 - discount * years);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Yielddisc && (argc == 4 || argc == 5)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const value =
      isNaN(price) || isNaN(redemption) || price <= 0.0 || redemption <= 0.0 || isNaN(years)
        ? NaN
        : (redemption - price) / price / years;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tbillprice && argc == 3) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const days =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : treasuryBillDaysValue(settlement, maturity);
    const value =
      isNaN(discount) || discount <= 0.0 || isNaN(days)
        ? NaN
        : 100.0 * (1.0 - (discount * days) / 360.0);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tbillyield && argc == 3) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const days =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : treasuryBillDaysValue(settlement, maturity);
    const value =
      isNaN(price) || price <= 0.0 || isNaN(days)
        ? NaN
        : ((100.0 - price) * 360.0) / (price * days);
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tbilleq && argc == 3) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const days =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : treasuryBillDaysValue(settlement, maturity);
    const denominator = isNaN(days) ? NaN : 360.0 - discount * days;
    const value =
      isNaN(discount) || discount <= 0.0 || !isFinite(denominator) || denominator == 0.0
        ? NaN
        : (365.0 * discount) / denominator;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Pricemat && (argc == 5 || argc == 6)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const yieldRate = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0;
    const issueToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || issue == i32.MIN_VALUE
        ? NaN
        : maturityIssueYearfracValue(issue, settlement, maturity, basis);
    const settlementToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const issueToSettlement =
      settlement == i32.MIN_VALUE || issue == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : accruedIssueYearfracValue(issue, settlement, maturity, basis);
    const denominator = isNaN(settlementToMaturity) ? NaN : 1.0 + yieldRate * settlementToMaturity;
    const value =
      isNaN(rate) ||
      isNaN(yieldRate) ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      isNaN(issueToMaturity) ||
      isNaN(issueToSettlement) ||
      !isFinite(denominator) ||
      denominator <= 0.0
        ? NaN
        : (100.0 * (1.0 + rate * issueToMaturity)) / denominator - 100.0 * rate * issueToSettlement;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Yieldmat && (argc == 5 || argc == 6)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const price = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0;
    const issueToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || issue == i32.MIN_VALUE
        ? NaN
        : maturityIssueYearfracValue(issue, settlement, maturity, basis);
    const settlementToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : securityAnnualizedYearfracValue(settlement, maturity, basis);
    const issueToSettlement =
      settlement == i32.MIN_VALUE || issue == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : accruedIssueYearfracValue(issue, settlement, maturity, basis);
    const settlementValue =
      isNaN(price) || isNaN(rate) || isNaN(issueToSettlement)
        ? NaN
        : price + 100.0 * rate * issueToSettlement;
    const value =
      isNaN(rate) ||
      isNaN(price) ||
      rate < 0.0 ||
      price <= 0.0 ||
      isNaN(issueToMaturity) ||
      isNaN(settlementToMaturity) ||
      !isFinite(settlementValue) ||
      settlementValue <= 0.0
        ? NaN
        : ((100.0 * (1.0 + rate * issueToMaturity)) / settlementValue - 1.0) / settlementToMaturity;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Oddlprice && (argc == 7 || argc == 8)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const lastInterest = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const yieldRate = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const redemption = toNumberExact(tagStack[base + 5], valueStack[base + 5]);
    const frequency = truncToInt(tagStack[base + 6], valueStack[base + 6]);
    const basis = argc == 8 ? truncToInt(tagStack[base + 7], valueStack[base + 7]) : 0;
    const value =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || lastInterest == i32.MIN_VALUE
        ? NaN
        : oddLastPriceValue(
            settlement,
            maturity,
            lastInterest,
            rate,
            yieldRate,
            redemption,
            frequency,
            basis,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Oddlyield && (argc == 7 || argc == 8)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const lastInterest = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const price = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const redemption = toNumberExact(tagStack[base + 5], valueStack[base + 5]);
    const frequency = truncToInt(tagStack[base + 6], valueStack[base + 6]);
    const basis = argc == 8 ? truncToInt(tagStack[base + 7], valueStack[base + 7]) : 0;
    const value =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || lastInterest == i32.MIN_VALUE
        ? NaN
        : oddLastYieldValue(
            settlement,
            maturity,
            lastInterest,
            rate,
            price,
            redemption,
            frequency,
            basis,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Oddfprice && (argc == 8 || argc == 9)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const firstCoupon = excelSerialWhole(tagStack[base + 3], valueStack[base + 3]);
    const rate = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const yieldRate = toNumberExact(tagStack[base + 5], valueStack[base + 5]);
    const redemption = toNumberExact(tagStack[base + 6], valueStack[base + 6]);
    const frequency = truncToInt(tagStack[base + 7], valueStack[base + 7]);
    const basis = argc == 9 ? truncToInt(tagStack[base + 8], valueStack[base + 8]) : 0;
    const value =
      settlement == i32.MIN_VALUE ||
      maturity == i32.MIN_VALUE ||
      issue == i32.MIN_VALUE ||
      firstCoupon == i32.MIN_VALUE ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      redemption <= 0.0
        ? NaN
        : oddFirstPriceValue(
            settlement,
            maturity,
            issue,
            firstCoupon,
            rate,
            yieldRate,
            redemption,
            frequency,
            basis,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Oddfyield && (argc == 8 || argc == 9)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2]);
    const firstCoupon = excelSerialWhole(tagStack[base + 3], valueStack[base + 3]);
    const rate = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const price = toNumberExact(tagStack[base + 5], valueStack[base + 5]);
    const redemption = toNumberExact(tagStack[base + 6], valueStack[base + 6]);
    const frequency = truncToInt(tagStack[base + 7], valueStack[base + 7]);
    const basis = argc == 9 ? truncToInt(tagStack[base + 8], valueStack[base + 8]) : 0;
    const value =
      settlement == i32.MIN_VALUE ||
      maturity == i32.MIN_VALUE ||
      issue == i32.MIN_VALUE ||
      firstCoupon == i32.MIN_VALUE
        ? NaN
        : oddFirstYieldValue(
            settlement,
            maturity,
            issue,
            firstCoupon,
            rate,
            price,
            redemption,
            frequency,
            basis,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Price && (argc == 6 || argc == 7)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const rate = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const yieldRate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const redemption = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const frequency = truncToInt(tagStack[base + 5], valueStack[base + 5]);
    const basis = argc == 7 ? truncToInt(tagStack[base + 6], valueStack[base + 6]) : 0;
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency);
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining, frequency);
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency);
    const accruedDays =
      previousCoupon == i32.MIN_VALUE
        ? NaN
        : couponDaysByBasisValue(previousCoupon, settlement, basis);
    const daysToNextCoupon =
      nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis);
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency);
    const value =
      isNaN(rate) ||
      isNaN(yieldRate) ||
      isNaN(redemption) ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      redemption <= 0.0 ||
      periodsRemaining == i32.MIN_VALUE
        ? NaN
        : couponPriceFromMetricsValue(
            periodsRemaining,
            accruedDays,
            daysToNextCoupon,
            daysInPeriod,
            rate,
            yieldRate,
            redemption,
            frequency,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Yield && (argc == 6 || argc == 7)) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const rate = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const price = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const redemption = toNumberExact(tagStack[base + 4], valueStack[base + 4]);
    const frequency = truncToInt(tagStack[base + 5], valueStack[base + 5]);
    const basis = argc == 7 ? truncToInt(tagStack[base + 6], valueStack[base + 6]) : 0;
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency);
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining, frequency);
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency);
    const accruedDays =
      previousCoupon == i32.MIN_VALUE
        ? NaN
        : couponDaysByBasisValue(previousCoupon, settlement, basis);
    const daysToNextCoupon =
      nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis);
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency);
    const value =
      isNaN(rate) ||
      isNaN(price) ||
      isNaN(redemption) ||
      rate < 0.0 ||
      price <= 0.0 ||
      redemption <= 0.0 ||
      periodsRemaining == i32.MIN_VALUE
        ? NaN
        : solveCouponYieldValue(
            periodsRemaining,
            accruedDays,
            daysToNextCoupon,
            daysInPeriod,
            rate,
            price,
            redemption,
            frequency,
          );
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.Duration || builtinId == BuiltinId.Mduration) &&
    (argc == 5 || argc == 6)
  ) {
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
    const settlement = excelSerialWhole(tagStack[base], valueStack[base]);
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    const couponRate = toNumberExact(tagStack[base + 2], valueStack[base + 2]);
    const yieldRate = toNumberExact(tagStack[base + 3], valueStack[base + 3]);
    const frequency = truncToInt(tagStack[base + 4], valueStack[base + 4]);
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0;
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency);
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining, frequency);
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency);
    const accruedDays =
      previousCoupon == i32.MIN_VALUE
        ? NaN
        : couponDaysByBasisValue(previousCoupon, settlement, basis);
    const daysToNextCoupon =
      nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis);
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency);
    const duration =
      isNaN(couponRate) ||
      isNaN(yieldRate) ||
      couponRate < 0.0 ||
      yieldRate < 0.0 ||
      periodsRemaining == i32.MIN_VALUE
        ? NaN
        : macaulayDurationValue(
            periodsRemaining,
            accruedDays,
            daysToNextCoupon,
            daysInPeriod,
            couponRate,
            yieldRate,
            frequency,
          );
    const value =
      builtinId == BuiltinId.Mduration && !isNaN(duration)
        ? duration / (1.0 + yieldRate / <f64>frequency)
        : duration;
    if (isNaN(value)) {
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
      value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Replace && argc == 4) {
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
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    const replacement = scalarText(
      tagStack[base + 3],
      valueStack[base + 3],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE || replacement == null) {
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
      replaceText(text, start, count, replacement),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Substitute && (argc == 3 || argc == 4)) {
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
    const oldText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const newText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || oldText == null || newText == null || oldText.length == 0) {
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
    if (argc == 3) {
      return writeStringResult(
        base,
        substituteText(text, oldText, newText),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const instance = coercePositiveStart(tagStack[base + 3], valueStack[base + 3], 1);
    if (instance == i32.MIN_VALUE) {
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
      substituteNthText(text, oldText, newText, instance),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rept && argc == 2) {
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
    const count = coerceNonNegativeLength(tagStack[base + 1], valueStack[base + 1]);
    if (text == null || count == i32.MIN_VALUE) {
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
      repeatText(text, count),
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
