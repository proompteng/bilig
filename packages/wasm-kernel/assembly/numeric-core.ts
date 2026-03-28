import { toNumberExact } from "./operands";

export function truncAbs(value: f64): f64 {
  if (!isFinite(value)) {
    return NaN;
  }
  return Math.abs(value < 0.0 ? Math.ceil(value) : Math.floor(value));
}

export function factorialCalc(value: f64): f64 {
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

export function doubleFactorialCalc(value: f64): f64 {
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

export function gcdPairCalc(left: f64, right: f64): f64 {
  let a = truncAbs(left);
  let b = truncAbs(right);
  while (b != 0.0) {
    const next = a % b;
    a = b;
    b = next;
  }
  return a;
}

export function lcmPairCalc(left: f64, right: f64): f64 {
  const a = truncAbs(left);
  const b = truncAbs(right);
  if (a == 0.0 || b == 0.0) {
    return 0.0;
  }
  return Math.abs((a * b) / gcdPairCalc(a, b));
}

export function evenCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0;
  const rounded = Math.ceil(Math.abs(value) / 2.0) * 2.0;
  return sign * rounded;
}

export function oddCalc(value: f64): f64 {
  const sign = value < 0.0 ? -1.0 : 1.0;
  const rounded = Math.ceil(Math.abs(value));
  const odd = rounded % 2.0 == 0.0 ? rounded + 1.0 : rounded;
  return sign * odd;
}

export function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric;
}

export function roundToDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits);
    return Math.round(value * factor) / factor;
  }
  const factor = Math.pow(10.0, <f64>-digits);
  return Math.round(value / factor) * factor;
}

export function roundTowardZeroDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits);
    return Math.trunc(value * factor) / factor;
  }
  const factor = Math.pow(10.0, <f64>-digits);
  return Math.trunc(value / factor) * factor;
}

const MAX_SAFE_INTEGER_F64: f64 = 9007199254740991.0;

export function coerceBitwiseUnsigned(tag: u8, value: f64): i64 {
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

export function coerceNonNegativeShift(tag: u8, value: f64): i64 {
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

export function coerceInteger(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  return <i32>numeric;
}
