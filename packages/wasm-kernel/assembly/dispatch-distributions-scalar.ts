import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { valueNumber } from './comparison'
import { toNumberExact } from './operands'
import { isNumericResult, rangeSupportedScalarOnly, scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
import {
  erfApprox,
  gammaFunction,
  gammaDistributionCdf,
  gammaDistributionDensity,
  inverseGammaDistribution,
  inverseStandardNormal,
  inverseStudentT,
  logGamma,
  standardNormalCdf,
  standardNormalPdf,
} from './distributions'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

export function tryApplyScalarDistributionBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if ((builtinId == BuiltinId.Gauss || builtinId == BuiltinId.Phi) && argc == 1) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
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
    )
    if (isNaN(numeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
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
    )
  }

  if (builtinId == BuiltinId.Erf && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const lower = toNumberExact(tagStack[base], valueStack[base])
    const upper = argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 0.0
    if (isNaN(lower) || (argc == 2 && isNaN(upper))) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
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
    )
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
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const value = toNumberExact(tagStack[base], valueStack[base])
    if (isNaN(value)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = NaN
    if (builtinId == BuiltinId.ErfPrecise) {
      result = erfApprox(value)
    } else if (builtinId == BuiltinId.Erfc || builtinId == BuiltinId.ErfcPrecise) {
      result = 1.0 - erfApprox(value)
    } else if (builtinId == BuiltinId.Fisher) {
      result = value <= -1.0 || value >= 1.0 ? NaN : 0.5 * Math.log((1.0 + value) / (1.0 - value))
    } else if (builtinId == BuiltinId.Fisherinv) {
      const exponent = Math.exp(2.0 * value)
      result = (exponent - 1.0) / (exponent + 1.0)
    } else if (builtinId == BuiltinId.Gammaln || builtinId == BuiltinId.GammalnPrecise) {
      result = logGamma(value)
    } else if (builtinId == BuiltinId.Gamma) {
      result = gammaFunction(value)
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
    )
  }

  if ((builtinId == BuiltinId.ConfidenceNorm || builtinId == BuiltinId.Confidence || builtinId == BuiltinId.ConfidenceT) && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const alpha = toNumberExact(tagStack[base], valueStack[base])
    const standardDeviation = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const size = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const useNormal = builtinId == BuiltinId.ConfidenceNorm || builtinId == BuiltinId.Confidence
    const critical = useNormal ? inverseStandardNormal(1.0 - alpha / 2.0) : inverseStudentT(1.0 - alpha / 2.0, size - 1.0)
    const result =
      isNaN(alpha) ||
      isNaN(standardDeviation) ||
      isNaN(size) ||
      !(alpha > 0.0 && alpha < 1.0) ||
      standardDeviation <= 0.0 ||
      (useNormal ? size < 1.0 : size < 2.0) ||
      isNaN(critical)
        ? NaN
        : (critical * standardDeviation) / Math.sqrt(size)
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
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
    ((builtinId == BuiltinId.Lognormdist || builtinId == BuiltinId.LognormDist) && (argc == 3 || argc == 4))
  ) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = NaN
    if (builtinId == BuiltinId.Standardize) {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      result = isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0) ? NaN : (x - mean) / standardDeviation
    } else if (builtinId == BuiltinId.Normdist || builtinId == BuiltinId.NormDist) {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = coerceBoolean(tagStack[base + 3], valueStack[base + 3])
      result =
        isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || cumulative < 0 || !(standardDeviation > 0.0)
          ? NaN
          : cumulative == 1
            ? standardNormalCdf((x - mean) / standardDeviation)
            : standardNormalPdf((x - mean) / standardDeviation) / standardDeviation
    } else if (builtinId == BuiltinId.Norminv || builtinId == BuiltinId.NormInv) {
      const probability = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const inverse = inverseStandardNormal(probability)
      result =
        isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0) || isNaN(inverse) ? NaN : mean + standardDeviation * inverse
    } else if (builtinId == BuiltinId.Normsdist || builtinId == BuiltinId.NormSDist) {
      const value = toNumberExact(tagStack[base], valueStack[base])
      const cumulative = builtinId == BuiltinId.NormSDist && argc == 2 ? coerceBoolean(tagStack[base + 1], valueStack[base + 1]) : 1
      result = isNaN(value) || cumulative < 0 ? NaN : cumulative == 1 ? standardNormalCdf(value) : standardNormalPdf(value)
    } else if (builtinId == BuiltinId.Normsinv || builtinId == BuiltinId.NormSInv) {
      result = inverseStandardNormal(toNumberExact(tagStack[base], valueStack[base]))
    } else if (builtinId == BuiltinId.Loginv || builtinId == BuiltinId.LognormInv) {
      const probability = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const inverse = inverseStandardNormal(probability)
      result =
        isNaN(mean) || isNaN(standardDeviation) || !(standardDeviation > 0.0) || isNaN(inverse)
          ? NaN
          : Math.exp(mean + standardDeviation * inverse)
    } else {
      const x = toNumberExact(tagStack[base], valueStack[base])
      const mean = toNumberExact(tagStack[base + 1], valueStack[base + 1])
      const standardDeviation = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      const cumulative = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1
      const z =
        isNaN(x) || isNaN(mean) || isNaN(standardDeviation) || x <= 0.0 || !(standardDeviation > 0.0)
          ? NaN
          : (Math.log(x) - mean) / standardDeviation
      result = isNaN(z) || cumulative < 0 ? NaN : cumulative == 1 ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation)
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
    )
  }

  if ((builtinId == BuiltinId.GammaInv || builtinId == BuiltinId.Gammainv) && argc == 3) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const probability = toNumberExact(tagStack[base], valueStack[base])
    const alpha = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const beta = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const result = inverseGammaDistribution(probability, alpha, beta)
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
