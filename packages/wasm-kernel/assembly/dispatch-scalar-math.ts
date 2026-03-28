import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  doubleFactorialCalc,
  evenCalc,
  factorialCalc,
  oddCalc,
  roundToDigits,
  roundTowardZeroDigits,
  truncToInt,
} from "./numeric-core";
import { toNumberExact, toNumberOrZero } from "./operands";
import { besselIValue, besselJValue, besselKValue, besselYValue } from "./distributions";
import { STACK_KIND_SCALAR, writeResult } from "./result-io";

function writeScalarMathError(
  base: i32,
  error: ErrorCode,
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

function writeScalarMathNumber(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
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

export function tryApplyScalarMathBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const order = <i32>orderNumeric;
    if (order < 0) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if ((builtinId == BuiltinId.Besselk || builtinId == BuiltinId.Bessely) && x <= 0.0) {
      return writeScalarMathError(
        base,
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
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Abs && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.abs(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    return isNaN(numeric)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(
          base,
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
    return isNaN(numeric) || isNaN(digits)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(
          base,
          roundToDigits(numeric, <i32>digits),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    return isNaN(numeric)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(
          base,
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (significance == 0.0) {
      return writeScalarMathError(
        base,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
      Math.floor(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    return isNaN(numeric)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(
          base,
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (significance == 0.0) {
      return writeScalarMathError(
        base,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
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
      return writeScalarMathError(
        base,
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
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.FloorPrecise && (argc == 1 || argc == 2)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significanceRaw =
      argc == 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 1.0;
    const significance = isNaN(significanceRaw) ? 1.0 : Math.abs(significanceRaw);
    if (isNaN(numeric) || significance == 0.0) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
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
      return writeScalarMathError(
        base,
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
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
      Math.ceil(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = toNumberOrZero(tagStack[base + 1], valueStack[base + 1]);
    if (divisor == 0.0) {
      return writeScalarMathError(
        base,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
      toNumberOrZero(tagStack[base], valueStack[base]) % divisor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Int && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    return isNaN(numeric)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(
          base,
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
      return writeScalarMathError(
        base,
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
          ? numeric >= 0.0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0.0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) / factor;
    } else {
      const factor = Math.pow(10.0, <f64>-digits);
      const scaled = numeric / factor;
      result =
        (builtinId == BuiltinId.RoundUp
          ? numeric >= 0.0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0.0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) * factor;
    }
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Trunc && (argc == 1 || argc == 2)) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (isNaN(numeric) || digits == i32.MIN_VALUE) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
      roundTowardZeroDigits(numeric, digits),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sin && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.sin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Cos && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.cos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Tan && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.tan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Asin && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.asin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Acos && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.acos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.atan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan2 && argc == 2) {
    return writeScalarMathNumber(
      base,
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
    return writeScalarMathNumber(
      base,
      (toNumberOrZero(tagStack[base], valueStack[base]) * 180.0) / Math.PI,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Radians && argc == 1) {
    return writeScalarMathNumber(
      base,
      (toNumberOrZero(tagStack[base], valueStack[base]) * Math.PI) / 180.0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Exp && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.exp(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Ln && argc == 1) {
    return writeScalarMathNumber(
      base,
      Math.log(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Log10 && argc == 1) {
    return writeScalarMathNumber(
      base,
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
    return writeScalarMathNumber(
      base,
      Math.log(num) / Math.log(baseVal),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Power && argc == 2) {
    return writeScalarMathNumber(
      base,
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
    return writeScalarMathNumber(
      base,
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
      return writeScalarMathError(
        base,
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
    return writeScalarMathNumber(base, sum, rangeIndexStack, valueStack, tagStack, kindStack);
  }
  if (builtinId == BuiltinId.Sqrtpi && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const result = isNaN(numeric) ? NaN : Math.sqrt(numeric * Math.PI);
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }
  if (builtinId == BuiltinId.Pi && argc == 0) {
    return writeScalarMathNumber(base, Math.PI, rangeIndexStack, valueStack, tagStack, kindStack);
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
      return writeScalarMathError(
        base,
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
      return writeScalarMathError(
        base,
        errorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (builtinId == BuiltinId.Quotient) {
      if (right == 0.0) {
        return writeScalarMathError(
          base,
          ErrorCode.Div0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      return writeScalarMathNumber(
        base,
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
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (builtinId == BuiltinId.Combin && chosenValue > numberValue) {
      return writeScalarMathError(
        base,
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
    } else if (chosenValue == 0.0) {
      result = 1.0;
    } else if (numberValue == 0.0) {
      result = 0.0;
    } else {
      const combined = numberValue + chosenValue - 1.0;
      result =
        factorialCalc(combined) / (factorialCalc(chosenValue) * factorialCalc(numberValue - 1.0));
    }
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if ((builtinId == BuiltinId.Permut || builtinId == BuiltinId.Permutationa) && argc == 2) {
    const left = toNumberExact(tagStack[base], valueStack[base]);
    const right = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    const numberValue = !isFinite(left) || left < 0.0 ? NaN : Math.floor(left);
    const chosenValue = !isFinite(right) || right < 0.0 ? NaN : Math.floor(right);
    if (!isFinite(numberValue) || !isFinite(chosenValue)) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (builtinId == BuiltinId.Permut && chosenValue > numberValue) {
      return writeScalarMathError(
        base,
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
    return !isFinite(result)
      ? writeScalarMathError(
          base,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Mround && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const multiple = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(multiple) || multiple == 0.0) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (numeric != 0.0 && Math.sign(numeric) != Math.sign(multiple)) {
      return writeScalarMathError(
        base,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeScalarMathNumber(
      base,
      Math.round(numeric / multiple) * multiple,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return -1;
}
