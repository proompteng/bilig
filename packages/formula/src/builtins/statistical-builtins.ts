import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import {
  inverseNormal,
  inverseStandardNormal,
  kurtosis,
  percentileNormal,
  skewPopulation,
  skewSample,
  standardNormalCdf,
  standardNormalPdf,
} from "./distributions.js";
import { collectNumericArgs, collectStatNumericArgs } from "./numeric.js";
import {
  collectAStyleNumericArgs,
  modeSingle,
  populationStandardDeviation,
  populationVariance,
  sampleStandardDeviation,
  sampleVariance,
} from "./statistics.js";
import type { EvaluationResult } from "../runtime-values.js";

type Builtin = (...args: CellValue[]) => EvaluationResult;

interface StatisticalBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined;
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined;
  firstError: (args: CellValue[]) => CellValue | undefined;
  numberResult: (value: number) => EvaluationResult;
  numericResultOrError: (value: number) => EvaluationResult;
  valueError: () => EvaluationResult;
}

function naError(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.NA };
}

export function createStatisticalBuiltins({
  toNumber,
  coerceBoolean,
  firstError,
  numberResult,
  numericResultOrError,
  valueError,
}: StatisticalBuiltinDeps): Record<string, Builtin> {
  const builtins: Record<string, Builtin> = {
    GAUSS: (value) => {
      const numeric = toNumber(value);
      return numeric === undefined ? valueError() : numberResult(standardNormalCdf(numeric) - 0.5);
    },
    PHI: (value) => {
      const numeric = toNumber(value);
      return numeric === undefined ? valueError() : numberResult(standardNormalPdf(numeric));
    },
    STANDARDIZE: (xArg, meanArg, standardDeviationArg) => {
      const x = toNumber(xArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      if (
        x === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        standardDeviation <= 0
      ) {
        return valueError();
      }
      return numberResult((x - mean) / standardDeviation);
    },
    MODE: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const mode = modeSingle(collectNumericArgs(args, toNumber));
      return mode === undefined ? naError() : numberResult(mode);
    },
    "MODE.SNGL": (...args) => builtins["MODE"]!(...args),
    STDEV: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectStatNumericArgs(args);
      return numbers.length < 2
        ? valueError()
        : numericResultOrError(sampleStandardDeviation(numbers));
    },
    "STDEV.S": (...args) => builtins["STDEV"]!(...args),
    STDEVP: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectStatNumericArgs(args);
      return numbers.length === 0
        ? valueError()
        : numericResultOrError(populationStandardDeviation(numbers));
    },
    "STDEV.P": (...args) => builtins["STDEVP"]!(...args),
    STDEVA: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectAStyleNumericArgs(args);
      return numbers.length < 2
        ? valueError()
        : numericResultOrError(sampleStandardDeviation(numbers));
    },
    STDEVPA: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectAStyleNumericArgs(args);
      return numbers.length === 0
        ? valueError()
        : numericResultOrError(populationStandardDeviation(numbers));
    },
    VAR: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectStatNumericArgs(args);
      return numbers.length < 2 ? valueError() : numberResult(sampleVariance(numbers));
    },
    "VAR.S": (...args) => builtins["VAR"]!(...args),
    VARP: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectStatNumericArgs(args);
      return numbers.length === 0 ? valueError() : numberResult(populationVariance(numbers));
    },
    "VAR.P": (...args) => builtins["VARP"]!(...args),
    VARA: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectAStyleNumericArgs(args);
      return numbers.length < 2 ? valueError() : numberResult(sampleVariance(numbers));
    },
    VARPA: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const numbers = collectAStyleNumericArgs(args);
      return numbers.length === 0 ? valueError() : numberResult(populationVariance(numbers));
    },
    SKEW: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const skew = skewSample(collectStatNumericArgs(args));
      return skew === undefined ? valueError() : numberResult(skew);
    },
    "SKEW.P": (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const skew = skewPopulation(collectStatNumericArgs(args));
      return skew === undefined ? valueError() : numberResult(skew);
    },
    SKEWP: (...args) => builtins["SKEW.P"]!(...args),
    KURT: (...args) => {
      const error = firstError(args);
      if (error) {
        return error;
      }
      const value = kurtosis(collectStatNumericArgs(args));
      return value === undefined ? valueError() : numberResult(value);
    },
    NORMDIST: (xArg, meanArg, standardDeviationArg, cumulativeArg) => {
      const x = toNumber(xArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      const cumulative = coerceBoolean(cumulativeArg, false);
      if (
        x === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        cumulative === undefined ||
        standardDeviation <= 0
      ) {
        return valueError();
      }
      return numberResult(
        cumulative
          ? percentileNormal(mean, standardDeviation, x)
          : standardNormalPdf((x - mean) / standardDeviation) / standardDeviation,
      );
    },
    "NORM.DIST": (xArg, meanArg, standardDeviationArg, cumulativeArg) =>
      builtins["NORMDIST"]!(xArg, meanArg, standardDeviationArg, cumulativeArg),
    NORMINV: (probabilityArg, meanArg, standardDeviationArg) => {
      const probability = toNumber(probabilityArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      if (
        probability === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        standardDeviation <= 0
      ) {
        return valueError();
      }
      const result = inverseNormal(probability, mean, standardDeviation);
      return result === undefined ? valueError() : numberResult(result);
    },
    "NORM.INV": (probabilityArg, meanArg, standardDeviationArg) =>
      builtins["NORMINV"]!(probabilityArg, meanArg, standardDeviationArg),
    NORMSDIST: (value) => {
      const numeric = toNumber(value);
      return numeric === undefined ? valueError() : numberResult(standardNormalCdf(numeric));
    },
    "LEGACY.NORMSDIST": (value) => builtins["NORMSDIST"]!(value),
    "NORM.S.DIST": (value, cumulativeArg = { tag: ValueTag.Boolean, value: true }) => {
      const numeric = toNumber(value);
      const cumulative = coerceBoolean(cumulativeArg, true);
      if (numeric === undefined || cumulative === undefined) {
        return valueError();
      }
      return numberResult(cumulative ? standardNormalCdf(numeric) : standardNormalPdf(numeric));
    },
    NORMSINV: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const result = inverseStandardNormal(numeric);
      return result === undefined ? valueError() : numberResult(result);
    },
    "LEGACY.NORMSINV": (value) => builtins["NORMSINV"]!(value),
    "NORM.S.INV": (value) => builtins["NORMSINV"]!(value),
    LOGINV: (probabilityArg, meanArg, standardDeviationArg) => {
      const probability = toNumber(probabilityArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      if (
        probability === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        standardDeviation <= 0
      ) {
        return valueError();
      }
      const normal = inverseNormal(probability, mean, standardDeviation);
      return normal === undefined ? valueError() : numberResult(Math.exp(normal));
    },
    "LOGNORM.INV": (probabilityArg, meanArg, standardDeviationArg) =>
      builtins["LOGINV"]!(probabilityArg, meanArg, standardDeviationArg),
    LOGNORMDIST: (xArg, meanArg, standardDeviationArg) => {
      const x = toNumber(xArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      if (
        x === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        standardDeviation <= 0 ||
        x <= 0
      ) {
        return valueError();
      }
      return numberResult(percentileNormal(mean, standardDeviation, Math.log(x)));
    },
    "LOGNORM.DIST": (
      xArg,
      meanArg,
      standardDeviationArg,
      cumulativeArg = { tag: ValueTag.Boolean, value: true },
    ) => {
      const x = toNumber(xArg);
      const mean = toNumber(meanArg);
      const standardDeviation = toNumber(standardDeviationArg);
      const cumulative = coerceBoolean(cumulativeArg, true);
      if (
        x === undefined ||
        mean === undefined ||
        standardDeviation === undefined ||
        cumulative === undefined ||
        standardDeviation <= 0 ||
        x <= 0
      ) {
        return valueError();
      }
      const z = (Math.log(x) - mean) / standardDeviation;
      return numberResult(
        cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation),
      );
    },
  };

  return builtins;
}
