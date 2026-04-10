import { ErrorCode, type CellValue } from "@bilig/protocol";
import type { LookupBuiltin, LookupBuiltinArgument } from "./lookup.js";

interface LookupHypothesisBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue;
  chiSquareTestResult: (
    actualArg: LookupBuiltinArgument,
    expectedArg: LookupBuiltinArgument,
  ) => CellValue;
  fTestResult: (firstArg: LookupBuiltinArgument, secondArg: LookupBuiltinArgument) => CellValue;
  zTestResult: (
    arrayArg: LookupBuiltinArgument,
    xArg: LookupBuiltinArgument,
    sigmaArg?: LookupBuiltinArgument,
  ) => CellValue;
  tTestResult: (
    firstArg: LookupBuiltinArgument,
    secondArg: LookupBuiltinArgument,
    tailsArg: LookupBuiltinArgument,
    typeArg: LookupBuiltinArgument,
  ) => CellValue;
}

export function createLookupHypothesisBuiltins(
  deps: LookupHypothesisBuiltinDeps,
): Record<string, LookupBuiltin> {
  return {
    "CHISQ.TEST": (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.chiSquareTestResult(actualArg, expectedArg);
    },
    CHITEST: (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.chiSquareTestResult(actualArg, expectedArg);
    },
    "LEGACY.CHITEST": (actualArg, expectedArg) => {
      return actualArg === undefined || expectedArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.chiSquareTestResult(actualArg, expectedArg);
    },
    "F.TEST": (firstArg, secondArg) => {
      return firstArg === undefined || secondArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.fTestResult(firstArg, secondArg);
    },
    FTEST: (firstArg, secondArg) => {
      return firstArg === undefined || secondArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.fTestResult(firstArg, secondArg);
    },
    "Z.TEST": (arrayArg, xArg, sigmaArg) => {
      return arrayArg === undefined || xArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.zTestResult(arrayArg, xArg, sigmaArg);
    },
    ZTEST: (arrayArg, xArg, sigmaArg) => {
      return arrayArg === undefined || xArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.zTestResult(arrayArg, xArg, sigmaArg);
    },
    "T.TEST": (firstArg, secondArg, tailsArg, typeArg) => {
      return firstArg === undefined ||
        secondArg === undefined ||
        tailsArg === undefined ||
        typeArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.tTestResult(firstArg, secondArg, tailsArg, typeArg);
    },
    TTEST: (firstArg, secondArg, tailsArg, typeArg) => {
      return firstArg === undefined ||
        secondArg === undefined ||
        tailsArg === undefined ||
        typeArg === undefined
        ? deps.errorValue(ErrorCode.Value)
        : deps.tTestResult(firstArg, secondArg, tailsArg, typeArg);
    },
  };
}
