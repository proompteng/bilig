import type { HeadlessCellAddress } from "./types.js";

class BaseHeadlessError extends Error {
  constructor(name: string, message: string) {
    super(message);
    this.name = name;
  }
}

export class ConfigValueTooBigError extends BaseHeadlessError {
  constructor(paramName: string, maximum: number) {
    super("ConfigValueTooBigError", `Config parameter ${paramName} should be at most ${maximum}`);
  }
}

export class ConfigValueTooSmallError extends BaseHeadlessError {
  constructor(paramName: string, minimum: number) {
    super(
      "ConfigValueTooSmallError",
      `Config parameter ${paramName} should be at least ${minimum}`,
    );
  }
}

export class EvaluationSuspendedError extends BaseHeadlessError {
  constructor(message = "Computations are suspended") {
    super("EvaluationSuspendedError", message);
  }
}

export class ExpectedOneOfValuesError extends BaseHeadlessError {
  constructor(values: string, paramName: string) {
    super(
      "ExpectedOneOfValuesError",
      `Expected one of ${values} for config parameter: ${paramName}`,
    );
  }
}

export class ExpectedValueOfTypeError extends BaseHeadlessError {
  constructor(expectedType: string, paramName: string) {
    super(
      "ExpectedValueOfTypeError",
      `Expected value of type: ${expectedType} for config parameter: ${paramName}`,
    );
  }
}

export class FunctionPluginValidationError extends BaseHeadlessError {
  constructor(message: string) {
    super("FunctionPluginValidationError", message);
  }

  static functionNotDeclaredInPlugin(
    functionId: string,
    pluginName: string,
  ): FunctionPluginValidationError {
    return new FunctionPluginValidationError(
      `Function with id ${functionId} not declared in plugin ${pluginName}`,
    );
  }

  static functionMethodNotFound(
    functionName: string,
    pluginName: string,
  ): FunctionPluginValidationError {
    return new FunctionPluginValidationError(
      `Function method ${functionName} not found in plugin ${pluginName}`,
    );
  }
}

export class InvalidAddressError extends BaseHeadlessError {
  constructor(address: HeadlessCellAddress) {
    super("InvalidAddressError", `Address (row = ${address.row}, col = ${address.col}) is invalid`);
  }
}

export class InvalidArgumentsError extends BaseHeadlessError {
  constructor(expectedArguments: string) {
    super("InvalidArgumentsError", `Invalid arguments, expected ${expectedArguments}`);
  }
}

export class LanguageAlreadyRegisteredError extends BaseHeadlessError {
  constructor(languageCode?: string) {
    super(
      "LanguageAlreadyRegisteredError",
      languageCode
        ? `Language '${languageCode}' is already registered`
        : "Language already registered.",
    );
  }
}

export class LanguageNotRegisteredError extends BaseHeadlessError {
  constructor(languageCode?: string) {
    super(
      "LanguageNotRegisteredError",
      languageCode ? `Language '${languageCode}' is not registered` : "Language not registered.",
    );
  }
}

export class MissingTranslationError extends BaseHeadlessError {
  constructor(key: string) {
    super(
      "MissingTranslationError",
      `Translation for ${key} is missing in the translation package you're using.`,
    );
  }
}

export class NamedExpressionDoesNotExistError extends BaseHeadlessError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionDoesNotExistError",
      `Named Expression '${expressionName}' does not exist`,
    );
  }
}

export class NamedExpressionNameIsAlreadyTakenError extends BaseHeadlessError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionNameIsAlreadyTakenError",
      `Name of Named Expression '${expressionName}' is already present`,
    );
  }
}

export class NamedExpressionNameIsInvalidError extends BaseHeadlessError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionNameIsInvalidError",
      `Name of Named Expression '${expressionName}' is invalid`,
    );
  }
}

export class NoOperationToRedoError extends BaseHeadlessError {
  constructor() {
    super("NoOperationToRedoError", "There is no operation to redo");
  }
}

export class NoOperationToUndoError extends BaseHeadlessError {
  constructor() {
    super("NoOperationToUndoError", "There is no operation to undo");
  }
}

export class NoRelativeAddressesAllowedError extends BaseHeadlessError {
  constructor() {
    super(
      "NoRelativeAddressesAllowedError",
      "Relative addresses not allowed in named expressions.",
    );
  }
}

export class NoSheetWithIdError extends BaseHeadlessError {
  constructor(sheetId: number) {
    super("NoSheetWithIdError", `There's no sheet with id = ${sheetId}`);
  }
}

export class NoSheetWithNameError extends BaseHeadlessError {
  constructor(sheetName: string) {
    super("NoSheetWithNameError", `There's no sheet with name '${sheetName}'`);
  }
}

export class NotAFormulaError extends BaseHeadlessError {
  constructor() {
    super("NotAFormulaError", "This is not a formula");
  }
}

export class NothingToPasteError extends BaseHeadlessError {
  constructor() {
    super("NothingToPasteError", "There is nothing to paste");
  }
}

export class ProtectedFunctionTranslationError extends BaseHeadlessError {
  constructor(functionId: string) {
    super(
      "ProtectedFunctionTranslationError",
      `Cannot register translation for function with id: ${functionId}`,
    );
  }
}

export class SheetNameAlreadyTakenError extends BaseHeadlessError {
  constructor(sheetName: string) {
    super("SheetNameAlreadyTakenError", `Sheet with name ${sheetName} already exists`);
  }
}

export class SheetSizeLimitExceededError extends BaseHeadlessError {
  constructor() {
    super("SheetSizeLimitExceededError", "Sheet size limit exceeded");
  }
}

export class SourceLocationHasArrayError extends BaseHeadlessError {
  constructor() {
    super(
      "SourceLocationHasArrayError",
      "Cannot perform this operation, source location has an array inside.",
    );
  }
}

export class TargetLocationHasArrayError extends BaseHeadlessError {
  constructor() {
    super(
      "TargetLocationHasArrayError",
      "Cannot perform this operation, target location has an array inside.",
    );
  }
}

function serializeParseValue(value: unknown): string {
  return JSON.stringify(
    value,
    (_key, current) => {
      if (typeof current === "function" || typeof current === "symbol") {
        return String(current);
      }
      if (typeof current === "bigint") {
        return `BigInt(${current.toString()})`;
      }
      if (current instanceof RegExp) {
        return `RegExp(${current.toString()})`;
      }
      return current;
    },
    4,
  );
}

export class UnableToParseError extends BaseHeadlessError {
  constructor(value: unknown) {
    super("UnableToParseError", `Unable to parse value: ${serializeParseValue(value)}`);
  }
}

export class HeadlessArgumentError extends InvalidArgumentsError {
  constructor(expectedArguments: string) {
    super(expectedArguments);
    this.name = "HeadlessArgumentError";
  }
}

export class HeadlessConfigError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessConfigError", message);
  }
}

export class HeadlessSheetError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessSheetError", message);
  }
}

export class HeadlessNamedExpressionError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessNamedExpressionError", message);
  }
}

export class HeadlessClipboardError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessClipboardError", message);
  }
}

export class HeadlessEvaluationSuspendedError extends EvaluationSuspendedError {
  constructor(message = "Computations are suspended") {
    super(message);
    this.name = "HeadlessEvaluationSuspendedError";
  }
}

export class HeadlessParseError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessParseError", message);
  }
}

export class HeadlessOperationError extends BaseHeadlessError {
  constructor(message: string) {
    super("HeadlessOperationError", message);
  }
}
