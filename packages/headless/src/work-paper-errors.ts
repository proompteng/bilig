import type { WorkPaperCellAddress } from "./work-paper-types.js";

class BaseWorkPaperError extends Error {
  constructor(name: string, message: string) {
    super(message);
    this.name = name;
  }
}

export class ConfigValueTooBigError extends BaseWorkPaperError {
  constructor(paramName: string, maximum: number) {
    super("ConfigValueTooBigError", `Config parameter ${paramName} should be at most ${maximum}`);
  }
}

export class ConfigValueTooSmallError extends BaseWorkPaperError {
  constructor(paramName: string, minimum: number) {
    super(
      "ConfigValueTooSmallError",
      `Config parameter ${paramName} should be at least ${minimum}`,
    );
  }
}

export class EvaluationSuspendedError extends BaseWorkPaperError {
  constructor(message = "Computations are suspended") {
    super("EvaluationSuspendedError", message);
  }
}

export class ExpectedOneOfValuesError extends BaseWorkPaperError {
  constructor(values: string, paramName: string) {
    super(
      "ExpectedOneOfValuesError",
      `Expected one of ${values} for config parameter: ${paramName}`,
    );
  }
}

export class ExpectedValueOfTypeError extends BaseWorkPaperError {
  constructor(expectedType: string, paramName: string) {
    super(
      "ExpectedValueOfTypeError",
      `Expected value of type: ${expectedType} for config parameter: ${paramName}`,
    );
  }
}

export class FunctionPluginValidationError extends BaseWorkPaperError {
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

export class InvalidAddressError extends BaseWorkPaperError {
  constructor(address: WorkPaperCellAddress) {
    super("InvalidAddressError", `Address (row = ${address.row}, col = ${address.col}) is invalid`);
  }
}

export class InvalidArgumentsError extends BaseWorkPaperError {
  constructor(expectedArguments: string) {
    super("InvalidArgumentsError", `Invalid arguments, expected ${expectedArguments}`);
  }
}

export class LanguageAlreadyRegisteredError extends BaseWorkPaperError {
  constructor(languageCode?: string) {
    super(
      "LanguageAlreadyRegisteredError",
      languageCode
        ? `Language '${languageCode}' is already registered`
        : "Language already registered.",
    );
  }
}

export class LanguageNotRegisteredError extends BaseWorkPaperError {
  constructor(languageCode?: string) {
    super(
      "LanguageNotRegisteredError",
      languageCode ? `Language '${languageCode}' is not registered` : "Language not registered.",
    );
  }
}

export class MissingTranslationError extends BaseWorkPaperError {
  constructor(key: string) {
    super(
      "MissingTranslationError",
      `Translation for ${key} is missing in the translation package you're using.`,
    );
  }
}

export class NamedExpressionDoesNotExistError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionDoesNotExistError",
      `Named Expression '${expressionName}' does not exist`,
    );
  }
}

export class NamedExpressionNameIsAlreadyTakenError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionNameIsAlreadyTakenError",
      `Name of Named Expression '${expressionName}' is already present`,
    );
  }
}

export class NamedExpressionNameIsInvalidError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super(
      "NamedExpressionNameIsInvalidError",
      `Name of Named Expression '${expressionName}' is invalid`,
    );
  }
}

export class NoOperationToRedoError extends BaseWorkPaperError {
  constructor() {
    super("NoOperationToRedoError", "There is no operation to redo");
  }
}

export class NoOperationToUndoError extends BaseWorkPaperError {
  constructor() {
    super("NoOperationToUndoError", "There is no operation to undo");
  }
}

export class NoRelativeAddressesAllowedError extends BaseWorkPaperError {
  constructor() {
    super(
      "NoRelativeAddressesAllowedError",
      "Relative addresses not allowed in named expressions.",
    );
  }
}

export class NoSheetWithIdError extends BaseWorkPaperError {
  constructor(sheetId: number) {
    super("NoSheetWithIdError", `There's no sheet with id = ${sheetId}`);
  }
}

export class NoSheetWithNameError extends BaseWorkPaperError {
  constructor(sheetName: string) {
    super("NoSheetWithNameError", `There's no sheet with name '${sheetName}'`);
  }
}

export class NotAFormulaError extends BaseWorkPaperError {
  constructor() {
    super("NotAFormulaError", "This is not a formula");
  }
}

export class NothingToPasteError extends BaseWorkPaperError {
  constructor() {
    super("NothingToPasteError", "There is nothing to paste");
  }
}

export class ProtectedFunctionTranslationError extends BaseWorkPaperError {
  constructor(functionId: string) {
    super(
      "ProtectedFunctionTranslationError",
      `Cannot register translation for function with id: ${functionId}`,
    );
  }
}

export class SheetNameAlreadyTakenError extends BaseWorkPaperError {
  constructor(sheetName: string) {
    super("SheetNameAlreadyTakenError", `Sheet with name ${sheetName} already exists`);
  }
}

export class SheetSizeLimitExceededError extends BaseWorkPaperError {
  constructor() {
    super("SheetSizeLimitExceededError", "Sheet size limit exceeded");
  }
}

export class SourceLocationHasArrayError extends BaseWorkPaperError {
  constructor() {
    super(
      "SourceLocationHasArrayError",
      "Cannot perform this operation, source location has an array inside.",
    );
  }
}

export class TargetLocationHasArrayError extends BaseWorkPaperError {
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

export class UnableToParseError extends BaseWorkPaperError {
  constructor(value: unknown) {
    super("UnableToParseError", `Unable to parse value: ${serializeParseValue(value)}`);
  }
}

export class WorkPaperArgumentError extends InvalidArgumentsError {
  constructor(expectedArguments: string) {
    super(expectedArguments);
    this.name = "WorkPaperArgumentError";
  }
}

export class WorkPaperConfigError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperConfigError", message);
  }
}

export class WorkPaperSheetError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperSheetError", message);
  }
}

export class WorkPaperNamedExpressionError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperNamedExpressionError", message);
  }
}

export class WorkPaperClipboardError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperClipboardError", message);
  }
}

export class WorkPaperEvaluationSuspendedError extends EvaluationSuspendedError {
  constructor(message = "Computations are suspended") {
    super(message);
    this.name = "WorkPaperEvaluationSuspendedError";
  }
}

export class WorkPaperParseError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperParseError", message);
  }
}

export class WorkPaperOperationError extends BaseWorkPaperError {
  constructor(message: string) {
    super("WorkPaperOperationError", message);
  }
}
