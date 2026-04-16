import type { WorkPaperCellAddress } from './work-paper-types.js'

class BaseWorkPaperError extends Error {
  constructor(name: string, message: string) {
    super(message)
    this.name = name
  }
}

export class WorkPaperConfigValueTooBigError extends BaseWorkPaperError {
  constructor(paramName: string, maximum: number) {
    super('WorkPaperConfigValueTooBigError', `Config parameter ${paramName} should be at most ${maximum}`)
  }
}

export class WorkPaperConfigValueTooSmallError extends BaseWorkPaperError {
  constructor(paramName: string, minimum: number) {
    super('WorkPaperConfigValueTooSmallError', `Config parameter ${paramName} should be at least ${minimum}`)
  }
}

export class WorkPaperEvaluationSuspendedError extends BaseWorkPaperError {
  constructor(message = 'Computations are suspended') {
    super('WorkPaperEvaluationSuspendedError', message)
  }
}

export class WorkPaperExpectedOneOfValuesError extends BaseWorkPaperError {
  constructor(values: string, paramName: string) {
    super('WorkPaperExpectedOneOfValuesError', `Expected one of ${values} for config parameter: ${paramName}`)
  }
}

export class WorkPaperExpectedValueOfTypeError extends BaseWorkPaperError {
  constructor(expectedType: string, paramName: string) {
    super('WorkPaperExpectedValueOfTypeError', `Expected value of type: ${expectedType} for config parameter: ${paramName}`)
  }
}

export class WorkPaperFunctionPluginValidationError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperFunctionPluginValidationError', message)
  }

  static functionNotDeclaredInPlugin(functionId: string, pluginName: string): WorkPaperFunctionPluginValidationError {
    return new WorkPaperFunctionPluginValidationError(`Function with id ${functionId} not declared in plugin ${pluginName}`)
  }

  static functionMethodNotFound(functionName: string, pluginName: string): WorkPaperFunctionPluginValidationError {
    return new WorkPaperFunctionPluginValidationError(`Function method ${functionName} not found in plugin ${pluginName}`)
  }
}

export class WorkPaperInvalidAddressError extends BaseWorkPaperError {
  constructor(address: WorkPaperCellAddress) {
    super('WorkPaperInvalidAddressError', `Address (row = ${address.row}, col = ${address.col}) is invalid`)
  }
}

export class WorkPaperInvalidArgumentsError extends BaseWorkPaperError {
  constructor(expectedArguments: string) {
    super('WorkPaperInvalidArgumentsError', `Invalid arguments, expected ${expectedArguments}`)
  }
}

export class WorkPaperLanguageAlreadyRegisteredError extends BaseWorkPaperError {
  constructor(languageCode?: string) {
    super(
      'WorkPaperLanguageAlreadyRegisteredError',
      languageCode ? `Language '${languageCode}' is already registered` : 'Language already registered.',
    )
  }
}

export class WorkPaperLanguageNotRegisteredError extends BaseWorkPaperError {
  constructor(languageCode?: string) {
    super('WorkPaperLanguageNotRegisteredError', languageCode ? `Language '${languageCode}' is not registered` : 'Language not registered.')
  }
}

export class WorkPaperMissingTranslationError extends BaseWorkPaperError {
  constructor(key: string) {
    super('WorkPaperMissingTranslationError', `Translation for ${key} is missing in the translation package you're using.`)
  }
}

export class WorkPaperNamedExpressionDoesNotExistError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super('WorkPaperNamedExpressionDoesNotExistError', `Named Expression '${expressionName}' does not exist`)
  }
}

export class WorkPaperNamedExpressionNameIsAlreadyTakenError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super('WorkPaperNamedExpressionNameIsAlreadyTakenError', `Name of Named Expression '${expressionName}' is already present`)
  }
}

export class WorkPaperNamedExpressionNameIsInvalidError extends BaseWorkPaperError {
  constructor(expressionName: string) {
    super('WorkPaperNamedExpressionNameIsInvalidError', `Name of Named Expression '${expressionName}' is invalid`)
  }
}

export class WorkPaperNoOperationToRedoError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperNoOperationToRedoError', 'There is no operation to redo')
  }
}

export class WorkPaperNoOperationToUndoError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperNoOperationToUndoError', 'There is no operation to undo')
  }
}

export class WorkPaperNoRelativeAddressesAllowedError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperNoRelativeAddressesAllowedError', 'Relative addresses not allowed in named expressions.')
  }
}

export class WorkPaperNoSheetWithIdError extends BaseWorkPaperError {
  constructor(sheetId: number) {
    super('WorkPaperNoSheetWithIdError', `There's no sheet with id = ${sheetId}`)
  }
}

export class WorkPaperNoSheetWithNameError extends BaseWorkPaperError {
  constructor(sheetName: string) {
    super('WorkPaperNoSheetWithNameError', `There's no sheet with name '${sheetName}'`)
  }
}

export class WorkPaperNotAFormulaError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperNotAFormulaError', 'This is not a formula')
  }
}

export class WorkPaperNothingToPasteError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperNothingToPasteError', 'There is nothing to paste')
  }
}

export class WorkPaperProtectedFunctionTranslationError extends BaseWorkPaperError {
  constructor(functionId: string) {
    super('WorkPaperProtectedFunctionTranslationError', `Cannot register translation for function with id: ${functionId}`)
  }
}

export class WorkPaperSheetNameAlreadyTakenError extends BaseWorkPaperError {
  constructor(sheetName: string) {
    super('WorkPaperSheetNameAlreadyTakenError', `Sheet with name ${sheetName} already exists`)
  }
}

export class WorkPaperSheetSizeLimitExceededError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperSheetSizeLimitExceededError', 'Sheet size limit exceeded')
  }
}

export class WorkPaperSourceLocationHasArrayError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperSourceLocationHasArrayError', 'Cannot perform this operation, source location has an array inside.')
  }
}

export class WorkPaperTargetLocationHasArrayError extends BaseWorkPaperError {
  constructor() {
    super('WorkPaperTargetLocationHasArrayError', 'Cannot perform this operation, target location has an array inside.')
  }
}

function serializeParseValue(value: unknown): string {
  return JSON.stringify(
    value,
    (_key, current) => {
      if (typeof current === 'function' || typeof current === 'symbol') {
        return String(current)
      }
      if (typeof current === 'bigint') {
        return `BigInt(${current.toString()})`
      }
      if (current instanceof RegExp) {
        return `RegExp(${current.toString()})`
      }
      return current
    },
    4,
  )
}

export class WorkPaperUnableToParseError extends BaseWorkPaperError {
  constructor(value: unknown) {
    super('WorkPaperUnableToParseError', `Unable to parse value: ${serializeParseValue(value)}`)
  }
}

export class WorkPaperConfigError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperConfigError', message)
  }
}

export class WorkPaperSheetError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperSheetError', message)
  }
}

export class WorkPaperNamedExpressionError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperNamedExpressionError', message)
  }
}

export class WorkPaperClipboardError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperClipboardError', message)
  }
}

export class WorkPaperParseError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperParseError', message)
  }
}

export class WorkPaperOperationError extends BaseWorkPaperError {
  constructor(message: string) {
    super('WorkPaperOperationError', message)
  }
}
