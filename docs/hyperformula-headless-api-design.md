# HyperFormula Headless API Design For `bilig`

## Goal

Define a stable public headless library API for `bilig` that covers the practical HyperFormula workflow:

- construct workbooks from arrays or named sheets
- mutate cells, ranges, rows, columns, sheets, and named expressions
- inspect values, formulas, dimensions, serialization, and dependencies
- batch mutations, suspend evaluation, and use undo/redo
- support clipboard-style workflows
- register custom functions and language packs

The target is functional parity for headless workflows. The target is not source compatibility with HyperFormula internals, and not a separate HyperFormula-compat package.

## Scope decisions

- Add one new public package: `@bilig/headless`.
- Do not add `@bilig/headless-hf-compat`. That is out of scope.
- Keep `@bilig/core` as the lower-level engine/runtime package.
- Use HyperFormula-like method names directly on the new headless facade where the fit is good.
- Keep sync, replica, and binary-transport APIs out of the initial headless facade. They remain in `@bilig/core`.
- Keep custom functions JS-only in the first headless release.

## Source basis

### HyperFormula reviewed locally

- `/Users/gregkonush/github.com/hyperformula/src/index.ts`
- `/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Emitter.ts`
- `/Users/gregkonush/github.com/hyperformula/src/ConfigParams.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/plugin/FunctionPlugin.ts`
- `/Users/gregkonush/github.com/hyperformula/src/NamedExpressions.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Exporter.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Serialization.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Operations.ts`
- `/Users/gregkonush/github.com/hyperformula/src/ClipboardOperations.ts`

### `bilig` reviewed locally

- `/Users/gregkonush/github.com/bilig/packages/core/src/engine.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/workbook-store.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/events.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/engine/errors.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/__tests__/sheet-ids.test.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/index.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/parser.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/translation.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/external-function-adapter.ts`
- `/Users/gregkonush/github.com/bilig/docs/public-api.md`

## Summary

`bilig` already has enough engine power for a serious headless library:

- `SpreadsheetEngine` covers workbook mutation, undo/redo, dependencies, snapshots, names, tables, spills, pivots, filters, sorts, freeze panes, styles, and formats.
- `WorkbookStore` already preserves stable numeric sheet IDs across export/import/rename.
- `@bilig/formula` already exposes parser, AST, address parsing, reference translation, and formula serialization utilities.

What `bilig` lacks is a cohesive library-facing facade with:

- ergonomic factory construction
- HyperFormula-style method surface
- mutation return values
- typed lifecycle events
- public config/update API
- JS custom-function plugin API
- language pack API

The right design is not to expand `SpreadsheetEngine` until it becomes HyperFormula-shaped. The right design is to wrap it with a dedicated headless facade.

## Public package decision

Create:

- `@bilig/headless`

Do not create:

- `@bilig/headless-hf-compat`

Reason:

- A separate compatibility package would duplicate naming and behavior policy.
- The simpler move is to make `@bilig/headless` directly expose the headless workflow methods that matter.
- `SpreadsheetEngine` keeps its existing low-level naming for product/runtime code.

## HyperFormula public surface inventory

This section inventories the public HyperFormula surface that matters for headless workflows.

### Public symbols

HyperFormula publicly exports these symbol families:

- workbook class: `HyperFormula`
- config: `ConfigParams`
- factories and static registries
- cell and range value/address types
- sheet and named-expression types
- change export types
- function plugin metadata and base class
- translation package types
- date/time and array helper types
- typed error classes

Concrete named exports from `src/index.ts` include:

- `HyperFormula`
- `ConfigParams`
- `AlwaysDense`
- `AlwaysSparse`
- `DenseSparseChooseBasedOnThreshold`
- `CellValue`
- `NoErrorCellValue`
- `RawCellContent`
- `FormatInfo`
- `Sheet`
- `Sheets`
- `SheetDimensions`
- `SimpleCellAddress`
- `SimpleCellRange`
- `ColumnRowIndex`
- `RawTranslationPackage`
- `FunctionPluginDefinition`
- `FunctionArgument`
- `NamedExpression`
- `NamedExpressionOptions`
- `CellType`
- `CellValueType`
- `CellValueDetailedType`
- `ErrorType`
- `ExportedChange`
- `ExportedCellChange`
- `ExportedNamedExpressionChange`
- `DetailedCellError`
- `CellError`
- `ArraySize`
- `FunctionPlugin`
- `ImplementedFunctions`
- `FunctionMetadata`
- `FunctionArgumentType`
- `SimpleRangeValue`
- `EmptyValue`
- `SerializedNamedExpression`
- `ConfigValueTooBigError`
- `ConfigValueTooSmallError`
- `EvaluationSuspendedError`
- `ExpectedOneOfValuesError`
- `ExpectedValueOfTypeError`
- `FunctionPluginValidationError`
- `InvalidAddressError`
- `InvalidArgumentsError`
- `LanguageAlreadyRegisteredError`
- `LanguageNotRegisteredError`
- `MissingTranslationError`
- `NamedExpressionDoesNotExistError`
- `NamedExpressionNameIsAlreadyTakenError`
- `NamedExpressionNameIsInvalidError`
- `NoOperationToRedoError`
- `NoOperationToUndoError`
- `NoRelativeAddressesAllowedError`
- `NoSheetWithIdError`
- `NoSheetWithNameError`
- `NotAFormulaError`
- `NothingToPasteError`
- `ProtectedFunctionTranslationError`
- `SheetNameAlreadyTakenError`
- `SheetSizeLimitExceededError`
- `SourceLocationHasArrayError`
- `TargetLocationHasArrayError`
- `UnableToParseError`

### Events

HyperFormula defines these semantic events:

- `sheetAdded`
- `sheetRemoved`
- `sheetRenamed`
- `namedExpressionAdded`
- `namedExpressionRemoved`
- `valuesUpdated`
- `evaluationSuspended`
- `evaluationResumed`

### Config keys

HyperFormula `ConfigParams` includes these keys:

- `accentSensitive`
- `caseSensitive`
- `caseFirst`
- `chooseAddressMappingPolicy`
- `context`
- `currencySymbol`
- `dateFormats`
- `functionArgSeparator`
- `decimalSeparator`
- `evaluateNullToZero`
- `functionPlugins`
- `ignorePunctuation`
- `language`
- `ignoreWhiteSpace`
- `leapYear1900`
- `licenseKey`
- `localeLang`
- `matchWholeCell`
- `arrayColumnSeparator`
- `arrayRowSeparator`
- `maxRows`
- `maxColumns`
- `nullDate`
- `nullYear`
- `parseDateTime`
- `precisionEpsilon`
- `precisionRounding`
- `stringifyDateTime`
- `stringifyDuration`
- `smartRounding`
- `thousandSeparator`
- `timeFormats`
- `useArrayArithmetic`
- `useColumnIndex`
- `useStats`
- `undoLimit`
- `useRegularExpressions`
- `useWildcards`

## Proposed `@bilig/headless` facade

### Public class

Expose one main class:

```ts
export class HeadlessWorkbook {}
```

This class intentionally adopts the HyperFormula-style method families directly.

### Relationship to `SpreadsheetEngine`

- `HeadlessWorkbook` owns one `SpreadsheetEngine`.
- `SpreadsheetEngine` remains public in `@bilig/core`.
- `HeadlessWorkbook` is the ergonomic headless API.
- `SpreadsheetEngine` is the low-level engine API.

`HeadlessWorkbook` must not expose the inner engine as a public field. That would make the facade unstable immediately.

## API mapping

The following tables list every relevant HyperFormula public method and the proposed `@bilig/headless` method.

Status values:

- `exact`: same method name and broad behavior
- `adapted`: same capability, behavior clarified for `bilig`
- `new`: `bilig` must add this capability on the facade

### Static factories and registries

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `version` | `HeadlessWorkbook.version` | `new` | derive from package version/build metadata |
| `buildDate` | `HeadlessWorkbook.buildDate` | `new` | facade build metadata |
| `releaseDate` | `HeadlessWorkbook.releaseDate` | `new` | facade build metadata |
| `languages` | `HeadlessWorkbook.languages` | `new` | registry owned by headless package |
| `defaultConfig` | `HeadlessWorkbook.defaultConfig` | `new` | map to `HeadlessConfig` defaults |
| `buildFromArray()` | `HeadlessWorkbook.buildFromArray()` | `new` | wrap engine construction and seed one sheet |
| `buildFromSheets()` | `HeadlessWorkbook.buildFromSheets()` | `new` | wrap engine construction and seed multiple sheets |
| `buildEmpty()` | `HeadlessWorkbook.buildEmpty()` | `new` | wrap empty engine construction |
| `getLanguage()` | `HeadlessWorkbook.getLanguage()` | `new` | headless-owned registry |
| `registerLanguage()` | `HeadlessWorkbook.registerLanguage()` | `new` | headless-owned registry |
| `unregisterLanguage()` | `HeadlessWorkbook.unregisterLanguage()` | `new` | headless-owned registry |
| `getRegisteredLanguagesCodes()` | `HeadlessWorkbook.getRegisteredLanguagesCodes()` | `new` | headless-owned registry |
| `registerFunctionPlugin()` | `HeadlessWorkbook.registerFunctionPlugin()` | `new` | static plugin catalog |
| `unregisterFunctionPlugin()` | `HeadlessWorkbook.unregisterFunctionPlugin()` | `new` | static plugin catalog |
| `registerFunction()` | `HeadlessWorkbook.registerFunction()` | `new` | static plugin catalog |
| `unregisterFunction()` | `HeadlessWorkbook.unregisterFunction()` | `new` | static plugin catalog |
| `unregisterAllFunctions()` | `HeadlessWorkbook.unregisterAllFunctions()` | `new` | static plugin catalog |
| `getRegisteredFunctionNames()` | `HeadlessWorkbook.getRegisteredFunctionNames()` | `new` | headless plugin/language registry |
| `getFunctionPlugin()` | `HeadlessWorkbook.getFunctionPlugin()` | `new` | headless plugin registry |
| `getAllFunctionPlugins()` | `HeadlessWorkbook.getAllFunctionPlugins()` | `new` | headless plugin registry |

### Read and serialization

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `getCellValue()` | `getCellValue()` | `adapted` | wraps `engine.getCellValue()` |
| `getCellFormula()` | `getCellFormula()` | `new` | read from cell snapshot / formula table |
| `getCellHyperlink()` | `getCellHyperlink()` | `new` | parse HYPERLINK formula serialization in JS path |
| `getCellSerialized()` | `getCellSerialized()` | `new` | return formula string or raw literal |
| `getRangeValues()` | `getRangeValues()` | `adapted` | wraps `engine.getRangeValues()` |
| `getRangeFormulas()` | `getRangeFormulas()` | `new` | dense rectangular formula read |
| `getRangeSerialized()` | `getRangeSerialized()` | `new` | dense rectangular raw-content read |
| `getSheetValues()` | `getSheetValues()` | `new` | serialize one sheet by id |
| `getSheetFormulas()` | `getSheetFormulas()` | `new` | serialize one sheet formulas by id |
| `getSheetSerialized()` | `getSheetSerialized()` | `new` | serialize one sheet raw content by id |
| `getAllSheetsValues()` | `getAllSheetsValues()` | `new` | all-sheet serialization |
| `getAllSheetsFormulas()` | `getAllSheetsFormulas()` | `new` | all-sheet serialization |
| `getAllSheetsSerialized()` | `getAllSheetsSerialized()` | `new` | all-sheet serialization |
| `getAllSheetsDimensions()` | `getAllSheetsDimensions()` | `new` | derive from workbook sheet grids |
| `getSheetDimensions()` | `getSheetDimensions()` | `new` | derive from one sheet grid |

### Config and runtime control

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `updateConfig()` | `updateConfig()` | `new` | facade-owned config, may rebuild parser/evaluator |
| `getConfig()` | `getConfig()` | `new` | facade-owned config |
| `rebuildAndRecalculate()` | `rebuildAndRecalculate()` | `new` | one-shot rebuild and recalc |
| `getStats()` | `getStats()` | `adapted` | map from engine metrics and future headless stats |
| `batch()` | `batch()` | `new` | defer event emission and coalesce changes |
| `suspendEvaluation()` | `suspendEvaluation()` | `new` | local headless evaluation gate |
| `resumeEvaluation()` | `resumeEvaluation()` | `new` | flush and recalc once |
| `isEvaluationSuspended()` | `isEvaluationSuspended()` | `new` | facade state |

### Undo, redo, and clipboard

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `undo()` | `undo()` | `adapted` | wraps engine undo and exports changes |
| `redo()` | `redo()` | `adapted` | wraps engine redo and exports changes |
| `isThereSomethingToUndo()` | `isThereSomethingToUndo()` | `new` | facade history inspection |
| `isThereSomethingToRedo()` | `isThereSomethingToRedo()` | `new` | facade history inspection |
| `copy()` | `copy()` | `new` | facade clipboard state |
| `cut()` | `cut()` | `new` | facade clipboard state + mutation |
| `paste()` | `paste()` | `new` | facade clipboard application |
| `isClipboardEmpty()` | `isClipboardEmpty()` | `new` | facade clipboard state |
| `clearClipboard()` | `clearClipboard()` | `new` | facade clipboard state |
| `clearRedoStack()` | `clearRedoStack()` | `new` | facade history |
| `clearUndoStack()` | `clearUndoStack()` | `new` | facade history |
| `getFillRangeData()` | `getFillRangeData()` | `new` | derive from existing range translation helpers |

### Mutation preflight

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `isItPossibleToSetCellContents()` | `isItPossibleToSetCellContents()` | `new` | validate address, size, collisions |
| `isItPossibleToSwapRowIndexes()` | `isItPossibleToSwapRowIndexes()` | `new` | validate row permutations |
| `isItPossibleToSetRowOrder()` | `isItPossibleToSetRowOrder()` | `new` | validate row permutations |
| `isItPossibleToSwapColumnIndexes()` | `isItPossibleToSwapColumnIndexes()` | `new` | validate column permutations |
| `isItPossibleToSetColumnOrder()` | `isItPossibleToSetColumnOrder()` | `new` | validate column permutations |
| `isItPossibleToAddRows()` | `isItPossibleToAddRows()` | `new` | validate bounds and shape constraints |
| `isItPossibleToRemoveRows()` | `isItPossibleToRemoveRows()` | `new` | validate bounds and dependency constraints |
| `isItPossibleToAddColumns()` | `isItPossibleToAddColumns()` | `new` | validate bounds and shape constraints |
| `isItPossibleToRemoveColumns()` | `isItPossibleToRemoveColumns()` | `new` | validate bounds and dependency constraints |
| `isItPossibleToMoveCells()` | `isItPossibleToMoveCells()` | `new` | validate source/target and spill collisions |
| `isItPossibleToMoveRows()` | `isItPossibleToMoveRows()` | `new` | validate row move |
| `isItPossibleToMoveColumns()` | `isItPossibleToMoveColumns()` | `new` | validate column move |
| `isItPossibleToAddSheet()` | `isItPossibleToAddSheet()` | `new` | validate name availability |
| `isItPossibleToRemoveSheet()` | `isItPossibleToRemoveSheet()` | `new` | validate sheet existence |
| `isItPossibleToClearSheet()` | `isItPossibleToClearSheet()` | `new` | validate sheet existence |
| `isItPossibleToReplaceSheetContent()` | `isItPossibleToReplaceSheetContent()` | `new` | validate content bounds |
| `isItPossibleToRenameSheet()` | `isItPossibleToRenameSheet()` | `new` | validate name availability |
| `isItPossibleToAddNamedExpression()` | `isItPossibleToAddNamedExpression()` | `new` | validate name, scope, relative refs |
| `isItPossibleToChangeNamedExpression()` | `isItPossibleToChangeNamedExpression()` | `new` | validate name, scope, relative refs |
| `isItPossibleToRemoveNamedExpression()` | `isItPossibleToRemoveNamedExpression()` | `new` | validate existence |

### Cell, row, column, and sheet mutation

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `setCellContents()` | `setCellContents()` | `new` | wraps `setCellValue`, `setCellFormula`, range writes |
| `swapRowIndexes()` | `swapRowIndexes()` | `new` | headless-level permutation helper |
| `setRowOrder()` | `setRowOrder()` | `new` | headless-level permutation helper |
| `swapColumnIndexes()` | `swapColumnIndexes()` | `new` | headless-level permutation helper |
| `setColumnOrder()` | `setColumnOrder()` | `new` | headless-level permutation helper |
| `addRows()` | `addRows()` | `adapted` | wraps `engine.insertRows()` |
| `removeRows()` | `removeRows()` | `adapted` | wraps `engine.deleteRows()` |
| `addColumns()` | `addColumns()` | `adapted` | wraps `engine.insertColumns()` |
| `removeColumns()` | `removeColumns()` | `adapted` | wraps `engine.deleteColumns()` |
| `moveCells()` | `moveCells()` | `adapted` | wraps `engine.moveRange()` |
| `moveRows()` | `moveRows()` | `adapted` | wraps `engine.moveRows()` |
| `moveColumns()` | `moveColumns()` | `adapted` | wraps `engine.moveColumns()` |
| `addSheet()` | `addSheet()` | `adapted` | wraps `engine.createSheet()` and returns name |
| `removeSheet()` | `removeSheet()` | `adapted` | wraps `engine.deleteSheet()` |
| `clearSheet()` | `clearSheet()` | `new` | clear all cells on a sheet via batch mutation |
| `setSheetContent()` | `setSheetContent()` | `new` | replace one sheet content from 2D array |
| `renameSheet()` | `renameSheet()` | `adapted` | wraps `engine.renameSheet()` |

### Address conversion, graph, identity, and introspection

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `simpleCellAddressFromString()` | `simpleCellAddressFromString()` | `new` | wrap formula address parsing + sheet id mapping |
| `simpleCellRangeFromString()` | `simpleCellRangeFromString()` | `new` | wrap parser helpers |
| `simpleCellRangeToString()` | `simpleCellRangeToString()` | `new` | formatter helper |
| `getCellDependents()` | `getCellDependents()` | `adapted` | wraps `engine.getDependents()` |
| `getCellPrecedents()` | `getCellPrecedents()` | `adapted` | wraps `engine.getDependencies()` |
| `getSheetName()` | `getSheetName()` | `adapted` | wraps stable sheet-id lookup |
| `getSheetNames()` | `getSheetNames()` | `new` | derive from workbook store |
| `getSheetId()` | `getSheetId()` | `adapted` | wraps stable sheet-id lookup |
| `doesSheetExist()` | `doesSheetExist()` | `adapted` | wraps workbook store |
| `countSheets()` | `countSheets()` | `new` | derive from workbook store |
| `getCellType()` | `getCellType()` | `new` | derive from cell snapshot |
| `doesCellHaveSimpleValue()` | `doesCellHaveSimpleValue()` | `new` | derive from cell snapshot |
| `doesCellHaveFormula()` | `doesCellHaveFormula()` | `new` | derive from cell snapshot |
| `isCellEmpty()` | `isCellEmpty()` | `new` | derive from cell snapshot |
| `isCellPartOfArray()` | `isCellPartOfArray()` | `new` | derive from spill metadata / snapshot |
| `getCellValueType()` | `getCellValueType()` | `new` | derive from cell value tag |
| `getCellValueDetailedType()` | `getCellValueDetailedType()` | `new` | derive from value + formatting/date subtype |
| `getCellValueFormat()` | `getCellValueFormat()` | `new` | derive from number-format/style metadata |

### Named expressions

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `addNamedExpression()` | `addNamedExpression()` | `adapted` | map to `setDefinedName()` with scope support |
| `getNamedExpressionValue()` | `getNamedExpressionValue()` | `new` | evaluate named-expression value |
| `getNamedExpressionFormula()` | `getNamedExpressionFormula()` | `new` | serialize named-expression formula |
| `getNamedExpression()` | `getNamedExpression()` | `new` | serialize one name record |
| `changeNamedExpression()` | `changeNamedExpression()` | `adapted` | overwrite named expression |
| `removeNamedExpression()` | `removeNamedExpression()` | `adapted` | delete named expression |
| `listNamedExpressions()` | `listNamedExpressions()` | `new` | list names in scope |
| `getAllNamedExpressionsSerialized()` | `getAllNamedExpressionsSerialized()` | `new` | serialize names in stable sheet-id scope |

### Formula utilities, function introspection, date helpers, lifecycle

| HyperFormula | `@bilig/headless` | Status | Basis / note |
| --- | --- | --- | --- |
| `normalizeFormula()` | `normalizeFormula()` | `adapted` | parse + serialize with `@bilig/formula` |
| `calculateFormula()` | `calculateFormula()` | `new` | evaluate one formula in workbook context |
| `getNamedExpressionsFromFormula()` | `getNamedExpressionsFromFormula()` | `new` | collect name refs from AST |
| `validateFormula()` | `validateFormula()` | `adapted` | parse-only validation |
| `getRegisteredFunctionNames()` | `getRegisteredFunctionNames()` | `new` | instance view of static registry |
| `getFunctionPlugin()` | `getFunctionPlugin()` | `new` | instance view of static registry |
| `getAllFunctionPlugins()` | `getAllFunctionPlugins()` | `new` | instance view of static registry |
| `numberToDateTime()` | `numberToDateTime()` | `new` | facade date/time helper using config |
| `numberToDate()` | `numberToDate()` | `new` | facade date helper using config |
| `numberToTime()` | `numberToTime()` | `new` | facade time helper using config |
| `destroy()` | `destroy()` | `new` | alias of `dispose()` |

## Additional public types

### Addresses and ranges

```ts
export interface HeadlessCellAddress {
  sheet: number;
  col: number;
  row: number;
}

export interface HeadlessCellRange {
  start: HeadlessCellAddress;
  end: HeadlessCellAddress;
}
```

### Changes

Keep HyperFormula’s operational shape: mutating methods return visible value changes only.

```ts
export type HeadlessChange =
  | HeadlessCellChange
  | HeadlessNamedExpressionChange;

export interface HeadlessCellChange {
  kind: "cell";
  address: HeadlessCellAddress;
  sheetName: string;
  a1: string;
  newValue: CellValue;
}

export interface HeadlessNamedExpressionChange {
  kind: "named-expression";
  name: string;
  scope?: number;
  newValue: CellValue | CellValue[][];
}
```

### Named expressions

```ts
export interface HeadlessNamedExpression {
  name: string;
  scope?: number;
  expression?: RawCellContent;
  options?: Record<string, string | number | boolean>;
}

export interface SerializedHeadlessNamedExpression {
  name: string;
  expression: RawCellContent;
  scope?: number;
  options?: Record<string, string | number | boolean>;
}
```

### Config

Expose one `HeadlessConfig` with the full HyperFormula-equivalent key set. The facade may internally ignore a key until its implementation phase lands, but the key names must be stable from the start.

```ts
export interface HeadlessConfig {
  accentSensitive?: boolean;
  caseSensitive?: boolean;
  caseFirst?: "upper" | "lower" | "false";
  chooseAddressMappingPolicy?: unknown;
  context?: unknown;
  currencySymbol?: string[];
  dateFormats?: string[];
  functionArgSeparator?: string;
  decimalSeparator?: "." | ",";
  evaluateNullToZero?: boolean;
  functionPlugins?: HeadlessFunctionPlugin[];
  ignorePunctuation?: boolean;
  language?: string;
  ignoreWhiteSpace?: "standard" | "any";
  leapYear1900?: boolean;
  licenseKey?: string;
  localeLang?: string;
  matchWholeCell?: boolean;
  arrayColumnSeparator?: "," | ";";
  arrayRowSeparator?: ";" | "|";
  maxRows?: number;
  maxColumns?: number;
  nullDate?: { year: number; month: number; day: number };
  nullYear?: number;
  parseDateTime?: (input: string) => unknown;
  precisionEpsilon?: number;
  precisionRounding?: number;
  stringifyDateTime?: (value: unknown) => string | undefined;
  stringifyDuration?: (value: unknown) => string | undefined;
  smartRounding?: boolean;
  thousandSeparator?: "" | "," | ".";
  timeFormats?: string[];
  useArrayArithmetic?: boolean;
  useColumnIndex?: boolean;
  useStats?: boolean;
  undoLimit?: number;
  useRegularExpressions?: boolean;
  useWildcards?: boolean;
}
```

Compatibility note:

- `licenseKey` exists only so the config surface can cover HyperFormula-style callers. It is a no-op in `bilig`.
- `chooseAddressMappingPolicy` is reserved for future storage-policy work. It is ignored in the first headless release.

### Events

```ts
export interface HeadlessWorkbookEventMap {
  sheetAdded: { sheetId: number; sheetName: string };
  sheetRemoved: { sheetId: number; sheetName: string; changes: HeadlessChange[] };
  sheetRenamed: { sheetId: number; oldName: string; newName: string };
  namedExpressionAdded: { name: string; scope?: number; changes: HeadlessChange[] };
  namedExpressionRemoved: { name: string; scope?: number; changes: HeadlessChange[] };
  valuesUpdated: { changes: HeadlessChange[] };
  evaluationSuspended: {};
  evaluationResumed: { changes: HeadlessChange[] };
}
```

### Custom function interfaces

```ts
export type HeadlessFunctionArgumentType =
  | "STRING"
  | "NUMBER"
  | "BOOLEAN"
  | "SCALAR"
  | "NOERROR"
  | "RANGE"
  | "INTEGER"
  | "COMPLEX"
  | "ANY";

export interface HeadlessFunctionArgument {
  argumentType: HeadlessFunctionArgumentType;
  passSubtype?: boolean;
  defaultValue?: unknown;
  optionalArg?: boolean;
  minValue?: number;
  maxValue?: number;
  lessThan?: number;
  greaterThan?: number;
}

export interface HeadlessFunctionMetadata {
  method: string;
  parameters?: HeadlessFunctionArgument[];
  repeatLastArgs?: number;
  expandRanges?: boolean;
  returnNumberType?: string;
  sizeOfResultArrayMethod?: string;
  isVolatile?: boolean;
  isDependentOnSheetStructureChange?: boolean;
  doesNotNeedArgumentsToBeComputed?: boolean;
  enableArrayArithmeticForArguments?: boolean;
  vectorizationForbidden?: boolean;
}

export interface HeadlessFunctionPlugin {
  implementedFunctions: Record<string, HeadlessFunctionMetadata>;
  aliases?: Record<string, string>;
}
```

## Behavior contracts

This section is the part the previous version lacked. These contracts are required before implementation.

### 1. Sheet identity contract

- Public headless APIs use numeric `sheetId` in address objects, matching HyperFormula.
- `sheetId` is stable across rename, snapshot export/import, and rebuild.
- This is supported by current `WorkbookStore` and tested in `packages/core/src/__tests__/sheet-ids.test.ts`.
- Name-based helpers remain available: `getSheetId(name)`, `getSheetName(id)`, `doesSheetExist(name)`.

### 2. Mutation return contract

Every mutating headless method returns `HeadlessChange[]`.

That array means:

- only externally visible value changes
- no metadata deltas for styles, filters, sorts, tables, pivots, freeze panes, or formats
- one `HeadlessCellChange` per changed visible cell
- one `HeadlessNamedExpressionChange` per changed named-expression value

Determinism rules:

- duplicate addresses are coalesced, final value wins
- spill/array outputs are expanded to one cell change per visible cell
- row/column/sheet structural edits only report value recalculation fallout, not tombstones
- metadata-only operations return `[]`
- ordering is stable:
  - cell changes first, sorted by sheet order, then row, then column
  - named-expression changes second, sorted by scope then name

### 3. Event contract

Events are synchronous.

For a mutating public method:

1. preflight runs
2. mutation applies
3. recalculation finishes
4. specific semantic event fires if relevant
5. `valuesUpdated` fires if and only if returned `HeadlessChange[]` is non-empty
6. method returns the same `HeadlessChange[]`

Examples:

- `addSheet()` emits `sheetAdded`; it emits `valuesUpdated` only if visible values changed
- `removeSheet()` emits `sheetRemoved(changes)` and then `valuesUpdated(changes)` if `changes.length > 0`
- `addNamedExpression()` emits `namedExpressionAdded(changes)` and then `valuesUpdated(changes)` if `changes.length > 0`
- `suspendEvaluation()` emits only `evaluationSuspended`
- `resumeEvaluation()` emits `evaluationResumed(changes)` and then `valuesUpdated(changes)` if `changes.length > 0`

### 4. Batch contract

`batch(fn)` is not transactional.

Rules:

- all operations inside `fn` are applied immediately to engine state
- recalculation may be deferred internally, but state mutations are not rolled back on throw
- semantic events are suppressed until the outermost batch finishes
- returned changes are the coalesced final visible changes across the whole batch
- nested batches are allowed
- one completed batch becomes one undo entry

### 5. Evaluation suspension contract

`suspendEvaluation()` applies only to the headless facade’s local mutation methods.

Rules:

- local mutating methods can still be called while suspended
- visible value getters throw `HeadlessEvaluationSuspendedError` while suspended
- preflight methods still work while suspended
- formula parse/validation helpers still work while suspended
- `resumeEvaluation()` performs one recalculation pass and one event flush

Initial scope rule:

- sync and remote-batch APIs are not part of `@bilig/headless` v1
- this avoids ambiguous interaction between suspension and remote mutation

### 6. Preflight contract

`isItPossibleTo*` methods:

- are pure
- emit no events
- mutate nothing
- return `false` for legal-but-impossible operations
- throw for malformed arguments

Examples of `false`:

- sheet name already taken
- moving cells into a blocked spill range
- adding rows beyond configured max rows
- removing a missing named expression

Examples of thrown argument errors:

- negative row count
- NaN sheet id
- malformed address object
- invalid 2D content shape

### 7. Undo/redo contract

- each public mutating call creates one undo entry
- `batch()` creates one undo entry
- `resumeEvaluation()` does not create a new undo entry by itself
- `undo()` and `redo()` return visible final changes, not raw ops
- `clearUndoStack()` and `clearRedoStack()` affect headless history only

### 8. Clipboard contract

Clipboard state is per workbook instance, not global.

Rules:

- `copy(range)` captures serialized cell content and translated formulas
- `cut(range)` captures clipboard data and clears source as one undoable operation
- `paste(target)` applies the last clipboard payload translated to the target anchor
- `isClipboardEmpty()` reports workbook-local clipboard state
- `getFillRangeData()` is pure and does not mutate

### 9. Named-expression contract

The headless facade exposes HyperFormula-style named expressions.

Rules:

- workbook scope: `scope === undefined`
- sheet scope: `scope === sheetId`
- names are case-insensitive for lookup, casing-preserving for serialization
- relative references in named expressions are rejected
- returned serialized scopes use stable `sheetId`, not sheet name

Mapping to current `bilig` internals:

- headless named expressions sit on top of defined names
- workbook scope maps directly
- sheet scope requires explicit facade-owned scoping support above core metadata

### 10. Custom function contract

This is a public contract, not just an adapter hook.

Rules:

- plugins are registered statically on `HeadlessWorkbook`
- each workbook captures a snapshot of the plugin registry at construction time
- later static registration changes do not affect existing workbooks
- workbook config may narrow the enabled plugin set with `functionPlugins`
- custom functions execute in JS only in v1
- `context` from `HeadlessConfig` is passed through to custom function execution
- volatility and structure-dependence flags are honored by recalc scheduling
- array-returning custom functions must declare `sizeOfResultArrayMethod`

Explicitly out of scope for the custom-function surface:

- reusing the current global `external-function-adapter` API as the primary public custom-function API
- WASM custom-function execution
- async custom functions

### 11. Config update contract

`updateConfig(next)` must preserve:

- workbook data
- sheet ids
- named expressions
- undo/redo history

It may rebuild:

- parser
- serializer
- coercion/date helpers
- custom-function registry view

If a config update changes formula syntax settings, formulas remain stored canonically and are reinterpreted under the new config.

### 12. Error contract

Public headless APIs throw typed library errors. They do not expose `effect` errors directly.

Recommended public classes:

- `HeadlessArgumentError`
- `HeadlessConfigError`
- `HeadlessSheetError`
- `HeadlessNamedExpressionError`
- `HeadlessClipboardError`
- `HeadlessEvaluationSuspendedError`
- `HeadlessParseError`
- `HeadlessOperationError`

Rules:

- formula/data errors such as `#REF!` remain returned as cell values, not thrown
- invalid API arguments throw typed library errors
- impossible operations return `false` from `isItPossibleTo*`
- execution failures in mutating methods throw `HeadlessOperationError`
- internal `Engine*Error` types are mapped into the public headless error classes

## What stays in `@bilig/core` for now

These capabilities are real and valuable, but they are not part of the first headless-parity target:

- `connectSyncClient()`
- `disconnectSyncClient()`
- `applyRemoteBatch()`
- `exportReplicaSnapshot()`
- `importReplicaSnapshot()`
- table/filter/sort/pivot/freeze-pane mutation APIs
- style and number-format mutation APIs

Reason:

- they are `bilig`-specific strengths
- they do not belong to the minimum HyperFormula-equivalent headless workflow
- keeping them out of the initial facade keeps the first public surface coherent

They remain available in `@bilig/core` and can be surfaced later under clearly named `bilig`-native extensions.

## Implementation plan

### Phase 1: construction, ids, and read surface

- create `@bilig/headless`
- implement factories
- implement stable `sheetId` mapping
- implement address conversion helpers
- implement read/serialization methods
- implement `dispose()` / `destroy()`

### Phase 2: mutation return values and events

- add change exporter
- wrap all sheet/cell/range/row/column mutations
- add synchronous semantic event emitter
- add deterministic event ordering

### Phase 3: undo/redo, clipboard, batch, suspend/resume

- expose undo/redo stack status
- add clipboard session
- add `batch()`
- add `suspendEvaluation()` / `resumeEvaluation()`

### Phase 4: named expressions and formula utilities

- add scoped named-expression model
- add name serialization
- add formula normalization, validation, and one-off evaluation helpers

### Phase 5: custom functions and config updates

- add plugin registry
- add function metadata validation
- route custom functions through JS evaluator
- add `updateConfig()`, `rebuildAndRecalculate()`, and stats

### Phase 6: language packs

- add language registry
- add translated function-name lookup
- add localized parser/serializer support

## Acceptance criteria

- a user can build a workbook from a 2D array or named sheets with one package
- the headless facade exposes every HyperFormula headless workflow method family
- every mutating method returns deterministic `HeadlessChange[]`
- event emission order is specified and tested
- stable numeric sheet ids survive rename and snapshot roundtrip
- `batch()` and evaluation suspension semantics are specified and tested
- named expressions support workbook and sheet scope
- JS custom functions can be registered with metadata
- no public headless method leaks `effect` error types

## Final recommendation

Proceed with `@bilig/headless` as a first-class library package.

Do not:

- overload `SpreadsheetEngine` with every HyperFormula-shaped concern
- build a separate compatibility package
- mix sync/replica APIs into the first headless facade

Build the headless facade around exact behavior contracts first. Without the contracts in this document, the implementation will look complete while remaining semantically unstable.
