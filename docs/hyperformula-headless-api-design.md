# HyperFormula Headless API Design

> Superseded by `docs/workpaper-platform-design.md`. The original `HeadlessWorkbook`
> parity design remains useful as historical context, but `WorkPaper` is now the
> canonical headless API contract.

This document defines the `bilig` headless workbook API target by reading the local HyperFormula checkout at `/Users/gregkonush/github.com/hyperformula` and translating that public surface into a `bilig` design that fits this repo.

Date: `2026-04-09`

## Source Corpus

Reviewed HyperFormula sources:

- `/Users/gregkonush/github.com/hyperformula/src/index.ts`
- `/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Emitter.ts`
- `/Users/gregkonush/github.com/hyperformula/src/ConfigParams.ts`
- `/Users/gregkonush/github.com/hyperformula/src/errors.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/FunctionRegistry.ts`
- `/Users/gregkonush/github.com/hyperformula/src/NamedExpressions.ts`
- `/Users/gregkonush/github.com/hyperformula/README.md`

Relevant `bilig` sources:

- `packages/core/src/engine.ts`
- `packages/formula/src/*`
- `docs/public-api.md`
- `docs/hyperformula-lessons-for-bilig.md`

## Goal

Ship a stable `@bilig/headless` package that can be used the way HyperFormula is used in server-side and non-UI spreadsheet workflows:

- construct a workbook from arrays or named sheets
- read cell, range, sheet, and named-expression values
- mutate workbook structure semantically
- use undo/redo, clipboard, and batching
- register languages and custom functions
- subscribe to workbook lifecycle and recalculation events
- inspect dependency relationships

The design target is functional parity for headless workflows, not a byte-for-byte clone of HyperFormula internals.

## Non-goals

- Copying HyperFormula implementation details or GPL-licensed code
- Reproducing HyperFormula's exact internal object model
- Moving UI, sync, or network APIs into `@bilig/headless`
- Freezing `@bilig/core` internals just to mirror HyperFormula raw getters

## Proposed Package Boundary

Add a new public package:

- `@bilig/headless`

Primary public class:

- `HeadlessWorkbook`

Supporting exports:

- cell, range, sheet, change, config, plugin, event, and named-expression types
- typed error classes
- address-mapping policy helpers once `@bilig/core` exposes them as public configuration primitives

The package is a facade on top of `@bilig/core` `SpreadsheetEngine` plus `@bilig/formula` parsing, formatting, and translation helpers.

## HyperFormula Entry Point Inventory

HyperFormula's package entry point exports these categories today:

- core class: `HyperFormula`
- workbook value and address types: `CellValue`, `RawCellContent`, `SimpleCellAddress`, `SimpleCellRange`, `Sheet`, `Sheets`, `SheetDimensions`
- mutation/change types: `ExportedChange`, `ExportedCellChange`, `ExportedNamedExpressionChange`, `ColumnRowIndex`
- formula and plugin types: `FunctionPluginDefinition`, `FunctionPlugin`, `FunctionMetadata`, `FunctionArgument`, `FunctionArgumentType`, `ImplementedFunctions`
- named-expression types: `NamedExpression`, `NamedExpressionOptions`, `SerializedNamedExpression`
- config and i18n types: `ConfigParams`, `RawTranslationPackage`
- value classification enums: `CellType`, `CellValueType`, `CellValueDetailedType`, `ErrorType`
- engine helper classes: `ArraySize`, `SimpleRangeValue`, `EmptyValue`
- address-mapping policy classes: `AlwaysDense`, `AlwaysSparse`, `DenseSparseChooseBasedOnThreshold`
- exported error classes from `src/errors.ts`

`bilig` should expose an equivalent category layout, even where names differ.

## HyperFormula Public Class Surface

HyperFormula's public class surface in `src/HyperFormula.ts` breaks down into these groups.

### Static properties and factories

| HyperFormula | `bilig` target | Notes |
| --- | --- | --- |
| `version` | `HeadlessWorkbook.version` | same role |
| `buildDate` | `HeadlessWorkbook.buildDate` | same role |
| `releaseDate` | `HeadlessWorkbook.releaseDate` | same role |
| `languages` | `HeadlessWorkbook.languages` | public registry snapshot |
| `defaultConfig` | `HeadlessWorkbook.defaultConfig` | getter or readonly clone |
| `buildEmpty()` | `HeadlessWorkbook.buildEmpty()` | same |
| `buildFromArray()` | `HeadlessWorkbook.buildFromArray()` | same |
| `buildFromSheets()` | `HeadlessWorkbook.buildFromSheets()` | same |

### Static localization and function registration

| HyperFormula | `bilig` target | Notes |
| --- | --- | --- |
| `getLanguage()` | `HeadlessWorkbook.getLanguage()` | same |
| `registerLanguage()` | `HeadlessWorkbook.registerLanguage()` | same |
| `unregisterLanguage()` | `HeadlessWorkbook.unregisterLanguage()` | same |
| `getRegisteredLanguagesCodes()` | `HeadlessWorkbook.getRegisteredLanguagesCodes()` | same |
| `registerFunctionPlugin(plugin, translations?)` | `HeadlessWorkbook.registerFunctionPlugin(plugin, translations?)` | translations supported |
| `unregisterFunctionPlugin(plugin)` | `HeadlessWorkbook.unregisterFunctionPlugin(plugin)` | same |
| `registerFunction(functionId, plugin, translations?)` | `HeadlessWorkbook.registerFunction(functionId, plugin, translations?)` | same |
| `unregisterFunction(functionId)` | `HeadlessWorkbook.unregisterFunction(functionId)` | same |
| `unregisterAllFunctions()` | `HeadlessWorkbook.unregisterAllFunctions()` | same |
| `getRegisteredFunctionNames(code)` | `HeadlessWorkbook.getRegisteredFunctionNames(code)` | same |
| `getFunctionPlugin(functionId)` | `HeadlessWorkbook.getFunctionPlugin(functionId)` | same |
| `getAllFunctionPlugins()` | `HeadlessWorkbook.getAllFunctionPlugins()` | same |

### Internal and debug accessors exposed publicly by HyperFormula

HyperFormula exposes these getters:

- `graph`
- `rangeMapping`
- `arrayMapping`
- `sheetMapping`
- `addressMapping`
- `dependencyGraph`
- `evaluator`
- `columnSearch`
- `lazilyTransformingAstService`
- `licenseKeyValidityState`

Design decision for `bilig`:

- `licenseKeyValidityState` belongs on the stable public surface.
- The remaining nine getters should not expose raw `@bilig/core` internals directly.
- Instead, `@bilig/headless` should expose a stable `internals` namespace with read-only adapter objects that cover the specific observable behaviors users rely on.
- The getter names should still exist for migration-sensitive users, but they should return adapter facades, not raw runtime singletons.

Proposed adapter getters:

- `get graph()`
- `get rangeMapping()`
- `get arrayMapping()`
- `get sheetMapping()`
- `get addressMapping()`
- `get dependencyGraph()`
- `get evaluator()`
- `get columnSearch()`
- `get lazilyTransformingAstService()`

These adapters are phase-two work because `@bilig/core` does not yet export stable equivalents.

## Workbook Read Surface

`bilig` should support the full HyperFormula headless read set.

| HyperFormula | `bilig` target | Required backing |
| --- | --- | --- |
| `getCellValue()` | `getCellValue()` | `SpreadsheetEngine.getCellValue()` |
| `getCellFormula()` | `getCellFormula()` | `SpreadsheetEngine.getCell()` + formula restore |
| `getCellHyperlink()` | `getCellHyperlink()` | formula parse helper |
| `getCellSerialized()` | `getCellSerialized()` | cell snapshot serialization |
| `getRangeValues()` | `getRangeValues()` | `SpreadsheetEngine.getRangeValues()` |
| `getRangeFormulas()` | `getRangeFormulas()` | dense range helper |
| `getRangeSerialized()` | `getRangeSerialized()` | dense range helper |
| `getSheetValues()` | `getSheetValues()` | range expansion |
| `getSheetFormulas()` | `getSheetFormulas()` | range expansion |
| `getSheetSerialized()` | `getSheetSerialized()` | range expansion |
| `getAllSheetsDimensions()` | `getAllSheetsDimensions()` | sheet iteration |
| `getSheetDimensions()` | `getSheetDimensions()` | workbook sheet grid introspection |
| `getAllSheetsValues()` | `getAllSheetsValues()` | sheet iteration |
| `getAllSheetsFormulas()` | `getAllSheetsFormulas()` | sheet iteration |
| `getAllSheetsSerialized()` | `getAllSheetsSerialized()` | sheet iteration |
| `getCellDependents()` | `getCellDependents()` | `SpreadsheetEngine.getDependents()` |
| `getCellPrecedents()` | `getCellPrecedents()` | `SpreadsheetEngine.getDependencies()` |
| `getSheetName()` | `getSheetName()` | return `undefined` on miss |
| `getSheetNames()` | `getSheetNames()` | workbook sheet enumeration |
| `getSheetId()` | `getSheetId()` | return `undefined` on miss |
| `doesSheetExist()` | `doesSheetExist()` | workbook sheet lookup |
| `countSheets()` | `countSheets()` | workbook sheet count |
| `getCellType()` | `getCellType()` | cell snapshot + spill metadata |
| `doesCellHaveSimpleValue()` | `doesCellHaveSimpleValue()` | cell snapshot |
| `doesCellHaveFormula()` | `doesCellHaveFormula()` | cell snapshot |
| `isCellEmpty()` | `isCellEmpty()` | cell value |
| `isCellPartOfArray()` | `isCellPartOfArray()` | spill metadata |
| `getCellValueType()` | `getCellValueType()` | value tag mapping |
| `getCellValueDetailedType()` | `getCellValueDetailedType()` | format-aware classification |
| `getCellValueFormat()` | `getCellValueFormat()` | cell snapshot format |

Return semantics:

- lookup and parse helpers return `undefined` on misses or invalid input where HyperFormula does
- read methods that target a valid cell or range still throw for structurally invalid sheet IDs and suspended evaluation

## Workbook Mutation Surface

| HyperFormula | `bilig` target | Required backing |
| --- | --- | --- |
| `updateConfig()` | `updateConfig()` | rebuild engine preserving serialized workbook |
| `rebuildAndRecalculate()` | `rebuildAndRecalculate()` | config rebuild with current config |
| `undo()` | `undo()` | `SpreadsheetEngine.undo()` + change diff |
| `redo()` | `redo()` | `SpreadsheetEngine.redo()` + change diff |
| `isThereSomethingToUndo()` | `isThereSomethingToUndo()` | history stack |
| `isThereSomethingToRedo()` | `isThereSomethingToRedo()` | history stack |
| `isItPossibleToSetCellContents()` | `isItPossibleToSetCellContents()` | range and bounds validation |
| `setCellContents()` | `setCellContents()` | `setCellValue`, `setCellFormula`, matrix apply |
| `swapRowIndexes()` | `swapRowIndexes()` | tuple mapping form |
| `isItPossibleToSwapRowIndexes()` | `isItPossibleToSwapRowIndexes()` | tuple mapping validation |
| `setRowOrder()` | `setRowOrder()` | reorder via moves |
| `isItPossibleToSetRowOrder()` | `isItPossibleToSetRowOrder()` | validation |
| `swapColumnIndexes()` | `swapColumnIndexes()` | tuple mapping form |
| `isItPossibleToSwapColumnIndexes()` | `isItPossibleToSwapColumnIndexes()` | validation |
| `setColumnOrder()` | `setColumnOrder()` | reorder via moves |
| `isItPossibleToSetColumnOrder()` | `isItPossibleToSetColumnOrder()` | validation |
| `isItPossibleToAddRows()` | `isItPossibleToAddRows()` | variadic interval validation |
| `addRows()` | `addRows()` | `insertRows` |
| `isItPossibleToRemoveRows()` | `isItPossibleToRemoveRows()` | validation |
| `removeRows()` | `removeRows()` | `deleteRows` |
| `isItPossibleToAddColumns()` | `isItPossibleToAddColumns()` | validation |
| `addColumns()` | `addColumns()` | `insertColumns` |
| `isItPossibleToRemoveColumns()` | `isItPossibleToRemoveColumns()` | validation |
| `removeColumns()` | `removeColumns()` | `deleteColumns` |
| `isItPossibleToMoveCells()` | `isItPossibleToMoveCells()` | range validation |
| `moveCells()` | `moveCells()` | `moveRange` |
| `isItPossibleToMoveRows()` | `isItPossibleToMoveRows()` | validation |
| `moveRows()` | `moveRows()` | `moveRows` |
| `isItPossibleToMoveColumns()` | `isItPossibleToMoveColumns()` | validation |
| `moveColumns()` | `moveColumns()` | `moveColumns` |
| `isItPossibleToAddSheet()` | `isItPossibleToAddSheet()` | name validation |
| `addSheet()` | `addSheet()` | `createSheet` |
| `isItPossibleToRemoveSheet()` | `isItPossibleToRemoveSheet()` | validation |
| `removeSheet()` | `removeSheet()` | `deleteSheet` |
| `isItPossibleToClearSheet()` | `isItPossibleToClearSheet()` | validation |
| `clearSheet()` | `clearSheet()` | `clearRange` |
| `isItPossibleToReplaceSheetContent()` | `isItPossibleToReplaceSheetContent()` | validation |
| `setSheetContent()` | `setSheetContent()` | sheet clear + matrix apply |
| `isItPossibleToRenameSheet()` | `isItPossibleToRenameSheet()` | validation |
| `renameSheet()` | `renameSheet()` | `renameSheet` |
| `batch()` | `batch()` | grouped change diff and history merge |
| `suspendEvaluation()` | `suspendEvaluation()` | readable-state suspension |
| `resumeEvaluation()` | `resumeEvaluation()` | change flush |
| `isEvaluationSuspended()` | `isEvaluationSuspended()` | state flag |

Change-return contract:

- all mutating methods return `HeadlessChange[]`
- the array is sorted by sheet order, then row, then column, then named-expression updates
- when HyperFormula would emit `valuesUpdated`, `@bilig/headless` returns the same change set and emits the same recalculation boundary

## Clipboard and Fill Surface

| HyperFormula | `bilig` target | Notes |
| --- | --- | --- |
| `copy()` | `copy()` | returns copied values and stores translated serialized content |
| `cut()` | `cut()` | `copy()` plus clear |
| `paste()` | `paste()` | uses stored serialized content |
| `isClipboardEmpty()` | `isClipboardEmpty()` | same |
| `clearClipboard()` | `clearClipboard()` | same |
| `clearRedoStack()` | `clearRedoStack()` | same |
| `clearUndoStack()` | `clearUndoStack()` | same |
| `getFillRangeData(source, target, offsetsFromTarget)` | `getFillRangeData(source, target, offsetsFromTarget)` | exact parameter list |

## Address and Formula Helper Surface

| HyperFormula | `bilig` target | Notes |
| --- | --- | --- |
| `simpleCellAddressFromString()` | `simpleCellAddressFromString()` | `undefined` on invalid input |
| `simpleCellRangeFromString()` | `simpleCellRangeFromString()` | `undefined` on invalid input |
| `simpleCellAddressToString()` | `simpleCellAddressToString()` | include-sheet and context-sheet modes |
| `simpleCellRangeToString()` | `simpleCellRangeToString()` | include-sheet and context-sheet modes |
| `normalizeFormula()` | `normalizeFormula()` | formula parser + serializer |
| `calculateFormula()` | `calculateFormula()` | scratch-sheet evaluation path |
| `getNamedExpressionsFromFormula()` | `getNamedExpressionsFromFormula()` | AST walk |
| `validateFormula()` | `validateFormula()` | parser validation |
| `numberToDateTime()` | `numberToDateTime()` | date serial helper |
| `numberToDate()` | `numberToDate()` | date helper |
| `numberToTime()` | `numberToTime()` | time helper |

## Named Expressions Surface

| HyperFormula | `bilig` target | Notes |
| --- | --- | --- |
| `isItPossibleToAddNamedExpression()` | `isItPossibleToAddNamedExpression()` | validation |
| `addNamedExpression()` | `addNamedExpression()` | same |
| `getNamedExpressionValue()` | `getNamedExpressionValue()` | `undefined` on miss |
| `getNamedExpressionFormula()` | `getNamedExpressionFormula()` | `undefined` on miss or literal |
| `getNamedExpression()` | `getNamedExpression()` | `undefined` on miss |
| `isItPossibleToChangeNamedExpression()` | `isItPossibleToChangeNamedExpression()` | validation |
| `changeNamedExpression()` | `changeNamedExpression()` | same |
| `isItPossibleToRemoveNamedExpression()` | `isItPossibleToRemoveNamedExpression()` | validation |
| `removeNamedExpression()` | `removeNamedExpression()` | same |
| `listNamedExpressions()` | `listNamedExpressions()` | same |
| `getAllNamedExpressionsSerialized()` | `getAllNamedExpressionsSerialized()` | same |

Scope semantics:

- workbook-scoped and sheet-scoped names both belong on the public surface
- lookup methods should follow HyperFormula scope rules, not a custom nearest-match rule
- formulas stored in workbook cells should be rewritten internally so public formulas keep original name text

## Event Surface

HyperFormula event names:

- `sheetAdded`
- `sheetRemoved`
- `sheetRenamed`
- `namedExpressionAdded`
- `namedExpressionRemoved`
- `valuesUpdated`
- `evaluationSuspended`
- `evaluationResumed`

HyperFormula listener methods:

- `on(event, listener)`
- `once(event, listener)`
- `off(event, listener)`

Design decision:

- `@bilig/headless` should expose the same event names and listener methods
- for compatibility, listener callbacks should accept HyperFormula-style positional arguments
- `bilig` may also expose richer structured events, but that should be additive through a separate `onDetailed()` or `events` namespace, not a replacement for `on()`

Target positional signatures:

| Event | Signature |
| --- | --- |
| `sheetAdded` | `(sheetName)` |
| `sheetRemoved` | `(sheetName, changes)` |
| `sheetRenamed` | `(oldName, newName)` |
| `namedExpressionAdded` | `(name, changes)` |
| `namedExpressionRemoved` | `(name, changes)` |
| `valuesUpdated` | `(changes)` |
| `evaluationSuspended` | `()` |
| `evaluationResumed` | `(changes)` |

## Config Surface

HyperFormula `ConfigParams` keys reviewed from `src/ConfigParams.ts`:

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

Design requirement:

- `HeadlessConfig` should preserve this full key set
- `updateConfig()` must rebuild the headless engine when a config change affects parsing, localization, plugins, precision, or engine limits
- config reads must return a stable clone, not internal mutable state

## Error Surface

HyperFormula exports these public error classes through its main package entry point:

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

`bilig` design requirement:

- `@bilig/headless` should export a parallel public error family with stable names and messages that are close enough for migration
- internal `@bilig/core` or `@bilig/formula` errors should be normalized into these headless errors at the facade boundary

## Compatibility Decisions

### Exact compatibility targets

- constructor/factory workflow
- static language and function registries
- config shape
- workbook reads
- workbook mutation semantics
- named-expression operations
- formula helpers
- event names and listener methods
- return `undefined` behavior for misses where HyperFormula does so

### Deliberate `bilig` differences

- change arrays may carry richer metadata than HyperFormula, but they must still be consumable in the same common headless workflows
- internal getter accessors should return adapters, not raw runtime internals
- sync, replica, snapshot replication, and UI-specific APIs stay in `@bilig/core`

## Proposed `@bilig/headless` Public Surface

Stable exports:

- `HeadlessWorkbook`
- `HeadlessConfig`
- `HeadlessCellAddress`
- `HeadlessCellRange`
- `HeadlessSheet`
- `HeadlessSheets`
- `HeadlessSheetDimensions`
- `HeadlessChange`
- `HeadlessCellChange`
- `HeadlessNamedExpressionChange`
- `HeadlessNamedExpression`
- `SerializedHeadlessNamedExpression`
- `HeadlessFunctionPluginDefinition`
- `HeadlessFunctionTranslationsPackage`
- `HeadlessLanguagePackage`
- `HeadlessWorkbookEventName`
- `HeadlessWorkbookEventMap`
- typed headless errors

Phase-two exports:

- address-mapping policy helpers once `@bilig/core` has stable public equivalents
- internal graph and evaluator adapters

## Implementation Plan

### Phase 1: workflow parity

- ship `HeadlessWorkbook` factories and full read/mutate/history/clipboard/named-expression surface
- normalize config rebuilds
- support function plugins and language registration
- implement formula helpers and dependency queries
- ship typed errors and acceptance tests

### Phase 2: event and compatibility polish

- make `on`, `once`, and `off` use HyperFormula positional callback signatures
- add structured event adapters without breaking compatibility
- close remaining return-shape mismatches
- add package-level exports that mirror HyperFormula categories more closely

### Phase 3: internal accessor adapters

- expose `graph`, `rangeMapping`, `arrayMapping`, `sheetMapping`, `addressMapping`, `dependencyGraph`, `evaluator`, `columnSearch`, and `lazilyTransformingAstService` through read-only adapters
- stabilize the minimum `@bilig/core` introspection APIs required for those adapters

### Phase 4: full migration validation

- add a parity matrix test suite that exercises every public HyperFormula method group
- add fixture-driven tests for:
  - config rebuilds
  - event ordering
  - named-expression scoping
  - array and spill behavior
  - clipboard translation
  - plugin translation registration

## Acceptance Criteria

This design is complete only when:

- `docs/public-api.md` lists `@bilig/headless` as a stable package
- `@bilig/headless` exposes every headless workflow API listed above
- HyperFormula-style event signatures are supported
- the adapter getter tranche is either shipped or explicitly removed from scope in a follow-up ADR
- a parity suite covers all public method groups and all config keys
- the package can be adopted without importing `@bilig/core` directly for standard headless workbook workflows

## Bottom Line

The correct `bilig` response to HyperFormula is not to copy its internals. It is to ship a first-class headless library package with the same practical workbook API surface, backed by `bilig`'s own engine.

That means:

- `@bilig/headless` becomes the public headless workbook library
- HyperFormula's public API is the compatibility checklist
- `@bilig/core` remains the lower-level engine and sync/runtime substrate
