import { HeadlessWorkbook } from "./headless-workbook.js";
import type {
  HeadlessAddressFormatOptions,
  HeadlessAddressLike,
  HeadlessArrayMappingAdapter,
  HeadlessAxisInterval,
  HeadlessAxisSwapMapping,
  HeadlessCellAddress,
  HeadlessCellChange,
  HeadlessCellRange,
  HeadlessCellType,
  HeadlessCellValueDetailedType,
  HeadlessCellValueType,
  HeadlessChange,
  HeadlessColumnSearchAdapter,
  HeadlessConfig,
  HeadlessDateTime,
  HeadlessDependencyGraphAdapter,
  HeadlessDependencyRef,
  HeadlessEvaluatorAdapter,
  HeadlessFunctionArgument,
  HeadlessFunctionArgumentType,
  HeadlessFunctionMetadata,
  HeadlessFunctionPlugin,
  HeadlessFunctionPluginDefinition,
  HeadlessFunctionTranslationsPackage,
  HeadlessGraphAdapter,
  HeadlessLanguagePackage,
  HeadlessLazilyTransformingAstServiceAdapter,
  HeadlessLicenseKeyValidityState,
  HeadlessNamedExpression,
  HeadlessNamedExpressionChange,
  HeadlessRangeMappingAdapter,
  HeadlessSheet,
  HeadlessSheetDimensions,
  HeadlessSheetMappingAdapter,
  HeadlessSheets,
  HeadlessStats,
  HeadlessWorkbookDetailedEventMap,
  HeadlessWorkbookDetailedListener,
  HeadlessWorkbookEventMap,
  HeadlessWorkbookEventName,
  HeadlessWorkbookInternals,
  HeadlessWorkbookListener,
  RawCellContent,
  SerializedHeadlessNamedExpression,
} from "./types.js";

export const WorkPaper = HeadlessWorkbook;

export type WorkPaper = HeadlessWorkbook;
export type WorkPaperConfig = HeadlessConfig;
export type WorkPaperSheet = HeadlessSheet;
export type WorkPaperSheets = HeadlessSheets;
export type WorkPaperCellAddress = HeadlessCellAddress;
export type WorkPaperCellRange = HeadlessCellRange;
export type WorkPaperAddressLike = HeadlessAddressLike;
export type WorkPaperAddressFormatOptions = HeadlessAddressFormatOptions;
export type WorkPaperAxisInterval = HeadlessAxisInterval;
export type WorkPaperAxisSwapMapping = HeadlessAxisSwapMapping;
export type WorkPaperChange = HeadlessChange;
export type WorkPaperCellChange = HeadlessCellChange;
export type WorkPaperNamedExpressionChange = HeadlessNamedExpressionChange;
export type WorkPaperNamedExpression = HeadlessNamedExpression;
export type SerializedWorkPaperNamedExpression = SerializedHeadlessNamedExpression;
export type WorkPaperFunctionArgumentType = HeadlessFunctionArgumentType;
export type WorkPaperFunctionArgument = HeadlessFunctionArgument;
export type WorkPaperFunctionMetadata = HeadlessFunctionMetadata;
export type WorkPaperFunctionPlugin = HeadlessFunctionPlugin;
export type WorkPaperFunctionPluginDefinition = HeadlessFunctionPluginDefinition;
export type WorkPaperFunctionTranslationsPackage = HeadlessFunctionTranslationsPackage;
export type WorkPaperLanguagePackage = HeadlessLanguagePackage;
export type WorkPaperLicenseKeyValidityState = HeadlessLicenseKeyValidityState;
export type WorkPaperEventMap = HeadlessWorkbookEventMap;
export type WorkPaperDetailedEventMap = HeadlessWorkbookDetailedEventMap;
export type WorkPaperEventName = HeadlessWorkbookEventName;
export type WorkPaperListener<EventName extends WorkPaperEventName> =
  HeadlessWorkbookListener<EventName>;
export type WorkPaperDetailedListener<EventName extends WorkPaperEventName> =
  HeadlessWorkbookDetailedListener<EventName>;
export type WorkPaperCellType = HeadlessCellType;
export type WorkPaperCellValueType = HeadlessCellValueType;
export type WorkPaperCellValueDetailedType = HeadlessCellValueDetailedType;
export type WorkPaperDependencyRef = HeadlessDependencyRef;
export type WorkPaperDateTime = HeadlessDateTime;
export type WorkPaperStats = HeadlessStats;
export type WorkPaperGraphAdapter = HeadlessGraphAdapter;
export type WorkPaperRangeMappingAdapter = HeadlessRangeMappingAdapter;
export type WorkPaperArrayMappingAdapter = HeadlessArrayMappingAdapter;
export type WorkPaperSheetMappingAdapter = HeadlessSheetMappingAdapter;
export type WorkPaperDependencyGraphAdapter = HeadlessDependencyGraphAdapter;
export type WorkPaperEvaluatorAdapter = HeadlessEvaluatorAdapter;
export type WorkPaperColumnSearchAdapter = HeadlessColumnSearchAdapter;
export type WorkPaperLazilyTransformingAstServiceAdapter =
  HeadlessLazilyTransformingAstServiceAdapter;
export type WorkPaperInternals = HeadlessWorkbookInternals;
export type WorkPaperSheetDimensions = HeadlessSheetDimensions;
export type WorkPaperRawCellContent = RawCellContent;
