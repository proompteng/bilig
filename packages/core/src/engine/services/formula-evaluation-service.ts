import { Effect } from "effect";
import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import {
  createLookupBuiltinResolver,
  evaluatePlanResult,
  formatAddress,
  isArrayValue,
  type EvaluationContext,
  type FormulaNode,
  type RangeBuiltinArgument,
  parseCellAddress,
  parseRangeAddress,
} from "@bilig/formula";
import { CellFlags } from "../../cell-store.js";
import { definedNameValueToCellValue } from "../../engine-metadata-utils.js";
import { emptyValue, errorValue } from "../../engine-value-utils.js";
import type { EngineRuntimeState, RuntimeFormula } from "../runtime-state.js";
import { EngineFormulaEvaluationError } from "../errors.js";
import type {
  EngineLookupService,
  PreparedApproximateVectorLookup,
  PreparedExactVectorLookup,
} from "./lookup-service.js";

export interface EngineFormulaEvaluationService {
  readonly evaluateUnsupportedFormula: (
    cellIndex: number,
  ) => Effect.Effect<number[], EngineFormulaEvaluationError>;
  readonly resolveStructuredReference: (
    tableName: string,
    columnName: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>;
  readonly resolveSpillReference: (
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>;
  readonly resolveMultipleOperations: (request: {
    formulaSheetName: string;
    formulaAddress: string;
    rowCellSheetName: string;
    rowCellAddress: string;
    rowReplacementSheetName: string;
    rowReplacementAddress: string;
    columnCellSheetName?: string;
    columnCellAddress?: string;
    columnReplacementSheetName?: string;
    columnReplacementAddress?: string;
  }) => Effect.Effect<CellValue, EngineFormulaEvaluationError>;
}

function evaluationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function referenceReplacementKey(sheetName: string, address: string): string {
  return `${sheetName.trim().toUpperCase()}!${address.trim().toUpperCase()}`;
}

type DirectLookupOperandInstruction =
  | { opcode: "push-cell"; address: string; sheetName?: string }
  | { opcode: "push-number"; value: number }
  | { opcode: "push-boolean"; value: boolean }
  | { opcode: "push-string"; value: string }
  | { opcode: "push-error"; code: ErrorCode }
  | { opcode: "push-name"; name: string };

type DirectExactLookupInstruction = {
  opcode: "lookup-exact-match";
  sheetName?: string;
  start: string;
  end: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  searchMode: 1 | -1;
};

type DirectApproximateLookupInstruction = {
  opcode: "lookup-approximate-match";
  sheetName?: string;
  start: string;
  end: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  matchMode: 1 | -1;
};

type DirectVectorLookupInstruction =
  | DirectExactLookupInstruction
  | DirectApproximateLookupInstruction;

type CachedDirectVectorLookup =
  | {
      kind: "exact";
      operandSheetName: string;
      operandRow: number;
      operandCol: number;
      prepared: PreparedExactVectorLookup;
      searchMode: 1 | -1;
    }
  | {
      kind: "approximate";
      operandSheetName: string;
      operandRow: number;
      operandCol: number;
      prepared: PreparedApproximateVectorLookup;
      matchMode: 1 | -1;
    };

function directLookupCacheKey(
  kind: "exact" | "approximate",
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  col: number,
): string {
  return `${kind}\t${sheetName}\t${rowStart}\t${rowEnd}\t${col}`;
}

function isDirectLookupOperandInstruction(value: unknown): value is DirectLookupOperandInstruction {
  if (!value || typeof value !== "object") {
    return false;
  }
  const opcode = Reflect.get(value, "opcode");
  switch (opcode) {
    case "push-cell":
      return typeof Reflect.get(value, "address") === "string";
    case "push-number":
      return typeof Reflect.get(value, "value") === "number";
    case "push-boolean":
      return typeof Reflect.get(value, "value") === "boolean";
    case "push-string":
      return typeof Reflect.get(value, "value") === "string";
    case "push-error":
      return typeof Reflect.get(value, "code") === "number";
    case "push-name":
      return typeof Reflect.get(value, "name") === "string";
    default:
      return false;
  }
}

function isDirectVectorLookupInstruction(value: unknown): value is DirectVectorLookupInstruction {
  if (!value || typeof value !== "object") {
    return false;
  }
  const opcode = Reflect.get(value, "opcode");
  if (
    (opcode !== "lookup-exact-match" && opcode !== "lookup-approximate-match") ||
    typeof Reflect.get(value, "start") !== "string" ||
    typeof Reflect.get(value, "end") !== "string" ||
    typeof Reflect.get(value, "startRow") !== "number" ||
    typeof Reflect.get(value, "endRow") !== "number" ||
    typeof Reflect.get(value, "startCol") !== "number" ||
    typeof Reflect.get(value, "endCol") !== "number"
  ) {
    return false;
  }
  if (opcode === "lookup-exact-match") {
    const searchMode = Reflect.get(value, "searchMode");
    return searchMode === 1 || searchMode === -1;
  }
  const matchMode = Reflect.get(value, "matchMode");
  return matchMode === 1 || matchMode === -1;
}

function isReturnInstruction(value: unknown): value is { opcode: "return" } {
  return !!value && typeof value === "object" && Reflect.get(value, "opcode") === "return";
}

export function createEngineFormulaEvaluationService(args: {
  readonly state: Pick<EngineRuntimeState, "workbook" | "strings" | "formulas" | "useColumnIndex">;
  readonly lookup: EngineLookupService;
  readonly materializeSpill: (
    cellIndex: number,
    arrayValue: { values: CellValue[]; rows: number; cols: number },
  ) => import("../runtime-state.js").SpillMaterialization;
  readonly clearOwnedSpill: (cellIndex: number) => number[];
  readonly resolvePivotData: (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ) => CellValue;
}): EngineFormulaEvaluationService {
  const preparedExactLookupCache = new WeakMap<
    RuntimeFormula,
    Map<string, PreparedExactVectorLookup>
  >();
  const preparedApproximateLookupCache = new WeakMap<
    RuntimeFormula,
    Map<string, PreparedApproximateVectorLookup>
  >();
  const directVectorLookupPlanCache = new WeakMap<
    RuntimeFormula,
    CachedDirectVectorLookup | null
  >();

  const readCellValue = (sheetName: string, address: string): CellValue => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id));
  };

  const readCellValueAt = (sheetName: string, row: number, col: number): CellValue => {
    const sheet = args.state.workbook.getSheet(sheetName);
    if (!sheet) {
      return emptyValue();
    }
    const cellIndex = sheet.grid.get(row, col);
    if (cellIndex === -1) {
      return emptyValue();
    }
    const cellStore = args.state.workbook.cellStore;
    const tag = cellStore.tags[cellIndex];
    switch (tag) {
      case undefined:
      case ValueTag.Empty:
        return emptyValue();
      case ValueTag.Number:
        return { tag: ValueTag.Number, value: cellStore.numbers[cellIndex] ?? 0 };
      case ValueTag.Boolean:
        return { tag: ValueTag.Boolean, value: (cellStore.numbers[cellIndex] ?? 0) !== 0 };
      case ValueTag.String: {
        const stringId = cellStore.stringIds[cellIndex] ?? 0;
        return {
          tag: ValueTag.String,
          value: stringId === 0 ? "" : args.state.strings.get(stringId),
          stringId,
        };
      }
      case ValueTag.Error:
        return errorValue(cellStore.errors[cellIndex] ?? ErrorCode.Value);
    }
    return emptyValue();
  };

  const readRangeValues = (
    sheetName: string,
    start: string,
    end: string,
    refKind: "cells" | "rows" | "cols",
    replacements?: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting?: Set<string>,
  ): CellValue[] => {
    if (refKind !== "cells") {
      return [];
    }
    const range = parseRangeAddress(`${start}:${end}`, sheetName);
    if (range.kind !== "cells") {
      return [];
    }
    const values: CellValue[] = [];
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        values.push(
          replacements && visiting
            ? evaluateCellWithReferenceReplacements(
                sheetName,
                formatAddress(row, col),
                replacements,
                visiting,
              )
            : readCellValue(sheetName, formatAddress(row, col)),
        );
      }
    }
    return values;
  };

  const resolveIndexedExactMatch = (
    lookupValue: CellValue,
    range: RangeBuiltinArgument,
  ): number | undefined => {
    if (!args.state.useColumnIndex || range.refKind !== "cells" || range.cols !== 1) {
      return undefined;
    }
    if (!range.sheetName || !range.start || !range.end) {
      return undefined;
    }
    const result = args.lookup.findExactVectorMatch({
      lookupValue,
      sheetName: range.sheetName,
      start: range.start,
      end: range.end,
      searchMode: 1,
    });
    return result.handled ? result.position : undefined;
  };

  const lookupBuiltinResolver = args.state.useColumnIndex
    ? createLookupBuiltinResolver({
        resolveIndexedExactMatch,
      })
    : undefined;

  const resolvePreparedExactVectorMatch = (
    formula: RuntimeFormula,
    request: {
      lookupValue: CellValue;
      sheetName: string;
      start: string;
      end: string;
      startRow: number;
      endRow: number;
      startCol: number;
      endCol: number;
      searchMode: 1 | -1;
    },
  ) => {
    if (request.startCol !== request.endCol) {
      return args.lookup.findExactVectorMatch(request);
    }
    let preparedByKey = preparedExactLookupCache.get(formula);
    if (!preparedByKey) {
      preparedByKey = new Map();
      preparedExactLookupCache.set(formula, preparedByKey);
    }
    const cacheKey = directLookupCacheKey(
      "exact",
      request.sheetName,
      request.startRow,
      request.endRow,
      request.startCol,
    );
    let prepared = preparedByKey.get(cacheKey);
    if (!prepared) {
      prepared = args.lookup.prepareExactVectorLookup({
        sheetName: request.sheetName,
        rowStart: request.startRow,
        rowEnd: request.endRow,
        col: request.startCol,
      });
      preparedByKey.set(cacheKey, prepared);
    }
    return args.lookup.findPreparedExactVectorMatch({
      lookupValue: request.lookupValue,
      prepared,
      searchMode: request.searchMode,
    });
  };

  const resolvePreparedApproximateVectorMatch = (
    formula: RuntimeFormula,
    request: {
      lookupValue: CellValue;
      sheetName: string;
      start: string;
      end: string;
      startRow: number;
      endRow: number;
      startCol: number;
      endCol: number;
      matchMode: 1 | -1;
    },
  ) => {
    if (request.startCol !== request.endCol) {
      return args.lookup.findApproximateVectorMatch(request);
    }
    let preparedByKey = preparedApproximateLookupCache.get(formula);
    if (!preparedByKey) {
      preparedByKey = new Map();
      preparedApproximateLookupCache.set(formula, preparedByKey);
    }
    const cacheKey = directLookupCacheKey(
      "approximate",
      request.sheetName,
      request.startRow,
      request.endRow,
      request.startCol,
    );
    let prepared = preparedByKey.get(cacheKey);
    if (!prepared) {
      prepared = args.lookup.prepareApproximateVectorLookup({
        sheetName: request.sheetName,
        rowStart: request.startRow,
        rowEnd: request.endRow,
        col: request.startCol,
      });
      preparedByKey.set(cacheKey, prepared);
    }
    return args.lookup.findPreparedApproximateVectorMatch({
      lookupValue: request.lookupValue,
      prepared,
      matchMode: request.matchMode,
    });
  };

  const getCachedDirectVectorLookup = (
    formula: RuntimeFormula,
    ownerSheetName: string,
    jsPlan: readonly unknown[],
  ): CachedDirectVectorLookup | null => {
    const cached = directVectorLookupPlanCache.get(formula);
    if (cached !== undefined) {
      return cached;
    }
    const [operand, lookup, terminal] = jsPlan;
    if (!operand || !lookup || !terminal || !isReturnInstruction(terminal)) {
      directVectorLookupPlanCache.set(formula, null);
      return null;
    }
    if (!isDirectLookupOperandInstruction(operand) || !isDirectVectorLookupInstruction(lookup)) {
      directVectorLookupPlanCache.set(formula, null);
      return null;
    }
    if (operand.opcode !== "push-cell") {
      directVectorLookupPlanCache.set(formula, null);
      return null;
    }
    const operandSheetName = operand.sheetName ?? ownerSheetName;
    const operandAddress = parseCellAddress(operand.address, operandSheetName);
    if (lookup.opcode === "lookup-exact-match") {
      const prepared = args.lookup.prepareExactVectorLookup({
        sheetName: lookup.sheetName ?? ownerSheetName,
        rowStart: lookup.startRow,
        rowEnd: lookup.endRow,
        col: lookup.startCol,
      });
      const directLookup: CachedDirectVectorLookup = {
        kind: "exact",
        operandSheetName,
        operandRow: operandAddress.row,
        operandCol: operandAddress.col,
        prepared,
        searchMode: lookup.searchMode,
      };
      directVectorLookupPlanCache.set(formula, directLookup);
      return directLookup;
    }
    if (lookup.opcode === "lookup-approximate-match") {
      const prepared = args.lookup.prepareApproximateVectorLookup({
        sheetName: lookup.sheetName ?? ownerSheetName,
        rowStart: lookup.startRow,
        rowEnd: lookup.endRow,
        col: lookup.startCol,
      });
      const directLookup: CachedDirectVectorLookup = {
        kind: "approximate",
        operandSheetName,
        operandRow: operandAddress.row,
        operandCol: operandAddress.col,
        prepared,
        matchMode: lookup.matchMode,
      };
      directVectorLookupPlanCache.set(formula, directLookup);
      return directLookup;
    }
    directVectorLookupPlanCache.set(formula, null);
    return null;
  };

  const tryEvaluateDirectVectorLookup = (
    formula: RuntimeFormula,
    ownerSheetName: string,
    jsPlan: readonly unknown[],
  ): CellValue | undefined => {
    const directLookup = getCachedDirectVectorLookup(formula, ownerSheetName, jsPlan);
    if (!directLookup) {
      return undefined;
    }
    const lookupValue = readCellValueAt(
      directLookup.operandSheetName,
      directLookup.operandRow,
      directLookup.operandCol,
    );
    if (directLookup.kind === "exact") {
      const result = args.lookup.findPreparedExactVectorMatch({
        lookupValue,
        prepared: directLookup.prepared,
        searchMode: directLookup.searchMode,
      });
      return result.handled
        ? result.position === undefined
          ? errorValue(ErrorCode.NA)
          : { tag: ValueTag.Number, value: result.position }
        : undefined;
    }
    const result = args.lookup.findPreparedApproximateVectorMatch({
      lookupValue,
      prepared: directLookup.prepared,
      matchMode: directLookup.matchMode,
    });
    return result.handled
      ? result.position === undefined
        ? errorValue(ErrorCode.NA)
        : { tag: ValueTag.Number, value: result.position }
      : undefined;
  };

  const resolveStructuredReferenceNow = (
    tableName: string,
    columnName: string,
  ): FormulaNode | undefined => {
    const table = args.state.workbook.getTable(tableName);
    if (!table) {
      return undefined;
    }
    const columnIndex = table.columnNames.findIndex(
      (name) => name.trim().toUpperCase() === columnName.trim().toUpperCase(),
    );
    if (columnIndex === -1) {
      return undefined;
    }
    const start = parseCellAddress(table.startAddress, table.sheetName);
    const end = parseCellAddress(table.endAddress, table.sheetName);
    const startRow = start.row + (table.headerRow ? 1 : 0);
    const endRow = end.row - (table.totalsRow ? 1 : 0);
    if (endRow < startRow) {
      return { kind: "ErrorLiteral", code: ErrorCode.Ref };
    }
    const column = start.col + columnIndex;
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: table.sheetName,
      start: formatAddress(startRow, column),
      end: formatAddress(endRow, column),
    };
  };

  const resolveSpillReferenceNow = (
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ): FormulaNode | undefined => {
    const targetSheetName = sheetName ?? currentSheetName;
    const spill = args.state.workbook.getSpill(targetSheetName, address);
    if (!spill) {
      return undefined;
    }
    const owner = parseCellAddress(address, targetSheetName);
    return {
      kind: "RangeRef",
      refKind: "cells",
      sheetName: targetSheetName,
      start: owner.text,
      end: formatAddress(owner.row + spill.rows - 1, owner.col + spill.cols - 1),
    };
  };

  const evaluateCellWithReferenceReplacements = (
    sheetName: string,
    address: string,
    replacements: ReadonlyMap<string, { sheetName: string; address: string }>,
    visiting: Set<string>,
  ): CellValue => {
    const replacementKey = referenceReplacementKey(sheetName, address);
    const replacement = replacements.get(replacementKey);
    if (replacement) {
      return evaluateCellWithReferenceReplacements(
        replacement.sheetName,
        replacement.address,
        replacements,
        visiting,
      );
    }

    const visitKey = referenceReplacementKey(sheetName, address);
    if (visiting.has(visitKey)) {
      return errorValue(ErrorCode.Cycle);
    }

    const cellIndex = args.state.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }

    const formula = args.state.formulas.get(cellIndex);
    if (!formula) {
      return args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id));
    }

    visiting.add(visitKey);
    const evaluationContext: EvaluationContext = {
      sheetName,
      currentAddress: address,
      resolveCell: (targetSheetName, targetAddress) =>
        evaluateCellWithReferenceReplacements(
          targetSheetName,
          targetAddress,
          replacements,
          visiting,
        ),
      resolveRange: (targetSheetName, start, end, refKind) =>
        readRangeValues(targetSheetName, start, end, refKind, replacements, visiting),
      resolveName: (name: string) => {
        const definedName = args.state.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings);
      },
      resolveFormula: (targetSheetName: string, targetAddress: string) => {
        const targetCellIndex = args.state.workbook.getCellIndex(targetSheetName, targetAddress);
        return targetCellIndex === undefined
          ? undefined
          : args.state.formulas.get(targetCellIndex)?.source;
      },
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address: pivotAddress,
        filters,
      }: {
        dataField: string;
        sheetName: string;
        address: string;
        filters: ReadonlyArray<{ field: string; item: CellValue }>;
      }) => args.resolvePivotData(pivotSheetName, pivotAddress, dataField, filters),
      resolveMultipleOperations: (nested: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) => resolveMultipleOperationsNow(nested),
      listSheetNames: () =>
        [...args.state.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
    };
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);
    visiting.delete(visitKey);
    return isArrayValue(result) ? (result.values[0] ?? emptyValue()) : result;
  };

  const resolveMultipleOperationsNow = (request: {
    formulaSheetName: string;
    formulaAddress: string;
    rowCellSheetName: string;
    rowCellAddress: string;
    rowReplacementSheetName: string;
    rowReplacementAddress: string;
    columnCellSheetName?: string;
    columnCellAddress?: string;
    columnReplacementSheetName?: string;
    columnReplacementAddress?: string;
  }): CellValue => {
    const replacements = new Map<string, { sheetName: string; address: string }>();
    replacements.set(referenceReplacementKey(request.rowCellSheetName, request.rowCellAddress), {
      sheetName: request.rowReplacementSheetName,
      address: request.rowReplacementAddress,
    });
    if (
      request.columnCellSheetName &&
      request.columnCellAddress &&
      request.columnReplacementSheetName &&
      request.columnReplacementAddress
    ) {
      replacements.set(
        referenceReplacementKey(request.columnCellSheetName, request.columnCellAddress),
        {
          sheetName: request.columnReplacementSheetName,
          address: request.columnReplacementAddress,
        },
      );
    }
    return evaluateCellWithReferenceReplacements(
      request.formulaSheetName,
      request.formulaAddress,
      replacements,
      new Set<string>(),
    );
  };

  const evaluateUnsupportedFormulaNow = (cellIndex: number): number[] => {
    const formula = args.state.formulas.get(cellIndex);
    const sheetName = args.state.workbook.getSheetNameById(
      args.state.workbook.cellStore.sheetIds[cellIndex]!,
    );
    if (!formula || !sheetName) {
      return [];
    }

    const evaluationContext: EvaluationContext = {
      sheetName,
      currentAddress: args.state.workbook.getAddress(cellIndex),
      resolveCell: (targetSheetName: string, address: string) =>
        readCellValue(targetSheetName, address),
      resolveRange: (
        targetSheetName: string,
        start: string,
        end: string,
        refKind: "cells" | "rows" | "cols",
      ) => readRangeValues(targetSheetName, start, end, refKind),
      resolveName: (name: string) => {
        const definedName = args.state.workbook.getDefinedName(name);
        if (!definedName) {
          return errorValue(ErrorCode.Name);
        }
        return definedNameValueToCellValue(definedName.value, args.state.strings);
      },
      resolveFormula: (targetSheetName: string, address: string) => {
        const targetCellIndex = args.state.workbook.getCellIndex(targetSheetName, address);
        return targetCellIndex === undefined
          ? undefined
          : args.state.formulas.get(targetCellIndex)?.source;
      },
      resolvePivotData: ({
        dataField,
        sheetName: pivotSheetName,
        address,
        filters,
      }: {
        dataField: string;
        sheetName: string;
        address: string;
        filters: ReadonlyArray<{ field: string; item: CellValue }>;
      }) => args.resolvePivotData(pivotSheetName, address, dataField, filters),
      resolveMultipleOperations: (request: {
        formulaSheetName: string;
        formulaAddress: string;
        rowCellSheetName: string;
        rowCellAddress: string;
        rowReplacementSheetName: string;
        rowReplacementAddress: string;
        columnCellSheetName?: string;
        columnCellAddress?: string;
        columnReplacementSheetName?: string;
        columnReplacementAddress?: string;
      }) => resolveMultipleOperationsNow(request),
      listSheetNames: () =>
        [...args.state.workbook.sheetsByName.values()]
          .toSorted((left, right) => left.order - right.order)
          .map((sheet) => sheet.name),
      resolveExactVectorMatch: (request) => {
        if (
          request.startRow === undefined ||
          request.endRow === undefined ||
          request.startCol === undefined ||
          request.endCol === undefined
        ) {
          return args.lookup.findExactVectorMatch(request);
        }
        return resolvePreparedExactVectorMatch(formula, request);
      },
      resolveApproximateVectorMatch: (request) => {
        if (
          request.startRow === undefined ||
          request.endRow === undefined ||
          request.startCol === undefined ||
          request.endCol === undefined
        ) {
          return args.lookup.findApproximateVectorMatch(request);
        }
        return resolvePreparedApproximateVectorMatch(formula, request);
      },
      ...(lookupBuiltinResolver
        ? {
            resolveLookupBuiltin: lookupBuiltinResolver,
          }
        : {}),
    };
    const result =
      tryEvaluateDirectVectorLookup(formula, sheetName, formula.compiled.jsPlan) ??
      evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);

    const materialization = isArrayValue(result)
      ? args.materializeSpill(cellIndex, result)
      : {
          changedCellIndices: args.clearOwnedSpill(cellIndex),
          ownerValue: result,
        };

    args.state.workbook.cellStore.flags[cellIndex] =
      (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
      ~(CellFlags.SpillChild | CellFlags.PivotOutput);
    args.state.workbook.cellStore.setValue(
      cellIndex,
      materialization.ownerValue,
      materialization.ownerValue.tag === ValueTag.String
        ? args.state.strings.intern(materialization.ownerValue.value)
        : 0,
    );
    return materialization.changedCellIndices;
  };

  return {
    evaluateUnsupportedFormula(cellIndex) {
      return Effect.try({
        try: () => evaluateUnsupportedFormulaNow(cellIndex),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to evaluate formula ${cellIndex}`, cause),
            cause,
          }),
      });
    },
    resolveStructuredReference(tableName, columnName) {
      return Effect.try({
        try: () => resolveStructuredReferenceNow(tableName, columnName),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(
              `Failed to resolve structured reference ${tableName}[${columnName}]`,
              cause,
            ),
            cause,
          }),
      });
    },
    resolveSpillReference(currentSheetName, sheetName, address) {
      return Effect.try({
        try: () => resolveSpillReferenceNow(currentSheetName, sheetName, address),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage(`Failed to resolve spill reference ${address}#`, cause),
            cause,
          }),
      });
    },
    resolveMultipleOperations(request) {
      return Effect.try({
        try: () => resolveMultipleOperationsNow(request),
        catch: (cause) =>
          new EngineFormulaEvaluationError({
            message: evaluationErrorMessage("Failed to resolve MULTIPLE.OPERATIONS", cause),
            cause,
          }),
      });
    },
  };
}
