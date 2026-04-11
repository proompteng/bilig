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
import type { EngineRuntimeState } from "../runtime-state.js";
import { EngineFormulaEvaluationError } from "../errors.js";
import type { EngineLookupService } from "./lookup-service.js";

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
  const readCellValue = (sheetName: string, address: string): CellValue => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address);
    if (cellIndex === undefined) {
      return emptyValue();
    }
    return args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id));
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
      resolveExactVectorMatch: (request) => args.lookup.findExactVectorMatch(request),
      resolveApproximateVectorMatch: (request) => args.lookup.findApproximateVectorMatch(request),
      ...(lookupBuiltinResolver
        ? {
            resolveLookupBuiltin: lookupBuiltinResolver,
          }
        : {}),
    };
    const result = evaluatePlanResult(formula.compiled.jsPlan, evaluationContext);

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
