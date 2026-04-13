import { Effect } from "effect";
import {
  ErrorCode,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  type CellRangeRef,
  type CellValue,
} from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { EngineOp } from "@bilig/workbook-domain";
import { CellFlags } from "../../cell-store.js";
import { normalizeRange } from "../../engine-range-utils.js";
import {
  areCellValuesEqual,
  cellValueDisplayText,
  emptyValue,
  errorValue,
  normalizePivotLookupText,
  pivotItemMatches,
} from "../../engine-value-utils.js";
import { materializePivotTable, type PivotDefinitionInput } from "../../pivot-engine.js";
import type { FormulaTable } from "../../formula-table.js";
import type { RangeRegistry } from "../../range-registry.js";
import type { StringPool } from "../../string-pool.js";
import type { WasmKernelFacade } from "../../wasm-facade.js";
import { pivotKey, type WorkbookPivotRecord, type WorkbookStore } from "../../workbook-store.js";
import type { RuntimeFormula } from "../runtime-state.js";
import { EnginePivotError } from "../errors.js";

interface EnginePivotState {
  readonly workbook: WorkbookStore;
  readonly strings: StringPool;
  readonly formulas: FormulaTable<RuntimeFormula>;
  readonly ranges: RangeRegistry;
  readonly wasm: WasmKernelFacade;
  readonly pivotOutputOwners: Map<number, string>;
}

export interface EnginePivotService {
  readonly materializePivot: (
    pivot: WorkbookPivotRecord,
  ) => Effect.Effect<number[], EnginePivotError>;
  readonly resolvePivotData: (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ) => Effect.Effect<CellValue, EnginePivotError>;
  readonly clearOwnedPivot: (
    pivot: WorkbookPivotRecord,
  ) => Effect.Effect<number[], EnginePivotError>;
  readonly clearPivotForCell: (cellIndex: number) => Effect.Effect<number[], EnginePivotError>;
  readonly materializePivotNow: (pivot: WorkbookPivotRecord) => number[];
  readonly resolvePivotDataNow: (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ) => CellValue;
  readonly clearOwnedPivotNow: (pivot: WorkbookPivotRecord) => number[];
  readonly clearPivotForCellNow: (cellIndex: number) => number[];
}

function toPivotDefinition(pivot: WorkbookPivotRecord): PivotDefinitionInput {
  return {
    groupBy: pivot.groupBy,
    values: pivot.values,
  };
}

export function createEnginePivotService(args: {
  readonly state: EnginePivotState;
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number;
  readonly forEachSheetCell: (
    sheetId: number,
    fn: (cellIndex: number, row: number, col: number) => void,
  ) => void;
  readonly flushDeferredKernelSync: () => void;
  readonly scheduleWasmProgramSync: () => void;
  readonly flushWasmProgramSync: () => void;
  readonly applyDerivedOp: (
    op: Extract<
      EngineOp,
      { kind: "upsertSpillRange" | "deleteSpillRange" | "upsertPivotTable" | "deletePivotTable" }
    >,
  ) => number[];
}): EnginePivotService {
  const clearPivotOutputCell = (cellIndex: number): boolean => {
    const currentFlags = args.state.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = args.state.workbook.cellStore.getValue(cellIndex, (id) =>
      args.state.strings.get(id),
    );
    if (currentValue.tag === ValueTag.Empty && (currentFlags & CellFlags.PivotOutput) === 0) {
      args.state.pivotOutputOwners.delete(cellIndex);
      return false;
    }
    args.state.pivotOutputOwners.delete(cellIndex);
    args.state.workbook.cellStore.setValue(cellIndex, emptyValue());
    args.state.workbook.cellStore.flags[cellIndex] =
      currentFlags &
      ~(
        CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput
      );
    return true;
  };

  const setPivotOutputCellValue = (
    cellIndex: number,
    value: CellValue,
    ownerKey: string,
  ): boolean => {
    const currentFlags = args.state.workbook.cellStore.flags[cellIndex] ?? 0;
    const currentValue = args.state.workbook.cellStore.getValue(cellIndex, (id) =>
      args.state.strings.get(id),
    );
    const nextFlags =
      (currentFlags &
        ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild)) |
      CellFlags.PivotOutput;
    if (
      areCellValuesEqual(currentValue, value) &&
      currentFlags === nextFlags &&
      args.state.pivotOutputOwners.get(cellIndex) === ownerKey
    ) {
      return false;
    }
    args.state.workbook.cellStore.setValue(
      cellIndex,
      value,
      value.tag === ValueTag.String ? args.state.strings.intern(value.value) : 0,
    );
    args.state.workbook.cellStore.flags[cellIndex] = nextFlags;
    args.state.pivotOutputOwners.set(cellIndex, ownerKey);
    return true;
  };

  const clearOwnedPivotNow = (pivot: WorkbookPivotRecord): number[] => {
    const changedCellIndices: number[] = [];
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    for (let rowOffset = 0; rowOffset < pivot.rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < pivot.cols; colOffset += 1) {
        const cellIndex = args.state.workbook.getCellIndex(
          pivot.sheetName,
          formatAddress(owner.row + rowOffset, owner.col + colOffset),
        );
        if (cellIndex === undefined || args.state.pivotOutputOwners.get(cellIndex) !== ownerKey) {
          continue;
        }
        if (clearPivotOutputCell(cellIndex)) {
          changedCellIndices.push(cellIndex);
        }
      }
    }
    return changedCellIndices;
  };

  const writePivotOutput = (
    pivot: WorkbookPivotRecord,
    rows: number,
    cols: number,
    values: readonly CellValue[],
    changedCellIndices: number[],
  ): number[] => {
    const sheet = args.state.workbook.getOrCreateSheet(pivot.sheetName);
    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    const changedSeen = new Set(changedCellIndices);

    for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < cols; colOffset += 1) {
        const valueIndex = rowOffset * cols + colOffset;
        const cellValue = values[valueIndex] ?? emptyValue();
        const cellIndex = args.ensureCellTrackedByCoords(
          sheet.id,
          owner.row + rowOffset,
          owner.col + colOffset,
        );
        if (setPivotOutputCellValue(cellIndex, cellValue, ownerKey)) {
          if (!changedSeen.has(cellIndex)) {
            changedSeen.add(cellIndex);
            changedCellIndices.push(cellIndex);
          }
        }
      }
    }

    if (pivot.rows !== rows || pivot.cols !== cols) {
      args.applyDerivedOp({
        kind: "upsertPivotTable",
        name: pivot.name,
        sheetName: pivot.sheetName,
        address: pivot.address,
        source: { ...pivot.source },
        groupBy: [...pivot.groupBy],
        values: pivot.values.map((value) => Object.assign({}, value)),
        rows,
        cols,
      });
    }
    return changedCellIndices;
  };

  const isPivotOutputBlocked = (
    pivot: WorkbookPivotRecord,
    startRow: number,
    startCol: number,
    rows: number,
    cols: number,
  ): boolean => {
    const ownerKey = pivotKey(pivot.sheetName, pivot.address);
    for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
      for (let colOffset = 0; colOffset < cols; colOffset += 1) {
        const targetIndex = args.state.workbook.getCellIndex(
          pivot.sheetName,
          formatAddress(startRow + rowOffset, startCol + colOffset),
        );
        if (targetIndex === undefined) {
          continue;
        }
        const pivotOwner = args.state.pivotOutputOwners.get(targetIndex);
        if (pivotOwner && pivotOwner !== ownerKey) {
          return true;
        }
        if (args.state.formulas.get(targetIndex)) {
          return true;
        }
        const targetFlags = args.state.workbook.cellStore.flags[targetIndex] ?? 0;
        if ((targetFlags & CellFlags.SpillChild) !== 0) {
          return true;
        }
        const targetValue = args.state.workbook.cellStore.getValue(targetIndex, (id) =>
          args.state.strings.get(id),
        );
        if (!pivotOwner && targetValue.tag !== ValueTag.Empty) {
          return true;
        }
      }
    }
    return false;
  };

  const guardPivotOutputWrite = (
    pivot: WorkbookPivotRecord,
    startRow: number,
    startCol: number,
    rows: number,
    cols: number,
    changedCellIndices: number[],
  ): number[] | undefined => {
    if (startRow + rows > MAX_ROWS || startCol + cols > MAX_COLS) {
      return writePivotOutput(pivot, 1, 1, [errorValue(ErrorCode.Spill)], changedCellIndices);
    }
    if (isPivotOutputBlocked(pivot, startRow, startCol, rows, cols)) {
      return writePivotOutput(pivot, 1, 1, [errorValue(ErrorCode.Blocked)], changedCellIndices);
    }
    return undefined;
  };

  const readPivotSourceRows = (range: CellRangeRef): CellValue[][] => {
    const bounds = normalizeRange(range);
    const rows: CellValue[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      const values: CellValue[] = [];
      for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
        const cellIndex = args.state.workbook.getCellIndex(
          range.sheetName,
          formatAddress(row, col),
        );
        values.push(
          cellIndex === undefined
            ? emptyValue()
            : args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id)),
        );
      }
      rows.push(values);
    }
    return rows;
  };

  const materializePivotNow = (pivot: WorkbookPivotRecord): number[] => {
    const changedCellIndices = clearOwnedPivotNow(pivot);
    const sourceSheet = args.state.workbook.getSheet(pivot.source.sheetName);
    if (!sourceSheet) {
      return writePivotOutput(pivot, 1, 1, [errorValue(ErrorCode.Ref)], changedCellIndices);
    }

    const materialized = materializePivotTable(
      toPivotDefinition(pivot),
      readPivotSourceRows(pivot.source),
    );
    if (materialized.kind === "error") {
      return writePivotOutput(
        pivot,
        materialized.rows,
        materialized.cols,
        materialized.values,
        changedCellIndices,
      );
    }

    const owner = parseCellAddress(pivot.address, pivot.sheetName);
    const blockedOutput = guardPivotOutputWrite(
      pivot,
      owner.row,
      owner.col,
      materialized.rows,
      materialized.cols,
      changedCellIndices,
    );
    if (blockedOutput) {
      return blockedOutput;
    }

    return writePivotOutput(
      pivot,
      materialized.rows,
      materialized.cols,
      materialized.values,
      changedCellIndices,
    );
  };

  const resolvePivotDataNow = (
    sheetName: string,
    address: string,
    dataField: string,
    filters: ReadonlyArray<{ field: string; item: CellValue }>,
  ): CellValue => {
    const target = parseCellAddress(address, sheetName);
    const pivot = args.state.workbook.listPivots().find((candidate) => {
      if (candidate.sheetName !== sheetName || candidate.rows <= 0 || candidate.cols <= 0) {
        return false;
      }
      const owner = parseCellAddress(candidate.address, candidate.sheetName);
      return (
        target.row >= owner.row &&
        target.row < owner.row + candidate.rows &&
        target.col >= owner.col &&
        target.col < owner.col + candidate.cols
      );
    });
    if (!pivot) {
      return errorValue(ErrorCode.Ref);
    }

    const normalizedDataField = normalizePivotLookupText(dataField);
    const valueField = pivot.values.find((field) => {
      const defaultLabel = `${field.summarizeBy.toUpperCase()} of ${field.sourceColumn}`;
      return (
        normalizePivotLookupText(field.sourceColumn) === normalizedDataField ||
        normalizePivotLookupText(field.outputLabel?.trim() ?? "") === normalizedDataField ||
        normalizePivotLookupText(defaultLabel) === normalizedDataField
      );
    });
    if (!valueField) {
      return errorValue(ErrorCode.Ref);
    }

    const sourceRows = readPivotSourceRows(pivot.source);
    const headerRow = sourceRows[0];
    if (!headerRow || headerRow.length === 0) {
      return errorValue(ErrorCode.Ref);
    }

    const headerLookup = new Map<string, number>();
    headerRow.forEach((cell, index) => {
      const normalized = normalizePivotLookupText(cellValueDisplayText(cell));
      if (normalized.length > 0 && !headerLookup.has(normalized)) {
        headerLookup.set(normalized, index);
      }
    });

    const valueColumnIndex = headerLookup.get(normalizePivotLookupText(valueField.sourceColumn));
    if (valueColumnIndex === undefined) {
      return errorValue(ErrorCode.Ref);
    }

    const materializedFilters = filters.map((filter) => ({
      fieldIndex: headerLookup.get(normalizePivotLookupText(filter.field)),
      item: filter.item,
    }));
    if (materializedFilters.some((filter) => filter.fieldIndex === undefined)) {
      return errorValue(ErrorCode.Ref);
    }

    for (let filterIndex = 0; filterIndex < materializedFilters.length; filterIndex += 1) {
      const filter = materializedFilters[filterIndex]!;
      const fieldIndex = filter.fieldIndex!;
      const itemSeen = sourceRows
        .slice(1)
        .some((row) => pivotItemMatches(row[fieldIndex] ?? emptyValue(), filter.item));
      if (!itemSeen) {
        return errorValue(ErrorCode.Ref);
      }
    }

    let matched = filters.length === 0;
    let aggregate = 0;
    for (let rowIndex = 1; rowIndex < sourceRows.length; rowIndex += 1) {
      const row = sourceRows[rowIndex] ?? [];
      const matches = materializedFilters.every((filter) =>
        pivotItemMatches(row[filter.fieldIndex!] ?? emptyValue(), filter.item),
      );
      if (!matches) {
        continue;
      }
      matched = true;
      const value = row[valueColumnIndex] ?? emptyValue();
      if (valueField.summarizeBy === "count") {
        aggregate += value.tag === ValueTag.Empty ? 0 : 1;
      } else if (value.tag === ValueTag.Number) {
        aggregate += value.value;
      }
    }

    return matched ? { tag: ValueTag.Number, value: aggregate } : errorValue(ErrorCode.Ref);
  };

  const clearPivotForCellNow = (cellIndex: number): number[] => {
    const ownerKey = args.state.pivotOutputOwners.get(cellIndex);
    if (!ownerKey) {
      return [];
    }
    const pivot = args.state.workbook.getPivotByKey(ownerKey);
    if (!pivot) {
      args.state.pivotOutputOwners.delete(cellIndex);
      return [];
    }
    return args.applyDerivedOp({
      kind: "deletePivotTable",
      sheetName: pivot.sheetName,
      address: pivot.address,
    });
  };

  return {
    materializePivot(pivot) {
      return Effect.try({
        try: () => materializePivotNow(pivot),
        catch: (cause) =>
          new EnginePivotError({
            message: `Failed to materialize pivot ${pivot.name}`,
            cause,
          }),
      });
    },
    resolvePivotData(sheetName, address, dataField, filters) {
      return Effect.try({
        try: () => resolvePivotDataNow(sheetName, address, dataField, filters),
        catch: (cause) =>
          new EnginePivotError({
            message: `Failed to resolve pivot data for ${sheetName}!${address}`,
            cause,
          }),
      });
    },
    clearOwnedPivot(pivot) {
      return Effect.try({
        try: () => clearOwnedPivotNow(pivot),
        catch: (cause) =>
          new EnginePivotError({
            message: `Failed to clear pivot output ownership for ${pivot.name}`,
            cause,
          }),
      });
    },
    clearPivotForCell(cellIndex) {
      return Effect.try({
        try: () => clearPivotForCellNow(cellIndex),
        catch: (cause) =>
          new EnginePivotError({
            message: `Failed to clear pivot ownership for cell ${cellIndex}`,
            cause,
          }),
      });
    },
    materializePivotNow,
    resolvePivotDataNow,
    clearOwnedPivotNow,
    clearPivotForCellNow,
  };
}
