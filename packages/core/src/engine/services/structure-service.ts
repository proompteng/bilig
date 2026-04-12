import { Effect } from "effect";
import type { SheetFormatRangeSnapshot, SheetStyleRangeSnapshot } from "@bilig/protocol";
import {
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from "@bilig/formula";
import type { EngineOp } from "@bilig/workbook-domain";
import { CellFlags } from "../../cell-store.js";
import { emptyValue } from "../../engine-value-utils.js";
import {
  mapStructuralAxisIndex,
  mapStructuralBoundary,
  structuralTransformForOp,
} from "../../engine-structural-utils.js";
import type { FormulaTable } from "../../formula-table.js";
import type { RuntimeFormula } from "../runtime-state.js";
import { EngineStructureError } from "../errors.js";
import type { WorkbookPivotRecord, WorkbookStore } from "../../workbook-store.js";

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind:
      | "insertRows"
      | "deleteRows"
      | "moveRows"
      | "insertColumns"
      | "deleteColumns"
      | "moveColumns";
  }
>;

interface EngineStructureState {
  readonly workbook: WorkbookStore;
  readonly formulas: FormulaTable<RuntimeFormula>;
  readonly pivotOutputOwners: Map<number, string>;
}

export interface EngineStructureService {
  readonly captureSheetCellState: (
    sheetName: string,
  ) => Effect.Effect<EngineOp[], EngineStructureError>;
  readonly captureRowRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => Effect.Effect<EngineOp[], EngineStructureError>;
  readonly captureColumnRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => Effect.Effect<EngineOp[], EngineStructureError>;
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => Effect.Effect<
    {
      changedCellIndices: number[];
      formulaCellIndices: number[];
    },
    EngineStructureError
  >;
}

export function createEngineStructureService(args: {
  readonly state: EngineStructureState;
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[];
  readonly removeFormula: (cellIndex: number) => boolean;
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[];
  readonly rebuildAllFormulaBindings: () => number[];
}): EngineStructureService {
  const captureStoredCellState = (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] =>
    args.captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress);

  const captureAxisRangeCellState = (
    sheetName: string,
    axis: "row" | "column",
    start: number,
    count: number,
  ): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName);
    if (!sheet) {
      return [];
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = [];
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      const index = axis === "row" ? row : col;
      if (index >= start && index < start + count) {
        captured.push({ cellIndex, row, col });
      }
    });
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex, row, col }) =>
        captureStoredCellState(cellIndex, sheetName, formatAddress(row, col)),
      );
  };

  const captureSheetCellState = (sheetName: string): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName);
    if (!sheet) {
      return [];
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = [];
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      captured.push({ cellIndex, row, col });
    });
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex }) =>
        captureStoredCellState(cellIndex, sheetName, args.state.workbook.getAddress(cellIndex)),
      );
  };

  const rewriteDefinedNamesForStructuralTransform = (
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void => {
    args.state.workbook.listDefinedNames().forEach((record) => {
      if (typeof record.value !== "string" || !record.value.startsWith("=")) {
        return;
      }
      const nextFormula = rewriteFormulaForStructuralTransform(
        record.value.slice(1),
        sheetName,
        sheetName,
        transform,
      );
      if (`=${nextFormula}` !== record.value) {
        args.state.workbook.setDefinedName(record.name, `=${nextFormula}`);
      }
    });
  };

  const rewriteCellFormulasForStructuralTransform = (
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void => {
    args.state.formulas.forEach((formula, cellIndex) => {
      const ownerSheetName = args.state.workbook.getSheetNameById(
        args.state.workbook.cellStore.sheetIds[cellIndex]!,
      );
      formula.source = rewriteFormulaForStructuralTransform(
        formula.source,
        ownerSheetName,
        sheetName,
        transform,
      );
    });
  };

  const clearAllSpillMetadata = (): void => {
    args.state.workbook.listSpills().forEach((spill) => {
      args.state.workbook.deleteSpill(spill.sheetName, spill.address);
    });
  };

  const clearPivotOutputsForSheet = (sheetName: string): void => {
    args.state.workbook
      .listPivots()
      .filter((pivot) => pivot.sheetName === sheetName)
      .forEach((pivot) => {
        args.clearOwnedPivot(pivot);
      });
  };

  const clearDerivedCellArtifacts = (cellIndex: number): void => {
    args.state.pivotOutputOwners.delete(cellIndex);
  };

  const rewriteWorkbookMetadataForStructuralTransform = (
    sheetName: string,
    transform: StructuralAxisTransform,
  ): void => {
    args.state.workbook
      .listTables()
      .filter((table) => table.sheetName === sheetName)
      .forEach((table) => {
        const range = rewriteRangeForStructuralTransform(
          table.startAddress,
          table.endAddress,
          transform,
        );
        if (!range) {
          args.state.workbook.deleteTable(table.name);
          return;
        }
        args.state.workbook.setTable({
          ...table,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        });
      });
    args.state.workbook.listFilters(sheetName).forEach((filter) => {
      const range = rewriteRangeForStructuralTransform(
        filter.range.startAddress,
        filter.range.endAddress,
        transform,
      );
      args.state.workbook.deleteFilter(sheetName, filter.range);
      if (range) {
        args.state.workbook.setFilter(sheetName, {
          ...filter.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        });
      }
    });
    args.state.workbook.listSorts(sheetName).forEach((sort) => {
      const range = rewriteRangeForStructuralTransform(
        sort.range.startAddress,
        sort.range.endAddress,
        transform,
      );
      args.state.workbook.deleteSort(sheetName, sort.range);
      if (!range) {
        return;
      }
      args.state.workbook.setSort(
        sheetName,
        { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
        sort.keys.map((key) => ({
          ...key,
          keyAddress:
            rewriteAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
        })),
      );
    });
    args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
      const range = rewriteRangeForStructuralTransform(
        validation.range.startAddress,
        validation.range.endAddress,
        transform,
      );
      args.state.workbook.deleteDataValidation(sheetName, validation.range);
      if (!range) {
        return;
      }
      const nextValidation = structuredClone(validation);
      nextValidation.range = {
        ...validation.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      };
      if (nextValidation.rule.kind === "list" && nextValidation.rule.source) {
        switch (nextValidation.rule.source.kind) {
          case "cell-ref": {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break;
            }
            const nextAddress = rewriteAddressForStructuralTransform(
              nextValidation.rule.source.address,
              transform,
            );
            if (!nextAddress) {
              return;
            }
            nextValidation.rule.source.address = nextAddress;
            break;
          }
          case "range-ref": {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break;
            }
            const nextSourceRange = rewriteRangeForStructuralTransform(
              nextValidation.rule.source.startAddress,
              nextValidation.rule.source.endAddress,
              transform,
            );
            if (!nextSourceRange) {
              return;
            }
            nextValidation.rule.source.startAddress = nextSourceRange.startAddress;
            nextValidation.rule.source.endAddress = nextSourceRange.endAddress;
            break;
          }
          case "named-range":
          case "structured-ref":
            break;
        }
      }
      args.state.workbook.setDataValidation(nextValidation);
    });
    args.state.workbook.listConditionalFormats(sheetName).forEach((format) => {
      const range = rewriteRangeForStructuralTransform(
        format.range.startAddress,
        format.range.endAddress,
        transform,
      );
      args.state.workbook.deleteConditionalFormat(format.id);
      if (!range) {
        return;
      }
      args.state.workbook.setConditionalFormat({
        ...format,
        range: {
          ...format.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      });
    });
    args.state.workbook.listRangeProtections(sheetName).forEach((protection) => {
      const range = rewriteRangeForStructuralTransform(
        protection.range.startAddress,
        protection.range.endAddress,
        transform,
      );
      args.state.workbook.deleteRangeProtection(protection.id);
      if (!range) {
        return;
      }
      args.state.workbook.setRangeProtection({
        ...protection,
        range: {
          ...protection.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      });
    });
    args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
      const nextAddress = rewriteAddressForStructuralTransform(thread.address, transform);
      args.state.workbook.deleteCommentThread(sheetName, thread.address);
      if (!nextAddress) {
        return;
      }
      args.state.workbook.setCommentThread({
        ...thread,
        address: nextAddress,
      });
    });
    args.state.workbook.listNotes(sheetName).forEach((note) => {
      const nextAddress = rewriteAddressForStructuralTransform(note.address, transform);
      args.state.workbook.deleteNote(sheetName, note.address);
      if (!nextAddress) {
        return;
      }
      args.state.workbook.setNote({
        ...note,
        address: nextAddress,
      });
    });
    const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = [];
    const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = [];
    args.state.workbook.listStyleRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(
        record.range.startAddress,
        record.range.endAddress,
        transform,
      );
      if (!range) {
        return;
      }
      rewrittenStyleRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        styleId: record.styleId,
      });
    });
    args.state.workbook.setStyleRanges(sheetName, rewrittenStyleRanges);
    args.state.workbook.listFormatRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(
        record.range.startAddress,
        record.range.endAddress,
        transform,
      );
      if (!range) {
        return;
      }
      rewrittenFormatRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        formatId: record.formatId,
      });
    });
    args.state.workbook.setFormatRanges(sheetName, rewrittenFormatRanges);
    const freezePane = args.state.workbook.getFreezePane(sheetName);
    if (freezePane) {
      const nextRows =
        transform.axis === "row"
          ? mapStructuralBoundary(freezePane.rows, transform)
          : freezePane.rows;
      const nextCols =
        transform.axis === "column"
          ? mapStructuralBoundary(freezePane.cols, transform)
          : freezePane.cols;
      if (nextRows <= 0 && nextCols <= 0) {
        args.state.workbook.clearFreezePane(sheetName);
      } else {
        args.state.workbook.setFreezePane(sheetName, nextRows, nextCols);
      }
    }
    args.state.workbook.listPivots().forEach((pivot) => {
      const nextAddress =
        pivot.sheetName === sheetName
          ? rewriteAddressForStructuralTransform(pivot.address, transform)
          : pivot.address;
      const nextSource =
        pivot.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(
              pivot.source.startAddress,
              pivot.source.endAddress,
              transform,
            )
          : { startAddress: pivot.source.startAddress, endAddress: pivot.source.endAddress };
      if (!nextAddress || !nextSource) {
        args.clearOwnedPivot(pivot);
        args.state.workbook.deletePivot(pivot.sheetName, pivot.address);
        return;
      }
      args.state.workbook.setPivot({
        ...pivot,
        address: nextAddress,
        source: {
          ...pivot.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      });
    });
    args.state.workbook.listCharts().forEach((chart) => {
      const nextAddress =
        chart.sheetName === sheetName
          ? rewriteAddressForStructuralTransform(chart.address, transform)
          : chart.address;
      const nextSource =
        chart.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(
              chart.source.startAddress,
              chart.source.endAddress,
              transform,
            )
          : { startAddress: chart.source.startAddress, endAddress: chart.source.endAddress };
      if (!nextAddress || !nextSource) {
        args.state.workbook.deleteChart(chart.id);
        return;
      }
      args.state.workbook.setChart({
        ...chart,
        address: nextAddress,
        source: {
          ...chart.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      });
    });
  };

  return {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => captureSheetCellState(sheetName),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture sheet cell state for ${sheetName}`,
            cause,
          }),
      });
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, "row", start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture row state for ${sheetName}`,
            cause,
          }),
      });
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, "column", start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture column state for ${sheetName}`,
            cause,
          }),
      });
    },
    applyStructuralAxisOp(op) {
      return Effect.try({
        try: () => {
          const axis = op.kind.includes("Rows") ? "row" : "column";
          const transform = structuralTransformForOp(op);
          const sheetName = op.sheetName;

          rewriteDefinedNamesForStructuralTransform(sheetName, transform);
          rewriteCellFormulasForStructuralTransform(sheetName, transform);
          rewriteWorkbookMetadataForStructuralTransform(sheetName, transform);

          switch (op.kind) {
            case "insertRows":
              args.state.workbook.insertRows(sheetName, op.start, op.count, op.entries);
              break;
            case "deleteRows":
              args.state.workbook.deleteRows(sheetName, op.start, op.count);
              break;
            case "moveRows":
              args.state.workbook.moveRows(sheetName, op.start, op.count, op.target);
              break;
            case "insertColumns":
              args.state.workbook.insertColumns(sheetName, op.start, op.count, op.entries);
              break;
            case "deleteColumns":
              args.state.workbook.deleteColumns(sheetName, op.start, op.count);
              break;
            case "moveColumns":
              args.state.workbook.moveColumns(sheetName, op.start, op.count, op.target);
              break;
          }

          const remapped = args.state.workbook.remapSheetCells(sheetName, axis, (index) =>
            mapStructuralAxisIndex(index, transform),
          );
          remapped.removedCellIndices.forEach((cellIndex) => {
            clearDerivedCellArtifacts(cellIndex);
            args.removeFormula(cellIndex);
            args.state.workbook.setCellFormat(cellIndex, null);
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue());
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(
                CellFlags.HasFormula |
                CellFlags.JsOnly |
                CellFlags.InCycle |
                CellFlags.SpillChild |
                CellFlags.PivotOutput
              );
          });

          clearAllSpillMetadata();
          clearPivotOutputsForSheet(sheetName);
          const formulaCellIndices = args.rebuildAllFormulaBindings();
          return {
            changedCellIndices: [...remapped.changedCellIndices, ...remapped.removedCellIndices],
            formulaCellIndices,
          };
        },
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to apply structural operation ${op.kind}`,
            cause,
          }),
      });
    },
  };
}
