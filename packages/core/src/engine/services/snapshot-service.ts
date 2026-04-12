import { Effect } from "effect";
import {
  ValueTag,
  type CellSnapshot,
  type WorkbookAxisEntrySnapshot,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import { CellFlags } from "../../cell-store.js";
import { cloneCellStyleRecord } from "../../engine-style-utils.js";
import { exportSheetMetadata } from "../../engine-snapshot-utils.js";
import {
  exportReplicaSnapshot as exportReplicaStateSnapshot,
  hydrateReplicaState,
} from "../../replica-state.js";
import type {
  EngineRuntimeState,
  EngineReplicaSnapshot,
  TransactionRecord,
} from "../runtime-state.js";
import { EngineSnapshotError } from "../errors.js";

export interface EngineSnapshotService {
  readonly exportWorkbook: () => Effect.Effect<WorkbookSnapshot, EngineSnapshotError>;
  readonly importWorkbook: (snapshot: WorkbookSnapshot) => Effect.Effect<void, EngineSnapshotError>;
  readonly exportReplica: () => Effect.Effect<EngineReplicaSnapshot, EngineSnapshotError>;
  readonly importReplica: (
    snapshot: EngineReplicaSnapshot,
  ) => Effect.Effect<void, EngineSnapshotError>;
}

export function createEngineSnapshotService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    "workbook" | "strings" | "formulas" | "replicaState" | "entityVersions" | "sheetDeleteVersions"
  >;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly resetWorkbook: (workbookName?: string) => void;
  readonly executeRestoreTransaction: (transaction: TransactionRecord) => void;
}): EngineSnapshotService {
  return {
    exportWorkbook() {
      return Effect.try({
        try: () => {
          const workbook: WorkbookSnapshot["workbook"] = {
            name: args.state.workbook.workbookName,
          };
          const properties = args.state.workbook
            .listWorkbookProperties()
            .map(({ key, value }) => ({ key, value }));
          const definedNames = args.state.workbook
            .listDefinedNames()
            .map(({ name, value }) => ({ name, value }));
          const calculationSettings = args.state.workbook.getCalculationSettings();
          const volatileContext = args.state.workbook.getVolatileContext();
          const tables = args.state.workbook.listTables().map((table) => ({
            name: table.name,
            sheetName: table.sheetName,
            startAddress: table.startAddress,
            endAddress: table.endAddress,
            columnNames: [...table.columnNames],
            headerRow: table.headerRow,
            totalsRow: table.totalsRow,
          }));
          const spills = args.state.workbook
            .listSpills()
            .map(({ sheetName, address, rows, cols }) => ({ sheetName, address, rows, cols }));
          const referencedStyleIds = new Set<string>();
          const referencedFormatIds = new Set<string>();
          args.state.workbook.sheetsByName.forEach((sheet) => {
            sheet.styleRanges.forEach((record) => referencedStyleIds.add(record.styleId));
            sheet.formatRanges.forEach((record) => referencedFormatIds.add(record.formatId));
          });
          for (let cellIndex = 0; cellIndex < args.state.workbook.cellStore.size; cellIndex += 1) {
            const explicitFormat = args.state.workbook.getCellFormat(cellIndex);
            if (explicitFormat !== undefined) {
              referencedFormatIds.add(
                args.state.workbook.internCellNumberFormat(explicitFormat).id,
              );
            }
          }
          const styles = args.state.workbook
            .listCellStyles()
            .filter((style) => referencedStyleIds.has(style.id))
            .map((style) => cloneCellStyleRecord(style));
          const formats = args.state.workbook
            .listCellNumberFormats()
            .filter((format) => referencedFormatIds.has(format.id))
            .map((format) => Object.assign({}, format));
          const pivots = args.state.workbook.listPivots().map((pivot) => ({
            name: pivot.name,
            sheetName: pivot.sheetName,
            address: pivot.address,
            source: { ...pivot.source },
            groupBy: [...pivot.groupBy],
            values: pivot.values.map((value) => Object.assign({}, value)),
            rows: pivot.rows,
            cols: pivot.cols,
          }));
          if (
            properties.length > 0 ||
            definedNames.length > 0 ||
            tables.length > 0 ||
            spills.length > 0 ||
            pivots.length > 0 ||
            styles.length > 0 ||
            formats.length > 0 ||
            calculationSettings.mode !== "automatic" ||
            calculationSettings.compatibilityMode !== "excel-modern" ||
            volatileContext.recalcEpoch !== 0
          ) {
            workbook.metadata = {};
            if (properties.length > 0) {
              workbook.metadata.properties = properties;
            }
            if (definedNames.length > 0) {
              workbook.metadata.definedNames = definedNames;
            }
            if (tables.length > 0) {
              workbook.metadata.tables = tables;
            }
            if (spills.length > 0) {
              workbook.metadata.spills = spills;
            }
            if (pivots.length > 0) {
              workbook.metadata.pivots = pivots;
            }
            if (styles.length > 0) {
              workbook.metadata.styles = styles;
            }
            if (formats.length > 0) {
              workbook.metadata.formats = formats;
            }
            if (
              calculationSettings.mode !== "automatic" ||
              calculationSettings.compatibilityMode !== "excel-modern"
            ) {
              workbook.metadata.calculationSettings = calculationSettings;
            }
            if (volatileContext.recalcEpoch !== 0) {
              workbook.metadata.volatileContext = volatileContext;
            }
          }

          return {
            version: 1,
            workbook,
            sheets: [...args.state.workbook.sheetsByName.values()]
              .toSorted((left, right) => left.order - right.order)
              .map((sheet) => {
                const metadata = exportSheetMetadata(args.state.workbook, sheet.name);
                const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
                sheet.grid.forEachCell((cellIndex) => {
                  const snapshot = args.getCellByIndex(cellIndex);
                  const explicitFormat = args.state.workbook.getCellFormat(cellIndex);
                  if ((snapshot.flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
                    return;
                  }
                  if (
                    snapshot.formula === undefined &&
                    explicitFormat === undefined &&
                    snapshot.version === 0 &&
                    (snapshot.value.tag === ValueTag.Empty || snapshot.value.tag === ValueTag.Error)
                  ) {
                    return;
                  }
                  const cell: WorkbookSnapshot["sheets"][number]["cells"][number] = {
                    address: snapshot.address,
                  };
                  if (explicitFormat !== undefined) {
                    cell.format = explicitFormat;
                  }
                  if (snapshot.formula) {
                    cell.formula = snapshot.formula;
                  } else if (snapshot.value.tag === ValueTag.Number) {
                    cell.value = snapshot.value.value;
                  } else if (snapshot.value.tag === ValueTag.Boolean) {
                    cell.value = snapshot.value.value;
                  } else if (snapshot.value.tag === ValueTag.String) {
                    cell.value = snapshot.value.value;
                  } else {
                    cell.value = null;
                  }
                  cells.push(cell);
                });
                return metadata
                  ? { id: sheet.id, name: sheet.name, order: sheet.order, metadata, cells }
                  : { id: sheet.id, name: sheet.name, order: sheet.order, cells };
              }),
          };
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: "Failed to export workbook snapshot",
            cause,
          }),
      });
    },
    importWorkbook(snapshot) {
      return Effect.try({
        try: () => {
          args.resetWorkbook();
          const ops: import("@bilig/workbook-domain").EngineOp[] = [
            { kind: "upsertWorkbook", name: snapshot.workbook.name },
          ];
          snapshot.workbook.metadata?.properties?.forEach(({ key, value }) => {
            ops.push({ kind: "setWorkbookMetadata", key, value });
          });
          if (snapshot.workbook.metadata?.calculationSettings) {
            ops.push({
              kind: "setCalculationSettings",
              settings: { ...snapshot.workbook.metadata.calculationSettings },
            });
          }
          if (snapshot.workbook.metadata?.volatileContext) {
            ops.push({
              kind: "setVolatileContext",
              context: { ...snapshot.workbook.metadata.volatileContext },
            });
          }
          snapshot.workbook.metadata?.definedNames?.forEach(({ name, value }) => {
            ops.push({ kind: "upsertDefinedName", name, value });
          });
          snapshot.workbook.metadata?.styles?.forEach((style) => {
            ops.push({ kind: "upsertCellStyle", style: cloneCellStyleRecord(style) });
          });
          snapshot.workbook.metadata?.formats?.forEach((format) => {
            ops.push({ kind: "upsertCellNumberFormat", format: { ...format } });
          });
          snapshot.sheets.forEach((sheet) => {
            ops.push({
              kind: "upsertSheet",
              name: sheet.name,
              order: sheet.order,
              ...(typeof sheet.id === "number" ? { id: sheet.id } : {}),
            });
          });
          snapshot.sheets.forEach((sheet) => {
            sheet.metadata?.rows?.forEach(({ index, id, size, hidden }) => {
              const entry: WorkbookAxisEntrySnapshot = { index, id };
              if (size !== undefined) {
                entry.size = size;
              }
              if (hidden !== undefined) {
                entry.hidden = hidden;
              }
              ops.push({
                kind: "insertRows",
                sheetName: sheet.name,
                start: index,
                count: 1,
                entries: [entry],
              });
            });
            sheet.metadata?.columns?.forEach(({ index, id, size, hidden }) => {
              const entry: WorkbookAxisEntrySnapshot = { index, id };
              if (size !== undefined) {
                entry.size = size;
              }
              if (hidden !== undefined) {
                entry.hidden = hidden;
              }
              ops.push({
                kind: "insertColumns",
                sheetName: sheet.name,
                start: index,
                count: 1,
                entries: [entry],
              });
            });
            sheet.metadata?.rowMetadata?.forEach(({ start, count, size, hidden }) => {
              ops.push({
                kind: "updateRowMetadata",
                sheetName: sheet.name,
                start,
                count,
                size: size ?? null,
                hidden: hidden ?? null,
              });
            });
            sheet.metadata?.columnMetadata?.forEach(({ start, count, size, hidden }) => {
              ops.push({
                kind: "updateColumnMetadata",
                sheetName: sheet.name,
                start,
                count,
                size: size ?? null,
                hidden: hidden ?? null,
              });
            });
            if (sheet.metadata?.freezePane) {
              ops.push({
                kind: "setFreezePane",
                sheetName: sheet.name,
                rows: sheet.metadata.freezePane.rows,
                cols: sheet.metadata.freezePane.cols,
              });
            }
            sheet.metadata?.styleRanges?.forEach((styleRange) => {
              ops.push({
                kind: "setStyleRange",
                range: { ...styleRange.range },
                styleId: styleRange.styleId,
              });
            });
            sheet.metadata?.formatRanges?.forEach((formatRange) => {
              ops.push({
                kind: "setFormatRange",
                range: { ...formatRange.range },
                formatId: formatRange.formatId,
              });
            });
            sheet.metadata?.filters?.forEach((range) => {
              ops.push({ kind: "setFilter", sheetName: sheet.name, range: { ...range } });
            });
            sheet.metadata?.sorts?.forEach((sort) => {
              ops.push({
                kind: "setSort",
                sheetName: sheet.name,
                range: { ...sort.range },
                keys: sort.keys.map((key) => Object.assign({}, key)),
              });
            });
            sheet.metadata?.validations?.forEach((validation) => {
              ops.push({
                kind: "setDataValidation",
                validation: structuredClone(validation),
              });
            });
            sheet.metadata?.conditionalFormats?.forEach((format) => {
              ops.push({
                kind: "upsertConditionalFormat",
                format: structuredClone(format),
              });
            });
            sheet.metadata?.commentThreads?.forEach((thread) => {
              ops.push({
                kind: "upsertCommentThread",
                thread: structuredClone(thread),
              });
            });
            sheet.metadata?.notes?.forEach((note) => {
              ops.push({
                kind: "upsertNote",
                note: structuredClone(note),
              });
            });
          });
          snapshot.sheets.forEach((sheet) => {
            sheet.cells.forEach((cell) => {
              if (cell.formula !== undefined) {
                ops.push({
                  kind: "setCellFormula",
                  sheetName: sheet.name,
                  address: cell.address,
                  formula: cell.formula,
                });
              } else {
                ops.push({
                  kind: "setCellValue",
                  sheetName: sheet.name,
                  address: cell.address,
                  value: cell.value ?? null,
                });
              }
              if (cell.format !== undefined) {
                ops.push({
                  kind: "setCellFormat",
                  sheetName: sheet.name,
                  address: cell.address,
                  format: cell.format,
                });
              }
            });
          });
          snapshot.workbook.metadata?.tables?.forEach((table) => {
            ops.push({
              kind: "upsertTable",
              table: {
                name: table.name,
                sheetName: table.sheetName,
                startAddress: table.startAddress,
                endAddress: table.endAddress,
                columnNames: [...table.columnNames],
                headerRow: table.headerRow,
                totalsRow: table.totalsRow,
              },
            });
          });
          snapshot.workbook.metadata?.spills?.forEach((spill) => {
            ops.push({
              kind: "upsertSpillRange",
              sheetName: spill.sheetName,
              address: spill.address,
              rows: spill.rows,
              cols: spill.cols,
            });
          });
          snapshot.workbook.metadata?.pivots?.forEach((pivot) => {
            ops.push({
              kind: "upsertPivotTable",
              name: pivot.name,
              sheetName: pivot.sheetName,
              address: pivot.address,
              source: { ...pivot.source },
              groupBy: [...pivot.groupBy],
              values: pivot.values.map((value) => Object.assign({}, value)),
              rows: pivot.rows,
              cols: pivot.cols,
            });
          });
          const potentialNewCells = snapshot.sheets.reduce(
            (count, sheet) => count + sheet.cells.length,
            0,
          );
          args.executeRestoreTransaction(
            potentialNewCells > 0 ? { ops, potentialNewCells } : { ops },
          );
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: "Failed to import workbook snapshot",
            cause,
          }),
      });
    },
    exportReplica() {
      return Effect.try({
        try: () => ({
          replica: exportReplicaStateSnapshot(args.state.replicaState),
          entityVersions: [...args.state.entityVersions.entries()].map(([entityKey, order]) => ({
            entityKey,
            order,
          })),
          sheetDeleteVersions: [...args.state.sheetDeleteVersions.entries()].map(
            ([sheetName, order]) => ({
              sheetName,
              order,
            }),
          ),
        }),
        catch: (cause) =>
          new EngineSnapshotError({
            message: "Failed to export replica snapshot",
            cause,
          }),
      });
    },
    importReplica(snapshot) {
      return Effect.try({
        try: () => {
          hydrateReplicaState(args.state.replicaState, snapshot.replica);
          args.state.entityVersions.clear();
          snapshot.entityVersions.forEach(({ entityKey, order }) => {
            args.state.entityVersions.set(entityKey, order);
          });
          args.state.sheetDeleteVersions.clear();
          snapshot.sheetDeleteVersions.forEach(({ sheetName, order }) => {
            args.state.sheetDeleteVersions.set(sheetName, order);
          });
        },
        catch: (cause) =>
          new EngineSnapshotError({
            message: "Failed to import replica snapshot",
            cause,
          }),
      });
    },
  };
}
