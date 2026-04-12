import { Cause, Effect, Exit } from "effect";
import type {
  CellRangeRef,
  LiteralInput,
  WorkbookCalculationSettingsSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookPivotSnapshot,
  WorkbookTableSnapshot,
  WorkbookVolatileContextSnapshot,
} from "@bilig/protocol";
import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from "./workbook-range-records.js";
import {
  cloneDataValidationRecord,
  cloneDefinedNameRecord,
  cloneDefinedNameValue,
  cloneFilterRecord,
  clonePivotRecord,
  clonePropertyRecord,
  cloneSortKeyRecord,
  cloneSortRecord,
  cloneSpillRecord,
  cloneTableRecord,
  dataValidationKey,
  deleteRecordsBySheet,
  filterKey,
  rekeyRecords,
  sortKey,
  spillKey,
  tableKey,
} from "./workbook-metadata-records.js";
import {
  createWorkbookMetadataRecord,
  normalizeDefinedName,
  pivotKey,
  type WorkbookCalculationSettingsRecord,
  type WorkbookDataValidationRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookMetadataRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookSortKeyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
} from "./workbook-metadata-types.js";

function metadataErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message;
}

function renameDataValidationSourceSheet(
  record: WorkbookDataValidationRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDataValidationRecord {
  const cloned = cloneDataValidationRecord(record);
  if (cloned.rule.kind !== "list" || !cloned.rule.source) {
    return cloned;
  }
  switch (cloned.rule.source.kind) {
    case "cell-ref":
    case "range-ref":
      if (cloned.rule.source.sheetName === oldSheetName) {
        cloned.rule.source.sheetName = newSheetName;
      }
      return cloned;
    case "named-range":
    case "structured-ref":
      return cloned;
  }
  return cloned;
}

export class WorkbookMetadataError extends Error {
  readonly _tag = "WorkbookMetadataError";
  override readonly cause: unknown;

  constructor(args: { message: string; cause: unknown }) {
    super(args.message);
    this.name = "WorkbookMetadataError";
    this.cause = args.cause;
  }
}

export interface WorkbookMetadataService {
  readonly renameSheet: (
    oldSheetName: string,
    newSheetName: string,
  ) => Effect.Effect<void, WorkbookMetadataError>;
  readonly deleteSheetRecords: (sheetName: string) => Effect.Effect<void, WorkbookMetadataError>;
  readonly reset: () => Effect.Effect<void, WorkbookMetadataError>;
  readonly setWorkbookProperty: (
    key: string,
    value: LiteralInput,
  ) => Effect.Effect<WorkbookPropertyRecord | undefined, WorkbookMetadataError>;
  readonly getWorkbookProperty: (
    key: string,
  ) => Effect.Effect<WorkbookPropertyRecord | undefined, WorkbookMetadataError>;
  readonly listWorkbookProperties: () => Effect.Effect<
    WorkbookPropertyRecord[],
    WorkbookMetadataError
  >;
  readonly setCalculationSettings: (
    settings: WorkbookCalculationSettingsSnapshot,
  ) => Effect.Effect<WorkbookCalculationSettingsRecord, WorkbookMetadataError>;
  readonly getCalculationSettings: () => Effect.Effect<
    WorkbookCalculationSettingsRecord,
    WorkbookMetadataError
  >;
  readonly setVolatileContext: (
    context: WorkbookVolatileContextSnapshot,
  ) => Effect.Effect<WorkbookVolatileContextRecord, WorkbookMetadataError>;
  readonly getVolatileContext: () => Effect.Effect<
    WorkbookVolatileContextRecord,
    WorkbookMetadataError
  >;
  readonly setDefinedName: (
    name: string,
    value: WorkbookDefinedNameValueSnapshot,
  ) => Effect.Effect<WorkbookDefinedNameRecord, WorkbookMetadataError>;
  readonly getDefinedName: (
    name: string,
  ) => Effect.Effect<WorkbookDefinedNameRecord | undefined, WorkbookMetadataError>;
  readonly deleteDefinedName: (name: string) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listDefinedNames: () => Effect.Effect<
    WorkbookDefinedNameRecord[],
    WorkbookMetadataError
  >;
  readonly setTable: (
    record: WorkbookTableSnapshot,
  ) => Effect.Effect<WorkbookTableRecord, WorkbookMetadataError>;
  readonly getTable: (
    name: string,
  ) => Effect.Effect<WorkbookTableRecord | undefined, WorkbookMetadataError>;
  readonly deleteTable: (name: string) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listTables: () => Effect.Effect<WorkbookTableRecord[], WorkbookMetadataError>;
  readonly setFreezePane: (
    sheetName: string,
    rows: number,
    cols: number,
  ) => Effect.Effect<WorkbookFreezePaneRecord, WorkbookMetadataError>;
  readonly getFreezePane: (
    sheetName: string,
  ) => Effect.Effect<WorkbookFreezePaneRecord | undefined, WorkbookMetadataError>;
  readonly clearFreezePane: (sheetName: string) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly setFilter: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<WorkbookFilterRecord, WorkbookMetadataError>;
  readonly getFilter: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<WorkbookFilterRecord | undefined, WorkbookMetadataError>;
  readonly deleteFilter: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listFilters: (
    sheetName: string,
  ) => Effect.Effect<WorkbookFilterRecord[], WorkbookMetadataError>;
  readonly setSort: (
    sheetName: string,
    range: CellRangeRef,
    keys: readonly WorkbookSortKeyRecord[],
  ) => Effect.Effect<WorkbookSortRecord, WorkbookMetadataError>;
  readonly getSort: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<WorkbookSortRecord | undefined, WorkbookMetadataError>;
  readonly deleteSort: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listSorts: (
    sheetName: string,
  ) => Effect.Effect<WorkbookSortRecord[], WorkbookMetadataError>;
  readonly setDataValidation: (
    record: WorkbookDataValidationSnapshot,
  ) => Effect.Effect<WorkbookDataValidationRecord, WorkbookMetadataError>;
  readonly getDataValidation: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<WorkbookDataValidationRecord | undefined, WorkbookMetadataError>;
  readonly deleteDataValidation: (
    sheetName: string,
    range: CellRangeRef,
  ) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listDataValidations: (
    sheetName: string,
  ) => Effect.Effect<WorkbookDataValidationRecord[], WorkbookMetadataError>;
  readonly setSpill: (
    sheetName: string,
    address: string,
    rows: number,
    cols: number,
  ) => Effect.Effect<WorkbookSpillRecord, WorkbookMetadataError>;
  readonly getSpill: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<WorkbookSpillRecord | undefined, WorkbookMetadataError>;
  readonly deleteSpill: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listSpills: () => Effect.Effect<WorkbookSpillRecord[], WorkbookMetadataError>;
  readonly setPivot: (
    record: WorkbookPivotSnapshot,
  ) => Effect.Effect<WorkbookPivotRecord, WorkbookMetadataError>;
  readonly getPivot: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<WorkbookPivotRecord | undefined, WorkbookMetadataError>;
  readonly getPivotByKey: (
    key: string,
  ) => Effect.Effect<WorkbookPivotRecord | undefined, WorkbookMetadataError>;
  readonly deletePivot: (
    sheetName: string,
    address: string,
  ) => Effect.Effect<boolean, WorkbookMetadataError>;
  readonly listPivots: () => Effect.Effect<WorkbookPivotRecord[], WorkbookMetadataError>;
}

export function createWorkbookMetadataService(
  metadata: WorkbookMetadataRecord,
): WorkbookMetadataService {
  const renameSheetNow = (oldSheetName: string, newSheetName: string): void => {
    rekeyRecords(metadata.freezePanes, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record,
    );
    rekeyRecords(metadata.rowMetadata, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record,
    );
    rekeyRecords(metadata.columnMetadata, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record,
    );
    rekeyRecords(metadata.filters, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName:
                record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
          }
        : cloneFilterRecord(record),
    );
    rekeyRecords(metadata.sorts, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName:
                record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
            keys: record.keys.map(cloneSortKeyRecord),
          }
        : cloneSortRecord(record),
    );
    rekeyRecords(metadata.dataValidations, (record) => {
      const cloned = renameDataValidationSourceSheet(record, oldSheetName, newSheetName);
      if (cloned.range.sheetName === oldSheetName) {
        cloned.range.sheetName = newSheetName;
      }
      return cloned;
    });
    rekeyRecords(metadata.tables, (record) =>
      record.sheetName === oldSheetName
        ? { ...record, sheetName: newSheetName }
        : cloneTableRecord(record),
    );
    rekeyRecords(metadata.spills, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : { ...record },
    );
    rekeyRecords(metadata.pivots, (record) =>
      record.sheetName === oldSheetName || record.source.sheetName === oldSheetName
        ? {
            ...record,
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            source: {
              ...record.source,
              sheetName:
                record.source.sheetName === oldSheetName ? newSheetName : record.source.sheetName,
            },
            groupBy: [...record.groupBy],
            values: record.values.map((value) => ({ ...value })),
          }
        : clonePivotRecord(record),
    );
  };

  const deleteSheetRecordsNow = (sheetName: string): void => {
    deleteRecordsBySheet(metadata.tables, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.spills, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.pivots, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.rowMetadata, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.columnMetadata, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.filters, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.sorts, sheetName, (record) => record.sheetName);
    deleteRecordsBySheet(metadata.dataValidations, sheetName, (record) => record.range.sheetName);
    metadata.freezePanes.delete(sheetName);
  };

  const resetNow = (): void => {
    const defaults = createWorkbookMetadataRecord();
    metadata.properties.clear();
    metadata.definedNames.clear();
    metadata.tables.clear();
    metadata.spills.clear();
    metadata.pivots.clear();
    metadata.rowMetadata.clear();
    metadata.columnMetadata.clear();
    metadata.freezePanes.clear();
    metadata.filters.clear();
    metadata.sorts.clear();
    metadata.dataValidations.clear();
    metadata.calculationSettings = defaults.calculationSettings;
    metadata.volatileContext = defaults.volatileContext;
  };

  return {
    renameSheet(oldSheetName, newSheetName) {
      return Effect.try({
        try: () => renameSheetNow(oldSheetName, newSheetName),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to rename workbook sheet metadata", cause),
            cause,
          }),
      });
    },
    deleteSheetRecords(sheetName) {
      return Effect.try({
        try: () => deleteSheetRecordsNow(sheetName),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete workbook sheet metadata", cause),
            cause,
          }),
      });
    },
    reset() {
      return Effect.try({
        try: resetNow,
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to reset workbook metadata", cause),
            cause,
          }),
      });
    },
    setWorkbookProperty(key, value) {
      return Effect.try({
        try: () => {
          const trimmedKey = normalizeMetadataKey(key);
          if (value === null) {
            metadata.properties.delete(trimmedKey);
            return undefined;
          }
          const record: WorkbookPropertyRecord = { key: trimmedKey, value };
          metadata.properties.set(trimmedKey, record);
          return { ...record };
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set workbook property", cause),
            cause,
          }),
      });
    },
    getWorkbookProperty(key) {
      return Effect.try({
        try: () => {
          const record = metadata.properties.get(normalizeMetadataKey(key));
          return record ? { ...record } : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get workbook property", cause),
            cause,
          }),
      });
    },
    listWorkbookProperties() {
      return Effect.try({
        try: () =>
          [...metadata.properties.values()]
            .toSorted((left, right) => left.key.localeCompare(right.key))
            .map(clonePropertyRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list workbook properties", cause),
            cause,
          }),
      });
    },
    setCalculationSettings(settings) {
      return Effect.try({
        try: () => {
          metadata.calculationSettings = {
            compatibilityMode: "excel-modern",
            ...settings,
          };
          return { ...metadata.calculationSettings };
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set calculation settings", cause),
            cause,
          }),
      });
    },
    getCalculationSettings() {
      return Effect.try({
        try: () => ({ ...metadata.calculationSettings }),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get calculation settings", cause),
            cause,
          }),
      });
    },
    setVolatileContext(context) {
      return Effect.try({
        try: () => {
          metadata.volatileContext = { ...context };
          return { ...metadata.volatileContext };
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set volatile context", cause),
            cause,
          }),
      });
    },
    getVolatileContext() {
      return Effect.try({
        try: () => ({ ...metadata.volatileContext }),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get volatile context", cause),
            cause,
          }),
      });
    },
    setDefinedName(name, value) {
      return Effect.try({
        try: () => {
          const trimmedName = name.trim();
          const record: WorkbookDefinedNameRecord = {
            name: trimmedName,
            value: cloneDefinedNameValue(value),
          };
          metadata.definedNames.set(normalizeDefinedName(trimmedName), record);
          return cloneDefinedNameRecord(record);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set defined name", cause),
            cause,
          }),
      });
    },
    getDefinedName(name) {
      return Effect.try({
        try: () => {
          const record = metadata.definedNames.get(normalizeDefinedName(name));
          return record ? cloneDefinedNameRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get defined name", cause),
            cause,
          }),
      });
    },
    deleteDefinedName(name) {
      return Effect.try({
        try: () => metadata.definedNames.delete(normalizeDefinedName(name)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete defined name", cause),
            cause,
          }),
      });
    },
    listDefinedNames() {
      return Effect.try({
        try: () =>
          [...metadata.definedNames.values()]
            .toSorted((left, right) =>
              normalizeDefinedName(left.name).localeCompare(normalizeDefinedName(right.name)),
            )
            .map(cloneDefinedNameRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list defined names", cause),
            cause,
          }),
      });
    },
    setTable(record) {
      return Effect.try({
        try: () => {
          const stored: WorkbookTableRecord = {
            name: record.name.trim(),
            sheetName: record.sheetName,
            startAddress: record.startAddress,
            endAddress: record.endAddress,
            columnNames: [...record.columnNames],
            headerRow: record.headerRow,
            totalsRow: record.totalsRow,
          };
          metadata.tables.set(tableKey(stored.name), stored);
          return cloneTableRecord(stored);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set table metadata", cause),
            cause,
          }),
      });
    },
    getTable(name) {
      return Effect.try({
        try: () => {
          const record = metadata.tables.get(tableKey(name));
          return record ? cloneTableRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get table metadata", cause),
            cause,
          }),
      });
    },
    deleteTable(name) {
      return Effect.try({
        try: () => metadata.tables.delete(tableKey(name)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete table metadata", cause),
            cause,
          }),
      });
    },
    listTables() {
      return Effect.try({
        try: () =>
          [...metadata.tables.values()]
            .toSorted((left, right) => tableKey(left.name).localeCompare(tableKey(right.name)))
            .map(cloneTableRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list table metadata", cause),
            cause,
          }),
      });
    },
    setFreezePane(sheetName, rows, cols) {
      return Effect.try({
        try: () => {
          const record: WorkbookFreezePaneRecord = { sheetName, rows, cols };
          metadata.freezePanes.set(sheetName, record);
          return { ...record };
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set freeze pane metadata", cause),
            cause,
          }),
      });
    },
    getFreezePane(sheetName) {
      return Effect.try({
        try: () => {
          const record = metadata.freezePanes.get(sheetName);
          return record ? { ...record } : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get freeze pane metadata", cause),
            cause,
          }),
      });
    },
    clearFreezePane(sheetName) {
      return Effect.try({
        try: () => metadata.freezePanes.delete(sheetName),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to clear freeze pane metadata", cause),
            cause,
          }),
      });
    },
    setFilter(sheetName, range) {
      return Effect.try({
        try: () => {
          const storedRange = canonicalWorkbookRangeRef(range);
          const record: WorkbookFilterRecord = { sheetName, range: storedRange };
          metadata.filters.set(filterKey(sheetName, storedRange), record);
          return cloneFilterRecord(record);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set filter metadata", cause),
            cause,
          }),
      });
    },
    getFilter(sheetName, range) {
      return Effect.try({
        try: () => {
          const record = metadata.filters.get(filterKey(sheetName, range));
          return record ? cloneFilterRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get filter metadata", cause),
            cause,
          }),
      });
    },
    deleteFilter(sheetName, range) {
      return Effect.try({
        try: () => metadata.filters.delete(filterKey(sheetName, range)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete filter metadata", cause),
            cause,
          }),
      });
    },
    listFilters(sheetName) {
      return Effect.try({
        try: () =>
          [...metadata.filters.values()]
            .filter((record) => record.sheetName === sheetName)
            .toSorted((left, right) =>
              filterKey(left.sheetName, left.range).localeCompare(
                filterKey(right.sheetName, right.range),
              ),
            )
            .map(cloneFilterRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list filter metadata", cause),
            cause,
          }),
      });
    },
    setSort(sheetName, range, keys) {
      return Effect.try({
        try: () => {
          const storedRange = canonicalWorkbookRangeRef(range);
          const record: WorkbookSortRecord = {
            sheetName,
            range: storedRange,
            keys: keys.map(cloneSortKeyRecord),
          };
          metadata.sorts.set(sortKey(sheetName, storedRange), record);
          return cloneSortRecord(record);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set sort metadata", cause),
            cause,
          }),
      });
    },
    getSort(sheetName, range) {
      return Effect.try({
        try: () => {
          const record = metadata.sorts.get(sortKey(sheetName, range));
          return record ? cloneSortRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get sort metadata", cause),
            cause,
          }),
      });
    },
    deleteSort(sheetName, range) {
      return Effect.try({
        try: () => metadata.sorts.delete(sortKey(sheetName, range)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete sort metadata", cause),
            cause,
          }),
      });
    },
    listSorts(sheetName) {
      return Effect.try({
        try: () =>
          [...metadata.sorts.values()]
            .filter((record) => record.sheetName === sheetName)
            .toSorted((left, right) =>
              sortKey(left.sheetName, left.range).localeCompare(
                sortKey(right.sheetName, right.range),
              ),
            )
            .map(cloneSortRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list sort metadata", cause),
            cause,
          }),
      });
    },
    setDataValidation(record) {
      return Effect.try({
        try: () => {
          const storedRange = canonicalWorkbookRangeRef(record.range);
          const nextRecord: WorkbookDataValidationRecord = {
            range: storedRange,
            rule: record.rule,
          };
          if (record.allowBlank !== undefined) {
            nextRecord.allowBlank = record.allowBlank;
          }
          if (record.showDropdown !== undefined) {
            nextRecord.showDropdown = record.showDropdown;
          }
          if (record.promptTitle !== undefined) {
            nextRecord.promptTitle = record.promptTitle;
          }
          if (record.promptMessage !== undefined) {
            nextRecord.promptMessage = record.promptMessage;
          }
          if (record.errorStyle !== undefined) {
            nextRecord.errorStyle = record.errorStyle;
          }
          if (record.errorTitle !== undefined) {
            nextRecord.errorTitle = record.errorTitle;
          }
          if (record.errorMessage !== undefined) {
            nextRecord.errorMessage = record.errorMessage;
          }
          const stored: WorkbookDataValidationRecord = cloneDataValidationRecord(nextRecord);
          metadata.dataValidations.set(
            dataValidationKey(storedRange.sheetName, storedRange),
            stored,
          );
          return cloneDataValidationRecord(stored);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set data validation metadata", cause),
            cause,
          }),
      });
    },
    getDataValidation(sheetName, range) {
      return Effect.try({
        try: () => {
          const record = metadata.dataValidations.get(dataValidationKey(sheetName, range));
          return record ? cloneDataValidationRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get data validation metadata", cause),
            cause,
          }),
      });
    },
    deleteDataValidation(sheetName, range) {
      return Effect.try({
        try: () => metadata.dataValidations.delete(dataValidationKey(sheetName, range)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete data validation metadata", cause),
            cause,
          }),
      });
    },
    listDataValidations(sheetName) {
      return Effect.try({
        try: () =>
          [...metadata.dataValidations.values()]
            .filter((record) => record.range.sheetName === sheetName)
            .toSorted((left, right) =>
              dataValidationKey(left.range.sheetName, left.range).localeCompare(
                dataValidationKey(right.range.sheetName, right.range),
              ),
            )
            .map(cloneDataValidationRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list data validation metadata", cause),
            cause,
          }),
      });
    },
    setSpill(sheetName, address, rows, cols) {
      return Effect.try({
        try: () => {
          const normalizedAddress = canonicalWorkbookAddress(sheetName, address);
          const record: WorkbookSpillRecord = { sheetName, address: normalizedAddress, rows, cols };
          metadata.spills.set(spillKey(sheetName, normalizedAddress), record);
          return { ...record };
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set spill metadata", cause),
            cause,
          }),
      });
    },
    getSpill(sheetName, address) {
      return Effect.try({
        try: () => {
          const record = metadata.spills.get(spillKey(sheetName, address));
          return record ? { ...record } : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get spill metadata", cause),
            cause,
          }),
      });
    },
    deleteSpill(sheetName, address) {
      return Effect.try({
        try: () => metadata.spills.delete(spillKey(sheetName, address)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete spill metadata", cause),
            cause,
          }),
      });
    },
    listSpills() {
      return Effect.try({
        try: () =>
          [...metadata.spills.values()]
            .toSorted((left, right) =>
              `${left.sheetName}!${left.address}`.localeCompare(
                `${right.sheetName}!${right.address}`,
              ),
            )
            .map(cloneSpillRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list spill metadata", cause),
            cause,
          }),
      });
    },
    setPivot(record) {
      return Effect.try({
        try: () => {
          const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address);
          const stored: WorkbookPivotRecord = {
            ...record,
            name: record.name.trim(),
            address: normalizedAddress,
            groupBy: [...record.groupBy],
            values: record.values.map((value) => ({ ...value })),
            source: canonicalWorkbookRangeRef(record.source),
          };
          metadata.pivots.set(pivotKey(record.sheetName, normalizedAddress), stored);
          return clonePivotRecord(stored);
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to set pivot metadata", cause),
            cause,
          }),
      });
    },
    getPivot(sheetName, address) {
      return Effect.try({
        try: () => {
          const record = metadata.pivots.get(pivotKey(sheetName, address));
          return record ? clonePivotRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get pivot metadata", cause),
            cause,
          }),
      });
    },
    getPivotByKey(key) {
      return Effect.try({
        try: () => {
          const record = metadata.pivots.get(key);
          return record ? clonePivotRecord(record) : undefined;
        },
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to get pivot metadata by key", cause),
            cause,
          }),
      });
    },
    deletePivot(sheetName, address) {
      return Effect.try({
        try: () => metadata.pivots.delete(pivotKey(sheetName, address)),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to delete pivot metadata", cause),
            cause,
          }),
      });
    },
    listPivots() {
      return Effect.try({
        try: () =>
          [...metadata.pivots.values()]
            .toSorted((left, right) =>
              `${left.sheetName}!${left.address}`.localeCompare(
                `${right.sheetName}!${right.address}`,
              ),
            )
            .map(clonePivotRecord),
        catch: (cause) =>
          new WorkbookMetadataError({
            message: metadataErrorMessage("Failed to list pivot metadata", cause),
            cause,
          }),
      });
    },
  };
}

export function runWorkbookMetadataEffect<Success, Failure>(
  effect: Effect.Effect<Success, Failure>,
): Success {
  const exit = Effect.runSyncExit(effect);
  if (Exit.isSuccess(exit)) {
    return exit.value;
  }
  throw Cause.squash(exit.cause);
}

function normalizeMetadataKey(key: string): string {
  const trimmed = key.trim();
  if (trimmed.length === 0) {
    throw new Error("Workbook metadata keys must be non-empty");
  }
  return trimmed;
}
