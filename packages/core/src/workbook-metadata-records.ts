import type { CellRangeRef, WorkbookDefinedNameValueSnapshot } from "@bilig/protocol";
import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from "./workbook-range-records.js";
import {
  normalizeWorkbookObjectName,
  pivotKey,
  type WorkbookAxisMetadataRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookSortKeyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
} from "./workbook-metadata-types.js";

export function cloneDefinedNameRecord(
  record: WorkbookDefinedNameRecord,
): WorkbookDefinedNameRecord {
  return {
    name: record.name,
    value: cloneDefinedNameValue(record.value),
  };
}

export function cloneDefinedNameValue(
  value: WorkbookDefinedNameValueSnapshot,
): WorkbookDefinedNameValueSnapshot {
  if (value === null || typeof value !== "object") {
    return value;
  }
  switch (value.kind) {
    case "scalar":
      return { kind: "scalar", value: value.value };
    case "cell-ref":
      return { kind: "cell-ref", sheetName: value.sheetName, address: value.address };
    case "range-ref":
      return {
        kind: "range-ref",
        sheetName: value.sheetName,
        startAddress: value.startAddress,
        endAddress: value.endAddress,
      };
    case "structured-ref":
      return {
        kind: "structured-ref",
        tableName: value.tableName,
        columnName: value.columnName,
      };
    case "formula":
      return { kind: "formula", formula: value.formula };
    default:
      return value;
  }
}

export function clonePropertyRecord(record: WorkbookPropertyRecord): WorkbookPropertyRecord {
  return { key: record.key, value: record.value };
}

export function cloneTableRecord(record: WorkbookTableRecord): WorkbookTableRecord {
  return {
    name: record.name,
    sheetName: record.sheetName,
    startAddress: record.startAddress,
    endAddress: record.endAddress,
    columnNames: [...record.columnNames],
    headerRow: record.headerRow,
    totalsRow: record.totalsRow,
  };
}

export function cloneFilterRecord(record: WorkbookFilterRecord): WorkbookFilterRecord {
  return {
    sheetName: record.sheetName,
    range: { ...record.range },
  };
}

export function cloneSortKeyRecord(record: WorkbookSortKeyRecord): WorkbookSortKeyRecord {
  return { ...record };
}

export function cloneSortRecord(record: WorkbookSortRecord): WorkbookSortRecord {
  return {
    sheetName: record.sheetName,
    range: { ...record.range },
    keys: record.keys.map(cloneSortKeyRecord),
  };
}

export function clonePivotRecord(record: WorkbookPivotRecord): WorkbookPivotRecord {
  return {
    ...record,
    source: { ...record.source },
    groupBy: [...record.groupBy],
    values: record.values.map((value) => ({ ...value })),
  };
}

export function cloneSpillRecord(record: WorkbookSpillRecord): WorkbookSpillRecord {
  return {
    sheetName: record.sheetName,
    address: record.address,
    rows: record.rows,
    cols: record.cols,
  };
}

function axisMetadataKey(sheetName: string, start: number, count: number): string {
  return `${sheetName}:${start}:${count}`;
}

export function filterKey(sheetName: string, range: CellRangeRef): string {
  const normalized = canonicalWorkbookRangeRef(range);
  return `${sheetName}:${normalized.startAddress}:${normalized.endAddress}`;
}

export function sortKey(sheetName: string, range: CellRangeRef): string {
  const normalized = canonicalWorkbookRangeRef(range);
  return `${sheetName}:${normalized.startAddress}:${normalized.endAddress}`;
}

export function tableKey(name: string): string {
  return normalizeWorkbookObjectName(name, "Tables");
}

export function spillKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`;
}

export function deleteRecordsBySheet<T>(
  bucket: Map<string, T>,
  sheetName: string,
  readSheetName: (record: T) => string,
): void {
  for (const [key, record] of bucket.entries()) {
    if (readSheetName(record) === sheetName) {
      bucket.delete(key);
    }
  }
}

export function rekeyRecords<T>(bucket: Map<string, T>, rewrite: (record: T) => T): void {
  const rewritten = [...bucket.values()].map((record) => rewrite(record));
  bucket.clear();
  rewritten.forEach((record) => {
    bucket.set(recordKey(record), record);
  });
}

function recordKey(record: unknown): string {
  if (isFreezePaneRecord(record)) {
    return record.sheetName;
  }
  if (isAxisMetadataRecord(record)) {
    return axisMetadataKey(record.sheetName, record.start, record.count);
  }
  if (isFilterRecord(record)) {
    return filterKey(record.sheetName, record.range);
  }
  if (isSortRecord(record)) {
    return sortKey(record.sheetName, record.range);
  }
  if (isTableRecord(record)) {
    return tableKey(record.name);
  }
  if (isSpillRecord(record)) {
    return spillKey(record.sheetName, record.address);
  }
  if (isPivotRecord(record)) {
    return pivotKey(record.sheetName, record.address);
  }
  throw new Error("Unsupported workbook metadata record");
}

function isFreezePaneRecord(record: unknown): record is WorkbookFreezePaneRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "rows" in record &&
    "cols" in record &&
    !("address" in record)
  );
}

function isAxisMetadataRecord(record: unknown): record is WorkbookAxisMetadataRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "start" in record &&
    "count" in record &&
    "size" in record &&
    "hidden" in record
  );
}

function isFilterRecord(record: unknown): record is WorkbookFilterRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "range" in record &&
    !("keys" in record)
  );
}

function isSortRecord(record: unknown): record is WorkbookSortRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "range" in record &&
    "keys" in record
  );
}

function isTableRecord(record: unknown): record is WorkbookTableRecord {
  return (
    typeof record === "object" && record !== null && "name" in record && "columnNames" in record
  );
}

function isSpillRecord(record: unknown): record is WorkbookSpillRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    "address" in record &&
    "rows" in record &&
    "cols" in record &&
    !("source" in record)
  );
}

function isPivotRecord(record: unknown): record is WorkbookPivotRecord {
  return (
    typeof record === "object" && record !== null && "sheetName" in record && "source" in record
  );
}
