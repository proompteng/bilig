import type {
  CellRangeRef,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookValidationListSourceSnapshot,
} from "@bilig/protocol";
import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from "./workbook-range-records.js";
import {
  type WorkbookCommentEntryRecord,
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookDataValidationRecord,
  normalizeWorkbookObjectName,
  type WorkbookNoteRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
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

function cloneValidationListSource(
  source: WorkbookValidationListSourceSnapshot,
): WorkbookValidationListSourceSnapshot {
  switch (source.kind) {
    case "named-range":
      return { kind: "named-range", name: source.name };
    case "cell-ref":
      return {
        kind: "cell-ref",
        sheetName: source.sheetName,
        address: canonicalWorkbookAddress(source.sheetName, source.address),
      };
    case "range-ref":
      return {
        kind: "range-ref",
        ...canonicalWorkbookRangeRef({
          sheetName: source.sheetName,
          startAddress: source.startAddress,
          endAddress: source.endAddress,
        }),
      };
    case "structured-ref":
      return {
        kind: "structured-ref",
        tableName: source.tableName,
        columnName: source.columnName,
      };
  }
}

export function cloneDataValidationRule(
  rule: WorkbookDataValidationRuleSnapshot,
): WorkbookDataValidationRuleSnapshot {
  switch (rule.kind) {
    case "list": {
      const cloned: Extract<WorkbookDataValidationRuleSnapshot, { kind: "list" }> = {
        kind: "list",
      };
      if (rule.values) {
        cloned.values = [...rule.values];
      }
      if (rule.source) {
        cloned.source = cloneValidationListSource(rule.source);
      }
      return cloned;
    }
    case "checkbox": {
      const cloned: Extract<WorkbookDataValidationRuleSnapshot, { kind: "checkbox" }> = {
        kind: "checkbox",
      };
      if (rule.checkedValue !== undefined) {
        cloned.checkedValue = rule.checkedValue;
      }
      if (rule.uncheckedValue !== undefined) {
        cloned.uncheckedValue = rule.uncheckedValue;
      }
      return cloned;
    }
    case "whole":
    case "decimal":
    case "date":
    case "time":
    case "textLength":
      return {
        kind: rule.kind,
        operator: rule.operator,
        values: [...rule.values],
      };
  }
}

export function cloneDataValidationRecord(
  record: WorkbookDataValidationRecord,
): WorkbookDataValidationRecord {
  const cloned: WorkbookDataValidationRecord = {
    range: { ...record.range },
    rule: cloneDataValidationRule(record.rule),
  };
  if (record.allowBlank !== undefined) {
    cloned.allowBlank = record.allowBlank;
  }
  if (record.showDropdown !== undefined) {
    cloned.showDropdown = record.showDropdown;
  }
  if (record.promptTitle !== undefined) {
    cloned.promptTitle = record.promptTitle;
  }
  if (record.promptMessage !== undefined) {
    cloned.promptMessage = record.promptMessage;
  }
  if (record.errorStyle !== undefined) {
    cloned.errorStyle = record.errorStyle;
  }
  if (record.errorTitle !== undefined) {
    cloned.errorTitle = record.errorTitle;
  }
  if (record.errorMessage !== undefined) {
    cloned.errorMessage = record.errorMessage;
  }
  return cloned;
}

export function cloneConditionalFormatRule(
  rule: WorkbookConditionalFormatRuleSnapshot,
): WorkbookConditionalFormatRuleSnapshot {
  switch (rule.kind) {
    case "cellIs":
      return {
        kind: "cellIs",
        operator: rule.operator,
        values: [...rule.values],
      };
    case "textContains": {
      const cloned: Extract<WorkbookConditionalFormatRuleSnapshot, { kind: "textContains" }> = {
        kind: "textContains",
        text: rule.text,
      };
      if (rule.caseSensitive !== undefined) {
        cloned.caseSensitive = rule.caseSensitive;
      }
      return cloned;
    }
    case "formula":
      return { kind: "formula", formula: rule.formula };
    case "blanks":
      return { kind: "blanks" };
    case "notBlanks":
      return { kind: "notBlanks" };
  }
}

export function cloneConditionalFormatRecord(
  record: WorkbookConditionalFormatRecord,
): WorkbookConditionalFormatRecord {
  const cloned: WorkbookConditionalFormatRecord = {
    id: record.id,
    range: { ...record.range },
    rule: cloneConditionalFormatRule(record.rule),
    style: structuredClone(record.style),
  };
  if (record.stopIfTrue !== undefined) {
    cloned.stopIfTrue = record.stopIfTrue;
  }
  if (record.priority !== undefined) {
    cloned.priority = record.priority;
  }
  return cloned;
}

export function cloneSheetProtectionRecord(
  record: WorkbookSheetProtectionRecord,
): WorkbookSheetProtectionRecord {
  return {
    sheetName: record.sheetName,
    ...(record.hideFormulas !== undefined ? { hideFormulas: record.hideFormulas } : {}),
  };
}

export function cloneRangeProtectionRecord(
  record: WorkbookRangeProtectionRecord,
): WorkbookRangeProtectionRecord {
  return {
    id: record.id,
    range: { ...record.range },
    ...(record.hideFormulas !== undefined ? { hideFormulas: record.hideFormulas } : {}),
  };
}

export function cloneCommentEntryRecord(
  record: WorkbookCommentEntryRecord,
): WorkbookCommentEntryRecord {
  const cloned: WorkbookCommentEntryRecord = {
    id: record.id,
    body: record.body,
  };
  if (record.authorUserId !== undefined) {
    cloned.authorUserId = record.authorUserId;
  }
  if (record.authorDisplayName !== undefined) {
    cloned.authorDisplayName = record.authorDisplayName;
  }
  if (record.createdAtUnixMs !== undefined) {
    cloned.createdAtUnixMs = record.createdAtUnixMs;
  }
  return cloned;
}

export function cloneCommentThreadRecord(
  record: WorkbookCommentThreadRecord,
): WorkbookCommentThreadRecord {
  const cloned: WorkbookCommentThreadRecord = {
    threadId: record.threadId,
    sheetName: record.sheetName,
    address: canonicalWorkbookAddress(record.sheetName, record.address),
    comments: record.comments.map(cloneCommentEntryRecord),
  };
  if (record.resolved !== undefined) {
    cloned.resolved = record.resolved;
  }
  if (record.resolvedByUserId !== undefined) {
    cloned.resolvedByUserId = record.resolvedByUserId;
  }
  if (record.resolvedAtUnixMs !== undefined) {
    cloned.resolvedAtUnixMs = record.resolvedAtUnixMs;
  }
  return cloned;
}

export function cloneNoteRecord(record: WorkbookNoteRecord): WorkbookNoteRecord {
  return {
    sheetName: record.sheetName,
    address: canonicalWorkbookAddress(record.sheetName, record.address),
    text: record.text,
  };
}

export function conditionalFormatKey(id: string): string {
  const normalized = id.trim();
  if (normalized.length === 0) {
    throw new Error("Conditional format id must be non-empty");
  }
  return normalized;
}

export function rangeProtectionKey(id: string): string {
  const normalized = id.trim();
  if (normalized.length === 0) {
    throw new Error("Range protection id must be non-empty");
  }
  return normalized;
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

export function dataValidationKey(sheetName: string, range: CellRangeRef): string {
  const normalized = canonicalWorkbookRangeRef(range);
  return `${sheetName}:${normalized.startAddress}:${normalized.endAddress}`;
}

export function commentThreadKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`;
}

export function noteKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`;
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
  if (isSheetProtectionRecord(record)) {
    return record.sheetName;
  }
  if (isFilterRecord(record)) {
    return filterKey(record.sheetName, record.range);
  }
  if (isSortRecord(record)) {
    return sortKey(record.sheetName, record.range);
  }
  if (isConditionalFormatRecord(record)) {
    return conditionalFormatKey(record.id);
  }
  if (isRangeProtectionRecord(record)) {
    return rangeProtectionKey(record.id);
  }
  if (isDataValidationRecord(record)) {
    return dataValidationKey(record.range.sheetName, record.range);
  }
  if (isCommentThreadRecord(record)) {
    return commentThreadKey(record.sheetName, record.address);
  }
  if (isNoteRecord(record)) {
    return noteKey(record.sheetName, record.address);
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

function isSheetProtectionRecord(record: unknown): record is WorkbookSheetProtectionRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "sheetName" in record &&
    !("rows" in record) &&
    !("cols" in record) &&
    !("range" in record) &&
    !("startAddress" in record) &&
    !("address" in record) &&
    !("start" in record)
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

function isDataValidationRecord(record: unknown): record is WorkbookDataValidationRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "range" in record &&
    "rule" in record &&
    !("keys" in record) &&
    !("id" in record && "style" in record)
  );
}

function isCommentThreadRecord(record: unknown): record is WorkbookCommentThreadRecord {
  return (
    typeof record === "object" && record !== null && "threadId" in record && "comments" in record
  );
}

function isConditionalFormatRecord(record: unknown): record is WorkbookConditionalFormatRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "id" in record &&
    "range" in record &&
    "rule" in record &&
    "style" in record
  );
}

function isRangeProtectionRecord(record: unknown): record is WorkbookRangeProtectionRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "id" in record &&
    "range" in record &&
    !("rule" in record) &&
    !("style" in record)
  );
}

function isNoteRecord(record: unknown): record is WorkbookNoteRecord {
  return (
    typeof record === "object" &&
    record !== null &&
    "address" in record &&
    "text" in record &&
    !("comments" in record)
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
