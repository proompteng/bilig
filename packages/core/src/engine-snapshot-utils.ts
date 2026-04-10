import type {
  CellRangeRef,
  SheetMetadataSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookFreezePaneSnapshot,
} from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";
import type {
  WorkbookAxisMetadataRecord,
  WorkbookFreezePaneRecord,
  WorkbookStore,
} from "./workbook-store.js";

function cloneSnapshotRangeRef(range: CellRangeRef): CellRangeRef {
  return {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  };
}

function axisMetadataToSnapshot(
  records: readonly WorkbookAxisMetadataRecord[],
): WorkbookAxisMetadataSnapshot[] {
  return records.map((record) => {
    const snapshot: WorkbookAxisMetadataSnapshot = {
      start: record.start,
      count: record.count,
    };
    if (record.size !== null) {
      snapshot.size = record.size;
    }
    if (record.hidden !== null) {
      snapshot.hidden = record.hidden;
    }
    return snapshot;
  });
}

function freezePaneToSnapshot(
  record: WorkbookFreezePaneRecord | undefined,
): WorkbookFreezePaneSnapshot | undefined {
  if (!record) {
    return undefined;
  }
  return { rows: record.rows, cols: record.cols };
}

export function exportSheetMetadata(
  workbook: WorkbookStore,
  sheetName: string,
): SheetMetadataSnapshot | undefined {
  const rows = workbook.listRowAxisEntries(sheetName);
  const columns = workbook.listColumnAxisEntries(sheetName);
  const rowMetadata = axisMetadataToSnapshot(workbook.listRowMetadata(sheetName));
  const columnMetadata = axisMetadataToSnapshot(workbook.listColumnMetadata(sheetName));
  const styleRanges = workbook.listStyleRanges(sheetName).map((record) => ({
    range: cloneSnapshotRangeRef(record.range),
    styleId: record.styleId,
  }));
  const formatRanges = workbook.listFormatRanges(sheetName).map((record) => ({
    range: cloneSnapshotRangeRef(record.range),
    formatId: record.formatId,
  }));
  const freezePane = freezePaneToSnapshot(workbook.getFreezePane(sheetName));
  const filters = workbook.listFilters(sheetName).map((filter) => Object.assign({}, filter.range));
  const sorts = workbook.listSorts(sheetName).map((sort) => ({
    range: { ...sort.range },
    keys: sort.keys.map((key) => ({ ...key })),
  }));

  if (
    rows.length === 0 &&
    columns.length === 0 &&
    rowMetadata.length === 0 &&
    columnMetadata.length === 0 &&
    styleRanges.length === 0 &&
    formatRanges.length === 0 &&
    freezePane === undefined &&
    filters.length === 0 &&
    sorts.length === 0
  ) {
    return undefined;
  }

  const metadata: SheetMetadataSnapshot = {};
  if (rows.length > 0) {
    metadata.rows = rows;
  }
  if (columns.length > 0) {
    metadata.columns = columns;
  }
  if (rowMetadata.length > 0) {
    metadata.rowMetadata = rowMetadata;
  }
  if (columnMetadata.length > 0) {
    metadata.columnMetadata = columnMetadata;
  }
  if (styleRanges.length > 0) {
    metadata.styleRanges = styleRanges;
  }
  if (formatRanges.length > 0) {
    metadata.formatRanges = formatRanges;
  }
  if (freezePane) {
    metadata.freezePane = freezePane;
  }
  if (filters.length > 0) {
    metadata.filters = filters;
  }
  if (sorts.length > 0) {
    metadata.sorts = sorts;
  }
  return metadata;
}

export function sheetMetadataToOps(workbook: WorkbookStore, sheetName: string): EngineOp[] {
  const ops: EngineOp[] = [];
  workbook.listRowAxisEntries(sheetName).forEach((entry) => {
    ops.push({ kind: "insertRows", sheetName, start: entry.index, count: 1, entries: [entry] });
  });
  workbook.listColumnAxisEntries(sheetName).forEach((entry) => {
    ops.push({
      kind: "insertColumns",
      sheetName,
      start: entry.index,
      count: 1,
      entries: [entry],
    });
  });
  workbook.listRowMetadata(sheetName).forEach((record) => {
    ops.push({
      kind: "updateRowMetadata",
      sheetName,
      start: record.start,
      count: record.count,
      size: record.size,
      hidden: record.hidden,
    });
  });
  workbook.listColumnMetadata(sheetName).forEach((record) => {
    ops.push({
      kind: "updateColumnMetadata",
      sheetName,
      start: record.start,
      count: record.count,
      size: record.size,
      hidden: record.hidden,
    });
  });
  workbook.listStyleRanges(sheetName).forEach((record) => {
    ops.push({ kind: "setStyleRange", range: { ...record.range }, styleId: record.styleId });
  });
  workbook.listFormatRanges(sheetName).forEach((record) => {
    ops.push({ kind: "setFormatRange", range: { ...record.range }, formatId: record.formatId });
  });
  const freezePane = workbook.getFreezePane(sheetName);
  if (freezePane) {
    ops.push({ kind: "setFreezePane", sheetName, rows: freezePane.rows, cols: freezePane.cols });
  }
  workbook.listFilters(sheetName).forEach((record) => {
    ops.push({ kind: "setFilter", sheetName, range: { ...record.range } });
  });
  workbook.listSorts(sheetName).forEach((record) => {
    ops.push({
      kind: "setSort",
      sheetName,
      range: { ...record.range },
      keys: record.keys.map((key) => ({ ...key })),
    });
  });
  return ops;
}
