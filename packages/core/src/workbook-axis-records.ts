import type { WorkbookAxisEntrySnapshot } from "@bilig/protocol";
import type {
  WorkbookAxisEntryRecord,
  WorkbookAxisMetadataRecord,
} from "./workbook-metadata-types.js";
import { axisMetadataKey, deleteRecordsBySheet } from "./workbook-store-records.js";

function makeAxisEntrySnapshot(
  entry: WorkbookAxisEntryRecord,
  index: number,
): WorkbookAxisEntrySnapshot {
  const snapshot: WorkbookAxisEntrySnapshot = { id: entry.id, index };
  if (entry.size !== null) {
    snapshot.size = entry.size;
  }
  if (entry.hidden !== null) {
    snapshot.hidden = entry.hidden;
  }
  return snapshot;
}

export function listAxisEntries(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
): WorkbookAxisEntrySnapshot[] {
  const result: WorkbookAxisEntrySnapshot[] = [];
  entries.forEach((entry, index) => {
    if (!entry) {
      return;
    }
    result.push(makeAxisEntrySnapshot(entry, index));
  });
  return result;
}

export function materializeAxisEntryRecords(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  createEntry: () => WorkbookAxisEntryRecord,
): WorkbookAxisEntryRecord[] {
  const materialized: WorkbookAxisEntryRecord[] = [];
  for (let index = 0; index < count; index += 1) {
    const position = start + index;
    let entry = entries[position];
    if (!entry) {
      entry = createEntry();
      entries[position] = entry;
    }
    materialized.push(entry);
  }
  return materialized;
}

export function materializeAxisEntries(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  createEntry: () => WorkbookAxisEntryRecord,
): WorkbookAxisEntrySnapshot[] {
  return materializeAxisEntryRecords(entries, start, count, createEntry).map((entry, offset) =>
    makeAxisEntrySnapshot(entry, start + offset),
  );
}

export function snapshotAxisEntriesInRange(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
): WorkbookAxisEntrySnapshot[] {
  if (count <= 0) {
    return [];
  }
  const snapshots: WorkbookAxisEntrySnapshot[] = [];
  for (let offset = 0; offset < count; offset += 1) {
    const index = start + offset;
    const entry = entries[index];
    if (!entry) {
      continue;
    }
    snapshots.push(makeAxisEntrySnapshot(entry, index));
  }
  return snapshots;
}

export function spliceAxisEntries(
  axisEntries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  deleteCount: number,
  insertCount: number,
  createEntry: () => WorkbookAxisEntryRecord,
  providedSnapshots?: readonly WorkbookAxisEntrySnapshot[],
): WorkbookAxisEntrySnapshot[] {
  const providedEntries = new Map<number, WorkbookAxisEntrySnapshot>();
  providedSnapshots?.forEach((entry) => {
    const offset = entry.index - start;
    if (offset < 0 || offset >= insertCount) {
      return;
    }
    providedEntries.set(offset, entry);
  });
  if (axisEntries.length < start) {
    axisEntries.length = start;
  }
  if (deleteCount > 0) {
    materializeAxisEntryRecords(axisEntries, start, deleteCount, createEntry);
  }
  const removed = axisEntries.splice(
    start,
    deleteCount,
    ...Array.from({ length: insertCount }, (_, index) => {
      const provided = providedEntries.get(index);
      if (provided) {
        return { id: provided.id, size: provided.size ?? null, hidden: provided.hidden ?? null };
      }
      if (providedSnapshots) {
        return undefined;
      }
      return createEntry();
    }),
  );
  return removed.flatMap((entry, index) =>
    entry ? [makeAxisEntrySnapshot(entry, start + index)] : [],
  );
}

export function moveAxisEntries(
  axisEntries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  target: number,
  createEntry: () => WorkbookAxisEntryRecord,
): void {
  if (count <= 0 || start === target) {
    return;
  }
  materializeAxisEntryRecords(axisEntries, start, count, createEntry);
  const moved = axisEntries.splice(start, count);
  axisEntries.splice(target, 0, ...moved);
}

export function getAxisMetadataRecord(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  sheetName: string,
  start: number,
  count: number,
): WorkbookAxisMetadataRecord | undefined {
  let size: number | null | undefined;
  let hidden: boolean | null | undefined;
  let sawMaterialized = false;
  for (let index = start; index < start + count; index += 1) {
    const entry = entries[index];
    if (!entry) {
      if (size === undefined) {
        size = null;
      }
      if (hidden === undefined) {
        hidden = null;
      }
      continue;
    }
    sawMaterialized = true;
    size ??= entry.size;
    hidden ??= entry.hidden;
    if (size !== entry.size || hidden !== entry.hidden) {
      return undefined;
    }
  }
  if (!sawMaterialized || ((size ?? null) === null && (hidden ?? null) === null)) {
    return undefined;
  }
  return { sheetName, start, count, size: size ?? null, hidden: hidden ?? null };
}

export function syncAxisMetadataBucket(
  bucket: Map<string, WorkbookAxisMetadataRecord>,
  sheetName: string,
  entries: Array<WorkbookAxisEntryRecord | undefined>,
): void {
  deleteRecordsBySheet(bucket, sheetName, (record) => record.sheetName);
  let cursor = 0;
  while (cursor < entries.length) {
    const entry = entries[cursor];
    if (!entry || (entry.size === null && entry.hidden === null)) {
      cursor += 1;
      continue;
    }
    const start = cursor;
    const size = entry.size;
    const hidden = entry.hidden;
    cursor += 1;
    while (cursor < entries.length) {
      const next = entries[cursor];
      if (!next || next.size !== size || next.hidden !== hidden) {
        break;
      }
      cursor += 1;
    }
    const record: WorkbookAxisMetadataRecord = {
      sheetName,
      start,
      count: cursor - start,
      size,
      hidden,
    };
    bucket.set(axisMetadataKey(sheetName, start, record.count), record);
  }
}
