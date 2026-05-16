import type { LiteralInput, WorkbookProtectionSnapshot } from '@bilig/protocol'
import { clonePropertyRecord, cloneWorkbookProtectionRecord } from './workbook-metadata-records.js'
import { normalizeMetadataKey } from './workbook-metadata-service-helpers.js'
import type { WorkbookMetadataRecord, WorkbookPropertyRecord, WorkbookProtectionRecord } from './workbook-metadata-types.js'

export function setWorkbookPropertyRecord(
  metadata: WorkbookMetadataRecord,
  key: string,
  value: LiteralInput,
): WorkbookPropertyRecord | undefined {
  const trimmedKey = normalizeMetadataKey(key)
  if (value === null) {
    metadata.properties.delete(trimmedKey)
    return undefined
  }
  const record: WorkbookPropertyRecord = { key: trimmedKey, value }
  metadata.properties.set(trimmedKey, record)
  return clonePropertyRecord(record)
}

export function getWorkbookPropertyRecord(metadata: WorkbookMetadataRecord, key: string): WorkbookPropertyRecord | undefined {
  const record = metadata.properties.get(normalizeMetadataKey(key))
  return record ? clonePropertyRecord(record) : undefined
}

export function listWorkbookPropertyRecords(metadata: WorkbookMetadataRecord): WorkbookPropertyRecord[] {
  return [...metadata.properties.values()].toSorted((left, right) => left.key.localeCompare(right.key)).map(clonePropertyRecord)
}

export function setWorkbookProtectionRecord(
  metadata: WorkbookMetadataRecord,
  record: WorkbookProtectionSnapshot,
): WorkbookProtectionRecord {
  const stored: WorkbookProtectionRecord = cloneWorkbookProtectionRecord(record)
  metadata.workbookProtection = stored
  return cloneWorkbookProtectionRecord(stored)
}

export function getWorkbookProtectionRecord(metadata: WorkbookMetadataRecord): WorkbookProtectionRecord | undefined {
  return metadata.workbookProtection ? cloneWorkbookProtectionRecord(metadata.workbookProtection) : undefined
}
