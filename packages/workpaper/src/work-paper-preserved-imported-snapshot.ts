import type { WorkbookSnapshot } from '@bilig/protocol'
import { attachWorkPaperRuntimeImage } from './work-paper-snapshot-clone.js'
import {
  attachImportedXlsxSourceMetadata,
  readImportedXlsxSource,
  readImportedXlsxSourceCellPatches,
} from './work-paper-imported-xlsx-source.js'

type WorkbookSnapshotWorkbook = WorkbookSnapshot['workbook']
type WorkbookSnapshotSheetMetadata = NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>

function isCloneRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function cloneSnapshotMetadataValue(value: unknown): unknown {
  if (!isCloneRecord(value)) {
    return value
  }
  if (Array.isArray(value)) {
    return value.map(cloneSnapshotMetadataValue)
  }
  if (ArrayBuffer.isView(value)) {
    return structuredClone(value)
  }

  const path = value['path']
  if (typeof path === 'string') {
    const storage = value['storage']
    const dataBase64 = value['dataBase64']
    const byteLength = value['byteLength']
    if (storage === 'base64' && typeof dataBase64 === 'string' && typeof byteLength === 'number') {
      return { path, storage, dataBase64, byteLength }
    }

    const xml = value['xml']
    if (typeof xml === 'string' && typeof value['readXml'] === 'function') {
      return { path, xml }
    }
  }

  const cloned: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    const child = value[key]
    if (typeof child !== 'function') {
      cloned[key] = cloneSnapshotMetadataValue(child)
    }
  }
  return cloned
}

function isWorkbookSnapshotWorkbook(value: unknown): value is WorkbookSnapshotWorkbook {
  return isCloneRecord(value) && typeof value['name'] === 'string'
}

function cloneWorkbookSnapshotWorkbook(workbook: WorkbookSnapshotWorkbook): WorkbookSnapshotWorkbook {
  const cloned = cloneSnapshotMetadataValue(workbook)
  return isWorkbookSnapshotWorkbook(cloned) ? cloned : { name: workbook.name }
}

function isWorkbookSnapshotSheetMetadata(value: unknown): value is WorkbookSnapshotSheetMetadata {
  return isCloneRecord(value)
}

function cloneWorkbookSnapshotSheetMetadata(metadata: WorkbookSnapshotSheetMetadata): WorkbookSnapshotSheetMetadata {
  const cloned = cloneSnapshotMetadataValue(metadata)
  return isWorkbookSnapshotSheetMetadata(cloned) ? cloned : {}
}

export function clonePreservedImportedSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const cloned: WorkbookSnapshot = {
    version: snapshot.version,
    workbook: cloneWorkbookSnapshotWorkbook(snapshot.workbook),
    sheets: snapshot.sheets.map((sheet) => ({
      ...sheet,
      ...(sheet.metadata === undefined ? {} : { metadata: cloneWorkbookSnapshotSheetMetadata(sheet.metadata) }),
      cells: sheet.cells,
    })),
  }
  const sourceBytes = readImportedXlsxSource(snapshot)
  const sourcePatches = readImportedXlsxSourceCellPatches(snapshot)
  attachImportedXlsxSourceMetadata(cloned, sourceBytes, sourcePatches)
  return attachWorkPaperRuntimeImage(snapshot, cloned)
}
