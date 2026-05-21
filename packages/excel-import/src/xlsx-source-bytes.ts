import type { WorkbookSnapshot } from '@bilig/protocol'

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')

export interface ImportedXlsxSourceReader {
  readonly byteLength: number
  readBytes(): Uint8Array
}

type ImportedXlsxSourceReference = Uint8Array | ImportedXlsxSourceReader

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: ImportedXlsxSourceReference
}

type MutableSnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  [importedXlsxSourceBytes]?: ImportedXlsxSourceReference
}

export function attachImportedXlsxSourceBytes(snapshot: WorkbookSnapshot, bytes: Uint8Array): WorkbookSnapshot {
  return attachImportedXlsxSourceReference(snapshot, bytes)
}

export function attachImportedXlsxSourceReader(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReader): WorkbookSnapshot {
  return attachImportedXlsxSourceReference(snapshot, source)
}

function attachImportedXlsxSourceReference(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReference): WorkbookSnapshot {
  Object.defineProperty(snapshot, importedXlsxSourceBytes, {
    configurable: true,
    enumerable: false,
    value: source,
  })
  return snapshot
}

export function readImportedXlsxSourceBytes(snapshot: WorkbookSnapshot): Uint8Array | undefined {
  const source = (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  return source instanceof Uint8Array ? source : source?.readBytes()
}

export function detachImportedXlsxSourceBytes(snapshot: WorkbookSnapshot): boolean {
  if (!Object.prototype.hasOwnProperty.call(snapshot, importedXlsxSourceBytes)) {
    return false
  }
  delete (snapshot as MutableSnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  return true
}
