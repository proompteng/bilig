import { ValueTag, type CellValue, type LiteralInput, type WorkbookSnapshot } from '@bilig/protocol'
import type { WorkPaperChange, WorkPaperCellAddress } from './work-paper-types.js'
import {
  attachImportedXlsxSourceMetadata,
  canRecordImportedXlsxLiteralPatch,
  readImportedXlsxSource,
  readImportedXlsxSourceCellPatches,
  releaseImportedXlsxSourceReaderSnapshotCells,
  setImportedXlsxFormulaCachePatch,
  setImportedXlsxLiteralPatch,
  type ImportedXlsxSourceCellPatch,
  type ImportedXlsxSourceReference,
} from './work-paper-imported-xlsx-source.js'
import { clonePreservedImportedSnapshot } from './work-paper-preserved-imported-snapshot.js'

function importedXlsxPatchValueFromCellValue(value: CellValue): LiteralInput | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return undefined
  }
}

export class WorkPaperImportedXlsxState {
  private preservedImportedSnapshot: WorkbookSnapshot | undefined
  private importedXlsxSource: ImportedXlsxSourceReference | undefined
  private importedXlsxStateActive = false
  private readonly importedXlsxSourceCellPatches = new Map<string, ImportedXlsxSourceCellPatch>()
  private recordingSourcePreservingImportedXlsxEdit = false

  get hasState(): boolean {
    return (
      this.preservedImportedSnapshot !== undefined || this.importedXlsxSource !== undefined || this.importedXlsxSourceCellPatches.size > 0
    )
  }

  get isActive(): boolean {
    return this.importedXlsxStateActive
  }

  get isRecordingSourcePreservingEdit(): boolean {
    return this.recordingSourcePreservingImportedXlsxEdit
  }

  initializeFromSnapshot(snapshot: WorkbookSnapshot): void {
    this.invalidate()
    const importedXlsxSource = readImportedXlsxSource(snapshot)
    if (importedXlsxSource === undefined) {
      this.preservedImportedSnapshot = snapshot
    } else {
      this.importedXlsxSource = importedXlsxSource
      this.importedXlsxStateActive = true
    }
    for (const patch of readImportedXlsxSourceCellPatches(snapshot)) {
      this.importedXlsxSourceCellPatches.set(`${patch.sheetName}!${patch.address}`, patch)
    }
    releaseImportedXlsxSourceReaderSnapshotCells(snapshot, importedXlsxSource)
  }

  tryExportSnapshot(createSourceSnapshot: () => WorkbookSnapshot): WorkbookSnapshot | null {
    if (this.preservedImportedSnapshot !== undefined) {
      return clonePreservedImportedSnapshot(this.preservedImportedSnapshot)
    }
    if (this.importedXlsxSource !== undefined) {
      return attachImportedXlsxSourceMetadata(createSourceSnapshot(), this.importedXlsxSource, [
        ...this.importedXlsxSourceCellPatches.values(),
      ])
    }
    return null
  }

  exportSourcePreservingSnapshot(baseSnapshot: WorkbookSnapshot): WorkbookSnapshot | null {
    if (this.importedXlsxSource === undefined || this.importedXlsxSourceCellPatches.size === 0) {
      return null
    }
    return attachImportedXlsxSourceMetadata(baseSnapshot, this.importedXlsxSource, [...this.importedXlsxSourceCellPatches.values()])
  }

  recordFormulaCachePatches(
    changes: readonly WorkPaperChange[],
    getCellFormula: (address: WorkPaperCellAddress) => string | undefined,
  ): void {
    if (this.importedXlsxSource === undefined) {
      return
    }
    for (const change of changes) {
      if (change.kind !== 'cell' || getCellFormula(change.address) === undefined) {
        continue
      }
      const value = importedXlsxPatchValueFromCellValue(change.newValue)
      if (value === undefined) {
        continue
      }
      setImportedXlsxFormulaCachePatch(this.importedXlsxSourceCellPatches, change.sheetName, change.a1, value)
    }
  }

  prepareForCellContentEdit(content: unknown): content is LiteralInput {
    const shouldRecordPatch = canRecordImportedXlsxLiteralPatch(this.importedXlsxSource, content)
    if (shouldRecordPatch) {
      this.preservedImportedSnapshot = undefined
    } else {
      this.invalidate()
    }
    this.recordingSourcePreservingImportedXlsxEdit = shouldRecordPatch
    return shouldRecordPatch
  }

  recordLiteralPatch(sheetName: string, address: string, value: LiteralInput): void {
    this.preservedImportedSnapshot = undefined
    setImportedXlsxLiteralPatch(this.importedXlsxSourceCellPatches, sheetName, address, value)
  }

  finishCellContentEdit(): void {
    this.recordingSourcePreservingImportedXlsxEdit = false
  }

  beforeCapturedMutation(): void {
    if (!this.hasState) {
      return
    }
    if (this.recordingSourcePreservingImportedXlsxEdit) {
      this.preservedImportedSnapshot = undefined
    } else {
      this.invalidate()
    }
  }

  invalidate(): void {
    this.preservedImportedSnapshot = undefined
    this.releaseSource()
    this.importedXlsxStateActive = false
    this.importedXlsxSourceCellPatches.clear()
  }

  dispose(): void {
    this.releaseSource()
    this.importedXlsxSourceCellPatches.clear()
    this.preservedImportedSnapshot = undefined
    this.importedXlsxStateActive = false
    this.recordingSourcePreservingImportedXlsxEdit = false
  }

  private releaseSource(): void {
    if (this.importedXlsxSource !== undefined && !(this.importedXlsxSource instanceof Uint8Array)) {
      this.importedXlsxSource.release?.()
    }
    this.importedXlsxSource = undefined
  }
}
