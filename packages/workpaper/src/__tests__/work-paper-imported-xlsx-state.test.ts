import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  attachImportedXlsxSourceMetadata,
  readImportedXlsxSource,
  readImportedXlsxSourceCellPatches,
} from '../work-paper-imported-xlsx-source.js'
import { WorkPaperImportedXlsxState } from '../work-paper-imported-xlsx-state.js'

function emptySnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Imported' },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        cells: [{ address: 'A1', value: 1 }],
      },
    ],
  }
}

describe('WorkPaperImportedXlsxState', () => {
  it('owns imported source release, active state, and source-preserving literal patches', () => {
    let releaseCount = 0
    const source = {
      byteLength: 3,
      readBytes: () => new Uint8Array([1, 2, 3]),
      release: () => {
        releaseCount += 1
      },
    }
    const snapshot = attachImportedXlsxSourceMetadata(emptySnapshot(), source)
    const state = new WorkPaperImportedXlsxState()

    state.initializeFromSnapshot(snapshot)

    expect(state.hasState).toBe(true)
    expect(state.isActive).toBe(true)
    expect(snapshot.sheets[0]?.cells).toEqual([])

    expect(state.prepareForCellContentEdit('patched')).toBe(true)
    expect(state.isRecordingSourcePreservingEdit).toBe(true)
    state.recordLiteralPatch('Sheet1', 'A1', 'patched')
    state.finishCellContentEdit()

    const exported = state.exportSourcePreservingSnapshot({
      version: 1,
      workbook: { name: 'Imported' },
      sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }],
    })

    expect(exported).not.toBeNull()
    expect(readImportedXlsxSource(exported!)).toBe(source)
    expect(readImportedXlsxSourceCellPatches(exported!)).toEqual([
      {
        kind: 'literal',
        sheetName: 'Sheet1',
        address: 'A1',
        value: 'patched',
      },
    ])

    expect(state.prepareForCellContentEdit('=A1')).toBe(false)
    state.finishCellContentEdit()

    expect(state.hasState).toBe(false)
    expect(state.isActive).toBe(false)
    expect(releaseCount).toBe(1)

    state.dispose()

    expect(releaseCount).toBe(1)
  })
})
