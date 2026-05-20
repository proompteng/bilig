import { describe, expect, it } from 'vitest'
import type { WorkbookPivotPackagePartSnapshot, WorkbookPreservedPackagePartSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'

type LazyBase64PackagePartSnapshot = WorkbookPreservedPackagePartSnapshot & {
  readBytes(): Uint8Array
}

type LazyXmlPackagePartSnapshot = WorkbookPivotPackagePartSnapshot & {
  readXml(): string
}

function lazyBase64PackagePart(): LazyBase64PackagePartSnapshot {
  return {
    path: 'xl/drawings/drawing1.xml',
    storage: 'base64',
    byteLength: 3,
    readBytes: () => new Uint8Array([1, 2, 3]),
    dataBase64: 'AQID',
  }
}

function lazyXmlPackagePart(): LazyXmlPackagePartSnapshot {
  return {
    path: 'xl/pivotTables/pivotTable1.xml',
    readXml: () => '<pivotTableDefinition/>',
    xml: '<pivotTableDefinition/>',
  }
}

function importedSnapshotWithLazyArtifacts(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Imported',
      metadata: {
        drawingArtifacts: {
          parts: [lazyBase64PackagePart()],
        },
        pivotArtifacts: {
          parts: [lazyXmlPackagePart()],
        },
      },
    },
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

describe('WorkPaper imported snapshot cloning', () => {
  it('materializes lazy XLSX artifact parts before exporting a preserved imported snapshot', () => {
    const workbook = WorkPaper.buildFromSnapshot(importedSnapshotWithLazyArtifacts())
    try {
      const exported = workbook.exportSnapshot()
      const drawingPart = exported.workbook.metadata?.drawingArtifacts?.parts[0]
      const pivotPart = exported.workbook.metadata?.pivotArtifacts?.parts[0]

      expect(drawingPart).toEqual({
        path: 'xl/drawings/drawing1.xml',
        storage: 'base64',
        dataBase64: 'AQID',
        byteLength: 3,
      })
      expect(Object.hasOwn(drawingPart ?? {}, 'readBytes')).toBe(false)
      expect(pivotPart).toEqual({
        path: 'xl/pivotTables/pivotTable1.xml',
        xml: '<pivotTableDefinition/>',
      })
      expect(Object.hasOwn(pivotPart ?? {}, 'readXml')).toBe(false)
    } finally {
      workbook.dispose()
    }
  })
})
