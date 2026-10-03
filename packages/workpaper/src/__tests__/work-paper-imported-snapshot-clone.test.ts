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
          workbookPivotCachesXml: '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches>',
          workbookRelationships: [
            {
              id: 'rIdPivotCache1',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
              target: 'pivotCache/pivotCacheDefinition1.xml',
            },
          ],
        },
        documentPropertyArtifacts: {
          core: {
            path: 'docProps/core.xml',
            xml: '<cp:coreProperties><dc:title>Imported</dc:title></cp:coreProperties>',
            relationship: {
              id: 'rIdCore',
              type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
              target: 'docProps/core.xml',
            },
            contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
          },
        },
        dataModelArtifacts: {
          parts: [{ path: 'xl/model/item.data', storage: 'base64', dataBase64: 'BAUG', byteLength: 3 }],
          workbookRelationships: [
            {
              id: 'rIdModel',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/model',
              target: 'model/item.data',
            },
          ],
        },
        slicerConnectionArtifacts: {
          parts: [{ path: 'xl/slicerCaches/slicerCache1.xml', storage: 'base64', dataBase64: 'BwgJ', byteLength: 3 }],
          workbookRelationships: [
            {
              id: 'rIdSlicer',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slicerCache',
              target: 'slicerCaches/slicerCache1.xml',
            },
          ],
          sheetArtifacts: [
            { sheetName: 'Sheet1', sheetSlicerListExtXml: '<extLst><ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"/></extLst>' },
          ],
        },
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="0"/></bookViews>',
        },
        styleArtifacts: {
          stylesXml: '<styleSheet><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',
          theme: {
            path: 'xl/theme/theme1.xml',
            xml: '<a:theme name="Office"/>',
            relationship: {
              id: 'rIdTheme',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
              target: 'theme/theme1.xml',
            },
          },
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        metadata: {
          drawingArtifacts: {
            relationshipTarget: '../drawings/drawing1.xml',
            preservedChartRelationshipIds: ['rId7'],
          },
          pivotArtifacts: {
            relationships: [
              {
                id: 'rIdPivotTable1',
                type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
                target: '../pivotTables/pivotTable1.xml',
              },
            ],
            pivotTableDefinitionsXml: '<pivotTableDefinition/>',
          },
          viewState: {
            sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
          },
          styleArtifacts: {
            cellStyleIndexes: [{ address: 'A1', styleIndex: 1 }],
            blankCellAddresses: ['C3'],
          },
        },
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
      expect(exported.sheets[0]?.metadata?.drawingArtifacts).toEqual({
        relationshipTarget: '../drawings/drawing1.xml',
        preservedChartRelationshipIds: ['rId7'],
      })
      expect(exported.workbook.metadata?.viewState).toEqual({ bookViewsXml: '<bookViews><workbookView activeTab="0"/></bookViews>' })
      expect(exported.sheets[0]?.metadata?.viewState).toEqual({
        sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
      })
      expect(exported.workbook.metadata?.documentPropertyArtifacts?.core?.path).toBe('docProps/core.xml')
      expect(exported.workbook.metadata?.dataModelArtifacts?.parts[0]).toEqual({
        path: 'xl/model/item.data',
        storage: 'base64',
        dataBase64: 'BAUG',
        byteLength: 3,
      })
      expect(exported.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts?.[0]?.sheetName).toBe('Sheet1')
      expect(exported.workbook.metadata?.styleArtifacts?.theme?.path).toBe('xl/theme/theme1.xml')
      expect(exported.sheets[0]?.metadata?.styleArtifacts).toEqual({
        cellStyleIndexes: [{ address: 'A1', styleIndex: 1 }],
        blankCellAddresses: ['C3'],
      })
    } finally {
      workbook.dispose()
    }
  })

  it('keeps imported package metadata after the WorkPaper snapshot is rebuilt by a headless edit', () => {
    const workbook = WorkPaper.buildFromSnapshot(importedSnapshotWithLazyArtifacts())
    try {
      workbook.setCellContents({ sheet: 1, row: 0, col: 1 }, 'headless edit')

      const exported = workbook.exportSnapshot()
      const drawingPart = exported.workbook.metadata?.drawingArtifacts?.parts[0]

      expect(drawingPart).toEqual({
        path: 'xl/drawings/drawing1.xml',
        storage: 'base64',
        dataBase64: 'AQID',
        byteLength: 3,
      })
      expect(Object.hasOwn(drawingPart ?? {}, 'readBytes')).toBe(false)
      expect(exported.sheets[0]?.metadata?.drawingArtifacts).toEqual({
        relationshipTarget: '../drawings/drawing1.xml',
        preservedChartRelationshipIds: ['rId7'],
      })
      expect(exported.workbook.metadata?.pivotArtifacts?.parts[0]).toEqual({
        path: 'xl/pivotTables/pivotTable1.xml',
        xml: '<pivotTableDefinition/>',
      })
      expect(Object.hasOwn(exported.workbook.metadata?.pivotArtifacts?.parts[0] ?? {}, 'readXml')).toBe(false)
      expect(exported.workbook.metadata?.viewState).toEqual({ bookViewsXml: '<bookViews><workbookView activeTab="0"/></bookViews>' })
      expect(exported.sheets[0]?.metadata?.pivotArtifacts).toEqual({
        relationships: [
          {
            id: 'rIdPivotTable1',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
            target: '../pivotTables/pivotTable1.xml',
          },
        ],
        pivotTableDefinitionsXml: '<pivotTableDefinition/>',
      })
      expect(exported.sheets[0]?.metadata?.viewState).toEqual({
        sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
      })
      expect(exported.workbook.metadata?.documentPropertyArtifacts?.core?.path).toBe('docProps/core.xml')
      expect(exported.workbook.metadata?.dataModelArtifacts?.parts[0]).toEqual({
        path: 'xl/model/item.data',
        storage: 'base64',
        dataBase64: 'BAUG',
        byteLength: 3,
      })
      expect(exported.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts?.[0]?.sheetName).toBe('Sheet1')
      expect(exported.workbook.metadata?.styleArtifacts?.theme?.path).toBe('xl/theme/theme1.xml')
      expect(exported.sheets[0]?.metadata?.styleArtifacts).toEqual({
        cellStyleIndexes: [{ address: 'A1', styleIndex: 1 }],
        blankCellAddresses: ['C3'],
      })
    } finally {
      workbook.dispose()
    }
  })
})
