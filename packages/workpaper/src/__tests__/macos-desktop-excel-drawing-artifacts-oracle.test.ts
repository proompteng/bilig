import { Buffer } from 'node:buffer'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookPreservedPackagePartSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const imageContentType = 'image/png'

describe('macOS Desktop Excel drawing artifact oracle', () => {
  it('preserves imported drawing package parts after a headless edit', () => {
    const imported = importXlsx(exportXlsx(workbookWithEmbeddedImageDrawing()), 'drawing-artifact-source.xlsx')
    const workpaper = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      const sheet = workpaper.getSheetId('Sheet1')
      if (sheet === undefined) {
        throw new Error('Expected Sheet1 to be available')
      }

      workpaper.setCellContents({ sheet, row: 0, col: 1 }, 'headless edit')

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'headless-drawing-artifact-roundtrip.xlsx')

      expect(drawingArtifactPaths(reimported.snapshot)).toEqual(
        expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png']),
      )
      expect(reimported.snapshot.sheets[0]?.metadata?.drawingArtifacts).toEqual({ relationshipTarget: '../drawings/drawing1.xml' })
      expect(strFromU8(unzipSync(exported)['xl/drawings/drawing1.xml'] ?? new Uint8Array())).toContain('<xdr:pic>')
    } finally {
      workpaper.dispose()
    }
  })

  it('structurally rewrites raw DrawingML anchors after a headless row insert', () => {
    const imported = importXlsx(exportXlsx(workbookWithEmbeddedImageDrawing()), 'drawing-artifact-structure-source.xlsx')
    const workpaper = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      const sheet = workpaper.getSheetId('Sheet1')
      if (sheet === undefined) {
        throw new Error('Expected Sheet1 to be available')
      }

      workpaper.addRows(sheet, 0, 1)

      const headlessSnapshot = workpaper.exportSnapshot()
      expect(drawingAnchorRows(headlessSnapshot)).toEqual([1, 5])
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel drawing package parts after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-drawing-artifacts-oracle-')
      try {
        const sourcePath = join(tempDir, 'drawing-artifact-source.xlsx')
        writeFileSync(sourcePath, exportXlsx(workbookWithEmbeddedImageDrawing()))

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Sheet1',
          formulaCells: [],
          inspectCells: ['A1', 'B1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: 'logo' },
          { kind: 'string', value: '' },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-saved-drawing-artifact-source.xlsx')
        expect(drawingArtifactPaths(excelTruth.snapshot)).toEqual(
          expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png']),
        )

        const workpaper = WorkPaper.buildFromSnapshot(excelTruth.snapshot)
        try {
          const sheet = workpaper.getSheetId('Sheet1')
          if (sheet === undefined) {
            throw new Error('Expected Sheet1 to be available')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 1 }, 'headless edit')

          const headlessPath = join(tempDir, 'headless-drawing-artifact-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Sheet1',
            formulaCells: [],
            inspectCells: ['A1', 'B1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells.map((cell) => cell.value)).toEqual([
            { kind: 'string', value: 'logo' },
            { kind: 'string', value: 'headless edit' },
          ])

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'excel-saved-headless-drawing-artifact.xlsx')
          expect(drawingArtifactPaths(headlessTruth.snapshot)).toEqual(
            expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png']),
          )
          expect(headlessTruth.snapshot.sheets[0]?.metadata?.drawingArtifacts?.relationshipTarget).toBe('../drawings/drawing1.xml')
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel raw DrawingML anchors after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-drawing-artifacts-structure-oracle-')
      try {
        const sourceBytes = exportXlsx(workbookWithEmbeddedImageDrawing())
        const excelWorkbookPath = join(tempDir, 'excel-drawing-artifact-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)

        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Sheet1',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-drawing-artifact-structure-truth.xlsx')
        const excelRows = drawingAnchorRows(excelTruth.snapshot)
        expect(excelRows).toEqual([1, 5])

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(sourceBytes, 'headless-drawing-artifact-structure-source.xlsx').snapshot)
        try {
          const sheet = workpaper.getSheetId('Sheet1')
          if (sheet === undefined) {
            throw new Error('Expected Sheet1 to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(drawingAnchorRows(headlessSnapshot)).toEqual(excelRows)

          const headlessPath = join(tempDir, 'headless-drawing-artifact-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Sheet1',
            formulaCells: [],
            inspectCells: ['A1', 'A2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-drawing-artifact-structure-truth.xlsx')
          expect(drawingAnchorRows(headlessTruth.snapshot)).toEqual(excelRows)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function drawingArtifactPaths(snapshot: WorkbookSnapshot): string[] {
  return snapshot.workbook.metadata?.drawingArtifacts?.parts.map((part) => part.path).toSorted() ?? []
}

function drawingAnchorRows(snapshot: WorkbookSnapshot): number[] {
  return [...drawingPartXml(snapshot).matchAll(/<xdr:row>(\d+)<\/xdr:row>/gu)].map((match) => Number(match[1]))
}

function drawingPartXml(snapshot: WorkbookSnapshot): string {
  const part = snapshot.workbook.metadata?.drawingArtifacts?.parts.find((candidate) => candidate.path === 'xl/drawings/drawing1.xml')
  return part ? Buffer.from(part.dataBase64, 'base64').toString('utf8') : ''
}

function workbookWithEmbeddedImageDrawing(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel drawing artifact oracle',
      metadata: {
        drawingArtifacts: {
          parts: [
            encodedPart('xl/drawings/drawing1.xml', drawingXml),
            encodedPart('xl/drawings/_rels/drawing1.xml.rels', drawingRelationshipsXml),
            encodedPart('xl/media/image1.png', tinyPngBase64, 'base64'),
          ],
          contentTypeDefaults: [{ extension: 'png', contentType: imageContentType }],
          contentTypeOverrides: [{ partName: '/xl/drawings/drawing1.xml', contentType: drawingContentType }],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        metadata: {
          drawingArtifacts: { relationshipTarget: '../drawings/drawing1.xml' },
        },
        cells: [{ address: 'A1', value: 'logo' }],
      },
    ],
  }
}

function encodedPart(path: string, data: string, encoding: 'utf8' | 'base64' = 'utf8'): WorkbookPreservedPackagePartSnapshot {
  const bytes = Buffer.from(data, encoding)
  return {
    path,
    storage: 'base64',
    dataBase64: bytes.toString('base64'),
    byteLength: bytes.byteLength,
  }
}

const tinyPngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='

const drawingRelationshipsXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
  '</Relationships>'

const drawingXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<xdr:twoCellAnchor editAs="oneCell">' +
  '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:pic>' +
  '<xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
  '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
  '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
  '</xdr:pic>' +
  '<xdr:clientData/>' +
  '</xdr:twoCellAnchor>' +
  '</xdr:wsDr>'
