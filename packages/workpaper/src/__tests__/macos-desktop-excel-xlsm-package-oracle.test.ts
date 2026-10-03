import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx, macroExecutionDeclinedWarning } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelPackageOpenSaveOracle } from '@bilig/excel-fixtures'
import type { WorkbookMacroPayloadSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const fixtureUrl = new URL('../../fixtures/excel-oracle/macos-excel-vba-project-source.xlsm', import.meta.url)
const vbaProjectRelationshipType = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const vbaProjectContentType = 'application/vnd.ms-office.vbaProject'
const macroEnabledWorkbookContentType = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
const expectedVbaProjectSha256 = '0ced1464b3677e98f5e3a8c5d80135e18dc98dca39299f1a8cfd2a00999fbf9f'
const expectedVbaProjectByteLength = 15_872

describe('macOS Desktop Excel XLSM/VBA package oracle', () => {
  it('preserves a real VBA project and VML macro button assignment through headless export', () => {
    const sourceBytes = readFixtureBytes()
    const sourceSummary = xlsmPackageSummary(sourceBytes)
    expect(sourceSummary).toEqual(expectedXlsmPackageSummary())

    const imported = importXlsx(sourceBytes, 'macos-excel-vba-project-source.xlsm')
    expect(imported.warnings).toContain(macroExecutionDeclinedWarning)
    expect(macroPayloadSummary(imported.snapshot)).toEqual({
      byteLength: expectedVbaProjectByteLength,
      sha256: expectedVbaProjectSha256,
      workbookCodeName: 'ThisWorkbook',
      sheetCodeNames: [{ sheetName: 'MacroAudit', codeName: 'Sheet1' }],
    })
    expect(imported.snapshot.sheets[0]?.metadata?.controlArtifacts?.relationships).toEqual([
      expect.objectContaining({ target: '../drawings/vmlDrawing1.vml', type: vmlDrawingRelationshipType }),
    ])

    const exportedBytes = exportAfterHeadlessEdit(imported.snapshot)
    expect(xlsmPackageSummary(exportedBytes)).toEqual(sourceSummary)

    const reimported = importXlsx(exportedBytes, 'headless-vba-project.xlsm')
    expect(reimported.warnings).toContain(macroExecutionDeclinedWarning)
    expect(cellValue(reimported.snapshot, 'MacroAudit', 'A1')).toBe('Headless reviewed')
    expect(cellValue(reimported.snapshot, 'MacroAudit', 'B1')).toBe('macro sentinel unchanged')
    expect(macroPayloadSummary(reimported.snapshot)).toEqual(macroPayloadSummary(imported.snapshot))
  })

  it('keeps VBA sheet code names aligned after a headless sheet rename', () => {
    const sourceBytes = readFixtureBytes()
    const imported = importXlsx(sourceBytes, 'macos-excel-vba-project-source.xlsm')

    const renamedBytes = exportAfterHeadlessRename(imported.snapshot, 'MacroReview')
    const reimported = importXlsx(renamedBytes, 'headless-renamed-vba-project.xlsm')

    expect(macroPayloadSummary(reimported.snapshot)).toEqual({
      byteLength: expectedVbaProjectByteLength,
      sha256: expectedVbaProjectSha256,
      workbookCodeName: 'ThisWorkbook',
      sheetCodeNames: [{ sheetName: 'MacroReview', codeName: 'Sheet1' }],
    })
    expect(cellValue(reimported.snapshot, 'MacroReview', 'B1')).toBe('macro sentinel unchanged')
    expect(xlsmPackageSummary(renamedBytes)).toEqual(expectedXlsmPackageSummary())
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves real VBA project and button macro assignment after Desktop Excel open/save and a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-xlsm-package-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-vba-project-source.xlsm')
        writeFileSync(sourcePath, readFixtureBytes())

        const excelSource = runMacosExcelPackageOpenSaveOracle({
          workbookPath: sourcePath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelSource.excelVersion).toMatch(/^\d+\./u)
        const excelSourceBytes = new Uint8Array(readFileSync(sourcePath))
        const excelSourceSummary = xlsmPackageSummary(excelSourceBytes)
        expect(excelSourceSummary).toEqual(expectedXlsmPackageSummary())

        const imported = importXlsx(excelSourceBytes, 'excel-saved-vba-project-source.xlsm')
        expect(imported.warnings).toContain(macroExecutionDeclinedWarning)

        const headlessPath = join(tempDir, 'headless-vba-project.xlsm')
        writeFileSync(headlessPath, exportAfterHeadlessEdit(imported.snapshot))
        expect(xlsmPackageSummary(new Uint8Array(readFileSync(headlessPath)))).toEqual(excelSourceSummary)

        const headlessExcel = runMacosExcelPackageOpenSaveOracle({
          workbookPath: headlessPath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(headlessExcel.excelVersion).toMatch(/^\d+\./u)

        const excelSavedHeadlessBytes = new Uint8Array(readFileSync(headlessPath))
        expect(xlsmPackageSummary(excelSavedHeadlessBytes)).toEqual(excelSourceSummary)

        const reimported = importXlsx(excelSavedHeadlessBytes, 'excel-saved-headless-vba-project.xlsm')
        expect(reimported.warnings).toContain(macroExecutionDeclinedWarning)
        expect(cellValue(reimported.snapshot, 'MacroAudit', 'A1')).toBe('Headless reviewed')
        expect(cellValue(reimported.snapshot, 'MacroAudit', 'B1')).toBe('macro sentinel unchanged')
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'survives Desktop Excel open/save after a headless sheet rename in a macro workbook',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-xlsm-rename-oracle-')
      try {
        const renamedPath = join(tempDir, 'headless-renamed-vba-project.xlsm')
        writeFileSync(renamedPath, exportAfterHeadlessRename(importXlsx(readFixtureBytes(), 'source.xlsm').snapshot, 'MacroReview'))

        const desktopExcel = runMacosExcelPackageOpenSaveOracle({
          workbookPath: renamedPath,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(desktopExcel.excelVersion).toMatch(/^\d+\./u)

        const excelSavedBytes = new Uint8Array(readFileSync(renamedPath))
        expect(xlsmPackageSummary(excelSavedBytes)).toEqual(expectedXlsmPackageSummary())

        const reimported = importXlsx(excelSavedBytes, 'excel-saved-headless-renamed-vba-project.xlsm')
        expect(reimported.warnings).toContain(macroExecutionDeclinedWarning)
        expect(reimported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['MacroReview'])
        expect(cellValue(reimported.snapshot, 'MacroReview', 'B1')).toBe('macro sentinel unchanged')
        expect(macroPayloadSummary(reimported.snapshot)?.sheetCodeNames).toEqual([{ sheetName: 'MacroReview', codeName: 'Sheet1' }])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})

function readFixtureBytes(): Uint8Array {
  return new Uint8Array(readFileSync(fixtureUrl))
}

function exportAfterHeadlessEdit(snapshot: WorkbookSnapshot): Uint8Array {
  const workpaper = WorkPaper.buildFromSnapshot(snapshot)
  try {
    const sheet = workpaper.getSheetId('MacroAudit')
    if (sheet === undefined) {
      throw new Error('Expected MacroAudit sheet to be available')
    }
    workpaper.setCellContents({ sheet, row: 0, col: 0 }, 'Headless reviewed')
    return exportXlsx(workpaper.exportSnapshot())
  } finally {
    workpaper.dispose()
  }
}

function exportAfterHeadlessRename(snapshot: WorkbookSnapshot, nextName: string): Uint8Array {
  const workpaper = WorkPaper.buildFromSnapshot(snapshot)
  try {
    const sheet = workpaper.getSheetId('MacroAudit')
    if (sheet === undefined) {
      throw new Error('Expected MacroAudit sheet to be available')
    }
    workpaper.renameSheet(sheet, nextName)
    const renamedSnapshot = workpaper.exportSnapshot()
    expect(macroPayloadSummary(renamedSnapshot)?.sheetCodeNames).toEqual([{ sheetName: nextName, codeName: 'Sheet1' }])
    return exportXlsx(renamedSnapshot)
  } finally {
    workpaper.dispose()
  }
}

function expectedXlsmPackageSummary(): XlsmPackageSummary {
  return {
    vbaProject: {
      byteLength: expectedVbaProjectByteLength,
      sha256: expectedVbaProjectSha256,
    },
    contentTypes: {
      binDefault: vbaProjectContentType,
      workbookOverride: macroEnabledWorkbookContentType,
    },
    workbookCodeName: 'ThisWorkbook',
    sheetCodeNames: ['Sheet1'],
    workbookVbaRelationship: {
      target: 'vbaProject.bin',
      type: vbaProjectRelationshipType,
    },
    vmlDrawingParts: ['xl/drawings/vmlDrawing1.vml'],
    vmlMacroAssignments: ['[0]!say_hello'],
    sheetVmlRelationships: [{ target: '../drawings/vmlDrawing1.vml', type: vmlDrawingRelationshipType }],
  }
}

interface XlsmPackageSummary {
  readonly vbaProject: {
    readonly byteLength: number
    readonly sha256: string
  }
  readonly contentTypes: {
    readonly binDefault: string | undefined
    readonly workbookOverride: string | undefined
  }
  readonly workbookCodeName: string | undefined
  readonly sheetCodeNames: readonly string[]
  readonly workbookVbaRelationship:
    | {
        readonly target: string
        readonly type: string
      }
    | undefined
  readonly vmlDrawingParts: readonly string[]
  readonly vmlMacroAssignments: readonly string[]
  readonly sheetVmlRelationships: readonly {
    readonly target: string
    readonly type: string
  }[]
}

function xlsmPackageSummary(bytes: Uint8Array): XlsmPackageSummary {
  const zip = unzipSync(bytes)
  const vbaProject = zip['xl/vbaProject.bin']
  if (!vbaProject) {
    throw new Error('Expected XLSM package to contain xl/vbaProject.bin')
  }
  const workbookXml = readZipText(zip, 'xl/workbook.xml')
  return {
    vbaProject: {
      byteLength: vbaProject.byteLength,
      sha256: sha256(vbaProject),
    },
    contentTypes: {
      binDefault: readContentTypeDefault(zip, 'bin'),
      workbookOverride: readContentTypeOverride(zip, '/xl/workbook.xml'),
    },
    workbookCodeName: readXmlAttribute(/<workbookPr\b([^>]*)>/u.exec(workbookXml)?.[1] ?? '', 'codeName') ?? undefined,
    sheetCodeNames: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/sheet\d+\.xml$/u.test(path))
      .flatMap(([, data]) => {
        const sheetXml = strFromU8(data)
        return [...sheetXml.matchAll(/<sheetPr\b([^>]*)>/gu)].flatMap((match) => {
          const codeName = readXmlAttribute(match[1] ?? '', 'codeName')
          return codeName ? [codeName] : []
        })
      }),
    workbookVbaRelationship: readWorkbookRelationship(zip, vbaProjectRelationshipType),
    vmlDrawingParts: Object.keys(zip)
      .filter((path) => /^xl\/drawings\/vmlDrawing\d+\.vml$/u.test(path))
      .toSorted(),
    vmlMacroAssignments: Object.entries(zip)
      .filter(([path]) => /^xl\/drawings\/vmlDrawing\d+\.vml$/u.test(path))
      .flatMap(([, data]) => [...strFromU8(data).matchAll(/<x:FmlaMacro>([\s\S]*?)<\/x:FmlaMacro>/gu)].map((match) => match[1] ?? ''))
      .toSorted(),
    sheetVmlRelationships: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/u.test(path))
      .flatMap(([, data]) => readRelationships(strFromU8(data), vmlDrawingRelationshipType))
      .toSorted((left, right) => left.target.localeCompare(right.target)),
  }
}

function macroPayloadSummary(snapshot: WorkbookSnapshot):
  | {
      readonly byteLength: number
      readonly sha256: string
      readonly workbookCodeName: string | undefined
      readonly sheetCodeNames: WorkbookMacroPayloadSnapshot['sheetCodeNames']
    }
  | undefined {
  const payload = snapshot.workbook.metadata?.macroPayloads?.[0]
  if (!payload) {
    return undefined
  }
  return {
    byteLength: payload.byteLength,
    sha256: sha256(Buffer.from(payload.dataBase64, 'base64')),
    workbookCodeName: payload.workbookCodeName,
    sheetCodeNames: payload.sheetCodeNames,
  }
}

function cellValue(snapshot: WorkbookSnapshot, sheetName: string, address: string): unknown {
  return snapshot.sheets.find((sheet) => sheet.name === sheetName)?.cells.find((cell) => cell.address === address)?.value
}

function readZipText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSM package part: ${path}`)
  }
  return strFromU8(bytes)
}

function readWorkbookRelationship(
  zip: Record<string, Uint8Array>,
  relationshipType: string,
): { readonly target: string; readonly type: string } | undefined {
  return readRelationships(readZipText(zip, 'xl/_rels/workbook.xml.rels'), relationshipType)[0]
}

function readRelationships(xml: string, relationshipType: string): { readonly target: string; readonly type: string }[] {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const type = readXmlAttribute(attributes, 'Type')
    const target = readXmlAttribute(attributes, 'Target')
    return type === relationshipType && target ? [{ target, type }] : []
  })
}

function readContentTypeDefault(zip: Record<string, Uint8Array>, extension: string): string | undefined {
  const contentTypesXml = readZipText(zip, '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Default\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Extension') === extension) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function readContentTypeOverride(zip: Record<string, Uint8Array>, partName: string): string | undefined {
  const contentTypesXml = readZipText(zip, '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}
