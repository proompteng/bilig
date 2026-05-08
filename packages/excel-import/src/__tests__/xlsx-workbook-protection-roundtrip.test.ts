import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'

import { exportXlsx, importXlsx } from '../index.js'

describe('workbook structure protection import/export', () => {
  it('preserves workbookProtection attributes across XLSX round trips', () => {
    const imported = importXlsx(buildWorkbookProtectionBytes(), 'protected-structure.xlsx')
    const workbookProtection = imported.snapshot.workbook.metadata?.workbookProtection

    expect(imported.warnings).toEqual([])
    expect(workbookProtection).toEqual({
      lockStructure: true,
      lockWindows: false,
      xmlAttributes: [
        { name: 'lockStructure', value: '1' },
        { name: 'lockWindows', value: '0' },
        { name: 'workbookPassword', value: 'AF2B' },
        { name: 'revisionsPassword', value: 'BC3D' },
      ],
    })

    const exportedXml = workbookXml(exportXlsx(imported.snapshot))
    expect(exportedXml).toContain('<workbookProtection ')
    expect(exportedXml).toContain('lockStructure="1"')
    expect(exportedXml).toContain('lockWindows="0"')
    expect(exportedXml).toContain('workbookPassword="AF2B"')
    expect(exportedXml).toContain('revisionsPassword="BC3D"')
  })

  it('keeps workbookProtection after engine import/export', async () => {
    const imported = importXlsx(buildWorkbookProtectionBytes(), 'protected-structure.xlsx')
    const engine = new SpreadsheetEngine({ workbookName: 'protected-structure' })
    await engine.ready()

    engine.importSnapshot(imported.snapshot)

    expect(engine.exportSnapshot().workbook.metadata?.workbookProtection).toEqual(imported.snapshot.workbook.metadata?.workbookProtection)
    expect(workbookXml(exportXlsx(engine.exportSnapshot()))).toContain('lockStructure="1"')
  })
})

function buildWorkbookProtectionBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Control', 'Value'],
    ['Revenue', 100],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const workbookProtectionXml = '<workbookProtection lockStructure="1" lockWindows="0" workbookPassword="AF2B" revisionsPassword="BC3D"/>'
  zip['xl/workbook.xml'] = strToU8(insertWorkbookProtection(sourceWorkbookXml, workbookProtectionXml))
  return zipSync(zip)
}

function insertWorkbookProtection(sourceWorkbookXml: string, workbookProtectionXml: string): string {
  const workbookPrMatch = /<workbookPr\b[^>]*(?:\/>|>[\s\S]*?<\/workbookPr>)/u.exec(sourceWorkbookXml)
  if (workbookPrMatch?.index !== undefined) {
    const insertIndex = workbookPrMatch.index + workbookPrMatch[0].length
    return `${sourceWorkbookXml.slice(0, insertIndex)}${workbookProtectionXml}${sourceWorkbookXml.slice(insertIndex)}`
  }
  return sourceWorkbookXml.replace(/<sheets\b/u, `${workbookProtectionXml}<sheets`)
}

function workbookXml(snapshotOrBytes: WorkbookSnapshot | Uint8Array): string {
  const bytes = snapshotOrBytes instanceof Uint8Array ? snapshotOrBytes : exportXlsx(snapshotOrBytes)
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}
