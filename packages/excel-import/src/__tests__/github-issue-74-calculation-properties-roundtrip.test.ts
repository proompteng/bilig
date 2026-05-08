import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx } from '../index.js'

describe('GitHub issue #74 XLSX calculation properties roundtrip', () => {
  it('preserves semantic workbook calcPr attributes through import, export, and engine snapshots', () => {
    const imported = importXlsx(buildCalculationPropertiesWorkbookBytes(), 'calculation-properties.xlsx')

    expect(imported.snapshot.workbook.metadata?.calculationSettings).toEqual({
      mode: 'manual',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 10200,
      iterateDelta: '9.9999999999999995E-7',
      fullCalcOnLoad: true,
      concurrentCalc: false,
    })

    const exportedWorkbookXml = workbookXml(exportXlsx(imported.snapshot))
    expect(exportedWorkbookXml).toContain(
      '<calcPr calcMode="manual" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" concurrentCalc="0"/>',
    )

    const engine = new SpreadsheetEngine({ workbookName: 'calculation-properties-engine' })
    engine.importSnapshot(imported.snapshot)
    const exportedFromEngineWorkbookXml = workbookXml(exportXlsx(engine.exportSnapshot()))
    expect(exportedFromEngineWorkbookXml).toContain(
      '<calcPr calcMode="manual" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" concurrentCalc="0"/>',
    )
  })
})

function buildCalculationPropertiesWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['rate'], [0.08]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const workbookXmlWithoutCalcPr = sourceWorkbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, '')
  zip['xl/workbook.xml'] = strToU8(
    workbookXmlWithoutCalcPr.replace(
      '</workbook>',
      '<calcPr calcMode="manual" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" concurrentCalc="0"/></workbook>',
    ),
  )
  return zipSync(zip)
}

function workbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}
