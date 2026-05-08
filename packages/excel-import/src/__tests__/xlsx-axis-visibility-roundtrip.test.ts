import { describe, expect, it } from 'vitest'
import { strFromU8, unzipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX hidden row and column roundtrip', () => {
  it('preserves hidden row and column state in metadata and exported worksheet XML', () => {
    const imported = importXlsx(buildHiddenAxisWorkbookBytes(), 'hidden-axis.xlsx')
    const metadata = imported.snapshot.sheets[0]?.metadata

    expect(metadata?.rows).toContainEqual(
      expect.objectContaining({
        id: 'row:2',
        index: 2,
        hidden: true,
      }),
    )
    expect(metadata?.columns).toContainEqual(
      expect.objectContaining({
        id: 'col:2',
        index: 2,
        hidden: true,
      }),
    )
    expect(metadata?.rowMetadata).toContainEqual(
      expect.objectContaining({
        start: 2,
        count: 1,
        hidden: true,
      }),
    )
    expect(metadata?.columnMetadata).toContainEqual(
      expect.objectContaining({
        start: 2,
        count: 1,
        hidden: true,
      }),
    )

    const exportedSheetXml = strFromU8(unzipSync(exportXlsx(imported.snapshot))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    expect(exportedSheetXml).toMatch(/<col\b(?=[^>]*\bmin="3")(?=[^>]*\bmax="3")(?=[^>]*\bhidden="1")[^>]*\/>/u)
    expect(exportedSheetXml).toMatch(/<row\b(?=[^>]*\br="3")(?=[^>]*\bhidden="1")[^>]*(?:\/|>)/u)
  })
})

function buildHiddenAxisWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['visible a', 'visible b', 'hidden c', 'visible d'],
    [1, 2, 3, 4],
    ['hidden row', 'hidden row', 'hidden row', 'hidden row'],
  ])
  sheet['!rows'] = [undefined, undefined, { hidden: true, hpx: 18 }]
  sheet['!cols'] = [undefined, undefined, { hidden: true, wpx: 96 }, undefined]
  XLSX.utils.book_append_sheet(workbook, sheet, 'Visibility')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
