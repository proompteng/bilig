import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('workbook view state roundtrip', () => {
  it('preserves workbook and worksheet view state across XLSX round trips', () => {
    const source = buildWorkbookWithViewState()
    const sourceSummary = readViewStateSummary(source)

    const imported = importXlsx(source, 'view-state.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedSummary = readViewStateSummary(exported)

    expect(exportedSummary).toEqual(sourceSummary)
    expect(exportedSummary.bookViewsXml).toContain('activeTab="1"')
    expect(exportedSummary.bookViewsXml).toContain('firstSheet="1"')
    expect(exportedSummary.sheetViewsXmlByPath).toContainEqual([
      'xl/worksheets/sheet1.xml',
      '<sheetViews><sheetView workbookViewId="0" showGridLines="0" topLeftCell="C10" zoomScale="85" zoomScaleNormal="90"><selection activeCell="D48" sqref="D48:K48"/></sheetView></sheetViews>',
    ])
    expect(exportedSummary.sheetViewsXmlByPath).toContainEqual([
      'xl/worksheets/sheet2.xml',
      '<sheetViews><sheetView workbookViewId="0" tabSelected="1" topLeftCell="B7" zoomScale="125" zoomScalePageLayoutView="80" view="pageLayout"><selection pane="topLeft" activeCell="B7" sqref="B7:C9"/></sheetView></sheetViews>',
    ])
  })
})

interface ViewStateSummary {
  readonly bookViewsXml: string
  readonly sheetViewsXmlByPath: readonly [string, string][]
}

function buildWorkbookWithViewState(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Metric', 'Value'],
      ['Revenue', 1250],
    ]),
    'Summary',
  )
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Input', 'Value'],
      ['Growth', 0.08],
    ]),
    'Inputs',
  )
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))

  upsertXmlSection(
    zip,
    'xl/workbook.xml',
    'bookViews',
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000" activeTab="1" firstSheet="1"/></bookViews>',
  )
  upsertXmlSection(
    zip,
    'xl/worksheets/sheet1.xml',
    'sheetViews',
    '<sheetViews><sheetView workbookViewId="0" showGridLines="0" topLeftCell="C10" zoomScale="85" zoomScaleNormal="90"><selection activeCell="D48" sqref="D48:K48"/></sheetView></sheetViews>',
  )
  upsertXmlSection(
    zip,
    'xl/worksheets/sheet2.xml',
    'sheetViews',
    '<sheetViews><sheetView workbookViewId="0" tabSelected="1" topLeftCell="B7" zoomScale="125" zoomScalePageLayoutView="80" view="pageLayout"><selection pane="topLeft" activeCell="B7" sqref="B7:C9"/></sheetView></sheetViews>',
  )
  return zipSync(zip)
}

function upsertXmlSection(zip: Record<string, Uint8Array>, path: string, localName: 'bookViews' | 'sheetViews', sectionXml: string): void {
  const currentXml = strFromU8(zip[path] ?? new Uint8Array())
  zip[path] = strToU8(replaceXmlSection(currentXml, localName, sectionXml))
}

function replaceXmlSection(xml: string, localName: 'bookViews' | 'sheetViews', sectionXml: string): string {
  const existing = readXmlSection(xml, localName)
  if (existing) {
    return xml.replace(existing, sectionXml)
  }
  const rootName = localName === 'bookViews' ? 'workbook' : 'worksheet'
  return xml.replace(new RegExp(`<${rootName}\\b([^>]*)>`, 'u'), `<${rootName}$1>${sectionXml}`)
}

function readViewStateSummary(bytes: Uint8Array): ViewStateSummary {
  const zip = unzipSync(bytes)
  return {
    bookViewsXml: readXmlSection(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array()), 'bookViews') ?? '',
    sheetViewsXmlByPath: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/sheet\d+\.xml$/u.test(path))
      .map(([path, data]): [string, string] => [path, readXmlSection(strFromU8(data), 'sheetViews') ?? ''])
      .toSorted(([left], [right]) => left.localeCompare(right)),
  }
}

function readXmlSection(xml: string, localName: 'bookViews' | 'sheetViews'): string | undefined {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${localName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)?.[0]
}
