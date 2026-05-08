import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('legacy comment rich text roundtrip', () => {
  it('preserves legacy note rich text runs on no-op roundtrip', () => {
    const source = buildRichTextLegacyCommentWorkbookBytes()
    const sourceCommentsXml = readCommentsXml(source)

    const imported = importXlsx(source, 'legacy-comment-rich-text.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedCommentsXml = readCommentsXml(exported)

    expect(imported.snapshot.sheets[0]?.metadata?.commentThreads?.[0]?.comments[0]).toMatchObject({
      authorDisplayName: 'Finance',
      body: 'Reviewed total needs CFO approval',
    })
    expect(countRichTextRuns(exportedCommentsXml)).toBe(3)
    expect(exportedCommentsXml).toBe(sourceCommentsXml)
  })

  it('does not preserve stale imported rich text comment XML after comment text changes', () => {
    const source = buildRichTextLegacyCommentWorkbookBytes()
    const sourceCommentsXml = readCommentsXml(source)
    const imported = importXlsx(source, 'legacy-comment-rich-text.xlsx')
    const firstComment = imported.snapshot.sheets[0]?.metadata?.commentThreads?.[0]?.comments[0]
    if (!firstComment) {
      throw new Error('Fixture import did not produce the expected comment thread.')
    }
    firstComment.body = 'Edited approval note'

    const exportedCommentsXml = readCommentsXml(exportXlsx(imported.snapshot))

    expect(exportedCommentsXml).not.toBe(sourceCommentsXml)
    expect(exportedCommentsXml).toContain('Edited approval note')
    expect(exportedCommentsXml).not.toContain('<b/>')
  })
})

function buildRichTextLegacyCommentWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const worksheet: XLSX.WorkSheet = {
    A1: {
      t: 'n',
      v: 42,
      c: [{ a: 'Finance', t: 'Reviewed total needs CFO approval' }],
    },
    '!ref': 'A1:A1',
  }
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Notes')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const commentsPath = readCommentsPath(zip)
  zip[commentsPath] = strToU8(
    strFromU8(zip[commentsPath] ?? new Uint8Array()).replace(
      /<text>[\s\S]*?<\/text>/u,
      [
        '<text>',
        '<r><rPr><b/><sz val="9"/><color rgb="FF1F4E79"/><rFont val="Calibri"/></rPr><t>Reviewed total</t></r>',
        '<r><rPr><i/><sz val="9"/><color rgb="FFC00000"/><rFont val="Calibri"/></rPr><t xml:space="preserve"> needs CFO </t></r>',
        '<r><rPr><u/><sz val="9"/><color rgb="FF7030A0"/><rFont val="Calibri"/></rPr><t>approval</t></r>',
        '</text>',
      ].join(''),
    ),
  )
  return zipSync(zip)
}

function readCommentsXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  return strFromU8(zip[readCommentsPath(zip)] ?? new Uint8Array())
}

function readCommentsPath(zip: Record<string, Uint8Array>): string {
  const commentsPath = Object.keys(zip).find((path) => /^xl\/(?:comments\d+|comments\/comment\d+)\.xml$/u.test(path))
  if (!commentsPath) {
    throw new Error('Workbook fixture did not include a comments part.')
  }
  return commentsPath
}

function countRichTextRuns(commentsXml: string): number {
  return [...commentsXml.matchAll(/<r>/gu)].length
}
