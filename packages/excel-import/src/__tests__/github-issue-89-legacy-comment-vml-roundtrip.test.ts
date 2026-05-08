import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

interface LegacyNoteVmlMetadata {
  readonly anchor: string
  readonly fillColor: string
  readonly height: string
  readonly marginLeft: string
  readonly marginTop: string
  readonly visibility: string
  readonly width: string
}

describe('GitHub issue #89 legacy comment VML roundtrip', () => {
  it('preserves legacy note VML anchors, geometry, fill, and visibility on no-op roundtrip', () => {
    const source = buildLegacyCommentVmlWorkbookBytes()
    const sourceMetadata = readLegacyNoteVmlMetadata(source)

    const imported = importXlsx(source, 'legacy-note-vml.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedMetadata = readLegacyNoteVmlMetadata(exported)

    expect(exportedMetadata).toEqual(sourceMetadata)
    expect(exportedMetadata.get('0:0')).toMatchObject({
      anchor: '1, 20, 2, 10, 4, 40, 6, 18',
      fillColor: '#fce4d6',
      visibility: 'hidden',
    })
    expect(exportedMetadata.get('2:2')).toMatchObject({
      anchor: '3, 14, 4, 8, 7, 28, 9, 12',
      fillColor: '#d9ead3',
      visibility: 'visible',
    })
  })

  it('does not preserve stale imported VML after comment text changes', () => {
    const source = buildLegacyCommentVmlWorkbookBytes()
    const sourceMetadata = readLegacyNoteVmlMetadata(source)
    const imported = importXlsx(source, 'legacy-note-vml.xlsx')
    const firstComment = imported.snapshot.sheets[0]?.metadata?.commentThreads?.[0]?.comments[0]
    if (!firstComment) {
      throw new Error('Fixture import did not produce the expected comment thread.')
    }
    firstComment.body = 'Edited assumption note'

    const exportedMetadata = readLegacyNoteVmlMetadata(exportXlsx(imported.snapshot))

    expect(exportedMetadata).not.toEqual(sourceMetadata)
    expect(exportedMetadata.get('0:0')?.anchor).not.toBe('1, 20, 2, 10, 4, 40, 6, 18')
  })
})

function buildLegacyCommentVmlWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const worksheet: XLSX.WorkSheet = {
    A1: { t: 's', v: 'Hidden note', c: [{ a: 'Audit', t: 'Hidden assumption note' }] },
    C3: { t: 'n', v: 42, c: [{ a: 'Review', t: 'Visible review note' }] },
    '!ref': 'A1:C3',
  }
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Notes')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const vmlPath = Object.keys(zip).find((path) => path.startsWith('xl/drawings/') && path.endsWith('.vml'))
  if (!vmlPath) {
    throw new Error('Fixture writer did not create a legacy comment VML drawing.')
  }
  const customizedVml = customizeNoteShape(
    customizeNoteShape(strFromU8(zip[vmlPath] ?? new Uint8Array()), {
      row: 0,
      column: 0,
      anchor: '1, 20, 2, 10, 4, 40, 6, 18',
      fillColor: '#fce4d6',
      marginLeft: '71.25pt',
      marginTop: '16.5pt',
      width: '192pt',
      height: '88.5pt',
      visible: false,
    }),
    {
      row: 2,
      column: 2,
      anchor: '3, 14, 4, 8, 7, 28, 9, 12',
      fillColor: '#d9ead3',
      marginLeft: '149.25pt',
      marginTop: '54pt',
      width: '216pt',
      height: '111pt',
      visible: true,
    },
  )
  zip[vmlPath] = strToU8(customizedVml)
  return zipSync(zip)
}

function customizeNoteShape(
  vmlXml: string,
  input: {
    readonly row: number
    readonly column: number
    readonly anchor: string
    readonly fillColor: string
    readonly marginLeft: string
    readonly marginTop: string
    readonly width: string
    readonly height: string
    readonly visible: boolean
  },
): string {
  return vmlXml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/gu, (shapeXml) => {
    if (!shapeXml.includes('<x:ClientData ObjectType="Note">')) {
      return shapeXml
    }
    if (!shapeXml.includes(`<x:Row>${String(input.row)}</x:Row>`) || !shapeXml.includes(`<x:Column>${String(input.column)}</x:Column>`)) {
      return shapeXml
    }
    const style = [
      'position:absolute',
      `margin-left:${input.marginLeft}`,
      `margin-top:${input.marginTop}`,
      `width:${input.width}`,
      `height:${input.height}`,
      'z-index:1',
      `visibility:${input.visible ? 'visible' : 'hidden'}`,
    ].join(';')
    const withStyle = shapeXml.replace(/\sstyle=(["'])[\s\S]*?\1/u, ` style="${style}"`)
    const withFill = withStyle.replace(/\sfillcolor=(["'])[\s\S]*?\1/u, ` fillcolor="${input.fillColor}"`)
    const withAnchor = withFill.replace(/<x:Anchor>[\s\S]*?<\/x:Anchor>/u, `<x:Anchor>${input.anchor}</x:Anchor>`)
    if (input.visible) {
      return /<x:Visible\s*\/>/u.test(withAnchor)
        ? withAnchor
        : withAnchor.replace('<x:ClientData ObjectType="Note">', '<x:ClientData ObjectType="Note"><x:Visible/>')
    }
    return withAnchor.replace(/<x:Visible\s*\/>/gu, '')
  })
}

function readLegacyNoteVmlMetadata(bytes: Uint8Array): Map<string, LegacyNoteVmlMetadata> {
  const zip = unzipSync(bytes)
  const metadata = new Map<string, LegacyNoteVmlMetadata>()
  for (const [path, data] of Object.entries(zip)) {
    if (!path.startsWith('xl/drawings/') || !path.endsWith('.vml')) {
      continue
    }
    const vmlXml = strFromU8(data)
    for (const shapeXml of vmlXml.match(/<v:shape\b[\s\S]*?<\/v:shape>/gu) ?? []) {
      if (!shapeXml.includes('<x:ClientData ObjectType="Note">')) {
        continue
      }
      const row = readXmlText(shapeXml, 'Row')
      const column = readXmlText(shapeXml, 'Column')
      if (row === null || column === null) {
        continue
      }
      const style = readAttribute(shapeXml, 'style') ?? ''
      metadata.set(`${row}:${column}`, {
        anchor: readXmlText(shapeXml, 'Anchor') ?? '',
        fillColor: readAttribute(shapeXml, 'fillcolor') ?? '',
        height: readStyleDeclaration(style, 'height') ?? '',
        marginLeft: readStyleDeclaration(style, 'margin-left') ?? '',
        marginTop: readStyleDeclaration(style, 'margin-top') ?? '',
        visibility: readStyleDeclaration(style, 'visibility') ?? '',
        width: readStyleDeclaration(style, 'width') ?? '',
      })
    }
  }
  return metadata
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function readXmlText(xml: string, localName: string): string | null {
  const match = new RegExp(`<x:${localName}>([\\s\\S]*?)</x:${localName}>`, 'u').exec(xml)
  return match?.[1]?.trim() ?? null
}

function readStyleDeclaration(style: string, propertyName: string): string | null {
  for (const declaration of style.split(';')) {
    const [rawName, ...rawValue] = declaration.split(':')
    if (rawName?.trim().toLowerCase() === propertyName) {
      return rawValue.join(':').trim()
    }
  }
  return null
}
