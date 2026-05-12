import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const threadedCommentRelationshipType = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'
const personRelationshipType = 'http://schemas.microsoft.com/office/2017/10/relationships/person'
const threadedCommentContentType = 'application/vnd.ms-excel.threadedcomments+xml'
const personContentType = 'application/vnd.ms-excel.person+xml'

describe('threaded comment package roundtrip', () => {
  it('preserves threaded comment parts, person parts, and package relationships across XLSX round trips', () => {
    const source = buildThreadedCommentWorkbookBytes()
    const sourceSummary = readThreadedCommentSummary(source)

    const imported = importXlsx(source, 'threaded-comments.xlsx')
    const exported = exportXlsx(imported.snapshot)
    const exportedSummary = readThreadedCommentSummary(exported)

    expect(exportedSummary).toEqual(sourceSummary)
    expect(exportedSummary.threadedCommentRecordCount).toBe(3)
    expect(exportedSummary.threadedCommentPartCount).toBe(2)
    expect(exportedSummary.personPartCount).toBe(1)
    expect(exportedSummary.sheetThreadedCommentRelationshipCount).toBe(2)
    expect(exportedSummary.workbookPersonRelationshipCount).toBe(1)
  })
})

interface ThreadedCommentSummary {
  readonly threadedCommentRecordCount: number
  readonly threadedCommentPartCount: number
  readonly personPartCount: number
  readonly sheetThreadedCommentRelationshipCount: number
  readonly workbookPersonRelationshipCount: number
  readonly contentTypeOverrides: readonly string[]
  readonly threadedCommentXmlByPath: readonly [string, string][]
  readonly personXmlByPath: readonly [string, string][]
}

function buildThreadedCommentWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Assumption', 42]]), 'Inputs')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Review', 84]]), 'Review')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))

  zip['xl/threadedComments/threadedComment1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<threadedComment ref="A1" dT="2026-05-01T10:00:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}"><text><x:r><x:t>Check revenue assumption</x:t></x:r></text></threadedComment>',
      '<threadedComment ref="B1" dT="2026-05-01T10:05:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb}"><text><x:r><x:t>Review model output</x:t></x:r></text></threadedComment>',
      '</ThreadedComments>',
    ].join(''),
  )
  zip['xl/threadedComments/threadedComment2.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<threadedComment ref="A1" dT="2026-05-01T11:00:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{cccccccc-cccc-cccc-cccc-cccccccccccc}"><text><x:r><x:t>Resolved by controller</x:t></x:r></text></threadedComment>',
      '</ThreadedComments>',
    ].join(''),
  )
  zip['xl/persons/person1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/person">',
      '<person displayName="Finance Reviewer" id="{11111111-1111-1111-1111-111111111111}" userId="finance@example.com" providerId="None"/>',
      '</personList>',
    ].join(''),
  )

  upsertRelationship(zip, 'xl/worksheets/_rels/sheet1.xml.rels', {
    id: 'rIdThreaded1',
    type: threadedCommentRelationshipType,
    target: '../threadedComments/threadedComment1.xml',
  })
  upsertRelationship(zip, 'xl/worksheets/_rels/sheet2.xml.rels', {
    id: 'rIdThreaded2',
    type: threadedCommentRelationshipType,
    target: '../threadedComments/threadedComment2.xml',
  })
  upsertRelationship(zip, 'xl/_rels/workbook.xml.rels', {
    id: 'rIdPerson1',
    type: personRelationshipType,
    target: 'persons/person1.xml',
  })
  addContentTypeOverride(zip, '/xl/threadedComments/threadedComment1.xml', threadedCommentContentType)
  addContentTypeOverride(zip, '/xl/threadedComments/threadedComment2.xml', threadedCommentContentType)
  addContentTypeOverride(zip, '/xl/persons/person1.xml', personContentType)
  return zipSync(zip)
}

function upsertRelationship(
  zip: Record<string, Uint8Array>,
  relsPath: string,
  relationship: {
    readonly id: string
    readonly type: string
    readonly target: string
  },
): void {
  const relationshipXml = `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`
  const currentXml = strFromU8(zip[relsPath] ?? new Uint8Array())
  zip[relsPath] = strToU8(
    currentXml.includes('</Relationships>')
      ? currentXml.replace('</Relationships>', `${relationshipXml}</Relationships>`)
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">${relationshipXml}</Relationships>`,
  )
}

function addContentTypeOverride(zip: Record<string, Uint8Array>, partName: string, contentType: string): void {
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.includes(`PartName="${partName}"`)
      ? contentTypesXml
      : contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`),
  )
}

function readThreadedCommentSummary(bytes: Uint8Array): ThreadedCommentSummary {
  const zip = unzipSync(bytes)
  const threadedCommentXmlByPath = readXmlParts(zip, /^xl\/threadedComments\/threadedComment\d+\.xml$/u)
  const personXmlByPath = readXmlParts(zip, /^xl\/persons\/person\d+\.xml$/u)
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  return {
    threadedCommentRecordCount: threadedCommentXmlByPath.reduce(
      (count, [_path, xml]) => count + (xml.match(/<threadedComment\b/gu)?.length ?? 0),
      0,
    ),
    threadedCommentPartCount: threadedCommentXmlByPath.length,
    personPartCount: personXmlByPath.length,
    sheetThreadedCommentRelationshipCount: Object.entries(zip)
      .filter(([path]) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/u.test(path))
      .reduce((count, [_path, rels]) => count + countOccurrences(strFromU8(rels), threadedCommentRelationshipType), 0),
    workbookPersonRelationshipCount: countOccurrences(
      strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()),
      personRelationshipType,
    ),
    contentTypeOverrides: [...contentTypesXml.matchAll(/<Override\b[^>]*PartName="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*\/>/gu)]
      .map((match) => `${match[1] ?? ''}:${match[2] ?? ''}`)
      .filter((entry) => entry.includes('threadedComments') || entry.includes('persons'))
      .toSorted(),
    threadedCommentXmlByPath,
    personXmlByPath,
  }
}

function readXmlParts(zip: Record<string, Uint8Array>, pattern: RegExp): readonly [string, string][] {
  return Object.entries(zip)
    .filter(([path]) => pattern.test(path))
    .map(([path, bytes]): [string, string] => [path, strFromU8(bytes)])
    .toSorted(([left], [right]) => left.localeCompare(right))
}

function countOccurrences(value: string, needle: string): number {
  return value.split(needle).length - 1
}
