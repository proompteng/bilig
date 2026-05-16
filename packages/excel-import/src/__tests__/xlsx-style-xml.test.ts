import { describe, expect, it } from 'vitest'
import {
  appendCustomCellXfsToStylesXml,
  readCellXfs,
  readXmlAttribute,
  readXmlNonNegativeIntegerAttribute,
  readXmlNumberAttribute,
  readXmlOptionalBooleanAttribute,
  readXmlPositiveIntegerAttribute,
} from '../xlsx-style-xml.js'

describe('xlsx style XML attribute helpers', () => {
  it('reads single and double quoted attributes without prefix collisions', () => {
    const tag = '<row r="3" customHeight=\'1\' hidden="false" outlineLevel="2" customHeightExtra="ignored"/>'

    expect(readXmlAttribute(tag, 'r')).toBe('3')
    expect(readXmlAttribute(tag, 'customHeight')).toBe('1')
    expect(readXmlAttribute(tag, 'height')).toBeNull()
  })

  it('accepts only finite safe integer values for integer helpers', () => {
    expect(readXmlNumberAttribute('<col width="12.5"/>', 'width')).toBe(12.5)
    expect(readXmlNumberAttribute('<col width="Infinity"/>', 'width')).toBeNull()
    expect(readXmlPositiveIntegerAttribute('<col min="1" max="0" style="-1"/>', 'min')).toBe(1)
    expect(readXmlPositiveIntegerAttribute('<col min="1" max="0" style="-1"/>', 'max')).toBeNull()
    expect(readXmlNonNegativeIntegerAttribute('<col min="1" max="0" style="-1"/>', 'max')).toBe(0)
    expect(readXmlNonNegativeIntegerAttribute('<col min="1" max="0" style="-1"/>', 'style')).toBeNull()
  })

  it('preserves tri-state optional booleans for metadata import', () => {
    expect(readXmlOptionalBooleanAttribute('<row hidden="1" customFormat="true" collapsed="0" thickTop="FALSE"/>', 'hidden')).toBe(true)
    expect(readXmlOptionalBooleanAttribute('<row hidden="1" customFormat="true" collapsed="0" thickTop="FALSE"/>', 'customFormat')).toBe(
      true,
    )
    expect(readXmlOptionalBooleanAttribute('<row hidden="1" customFormat="true" collapsed="0" thickTop="FALSE"/>', 'collapsed')).toBe(false)
    expect(readXmlOptionalBooleanAttribute('<row hidden="1" customFormat="true" collapsed="0" thickTop="FALSE"/>', 'thickTop')).toBe(false)
    expect(readXmlOptionalBooleanAttribute('<row hidden="1"/>', 'missing')).toBeNull()
  })

  it('reads self-closing and child-bearing cellXfs entries with XML namespaces', () => {
    const stylesXml = [
      '<x:styleSheet xmlns:x="urn:test">',
      '<x:cellXfs count="2">',
      '<x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '<x:xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0"><x:alignment horizontal="center"/></x:xf>',
      '</x:cellXfs>',
      '</x:styleSheet>',
    ].join('')

    expect(readCellXfs(stylesXml)).toEqual([
      '<x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '<x:xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0"><x:alignment horizontal="center"/></x:xf>',
    ])
  })

  it('appends custom cellXfs and repairs stale count metadata', () => {
    const stylesXml = '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

    expect(appendCustomCellXfsToStylesXml(stylesXml, ['<xf numFmtId="164" applyNumberFormat="1"/>'])).toBe(
      '<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>',
    )
  })
})
