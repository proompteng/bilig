import { describe, expect, test } from 'vitest'
import {
  WORKBOOK_DEFAULT_FONT_SIZE,
  WORKBOOK_FONT_POINT_TO_CSS_PX,
  WORKBOOK_FONT_SANS,
  workbookFontPointSizeToCssPx,
} from '../workbookTheme.js'

describe('workbookTheme', () => {
  test('uses a spreadsheet-native font stack for every workbook text renderer', () => {
    expect(WORKBOOK_FONT_SANS.startsWith('Arial')).toBe(true)
    expect(WORKBOOK_FONT_SANS).toContain('"Helvetica Neue"')
    expect(WORKBOOK_FONT_SANS).toContain('Helvetica')
    expect(WORKBOOK_FONT_SANS).toContain('"Segoe UI"')
    expect(WORKBOOK_FONT_SANS).not.toContain('"Bilig Sans"')
    expect(WORKBOOK_FONT_SANS).not.toContain('"IBM Plex Sans"')
    expect(WORKBOOK_FONT_SANS).not.toContain('Inter')
    expect(WORKBOOK_FONT_SANS).not.toContain('ui-sans-serif')
    expect(WORKBOOK_FONT_SANS).not.toContain('system-ui')
    expect(WORKBOOK_FONT_SANS).not.toContain('-apple-system')
    expect(WORKBOOK_FONT_SANS).not.toContain('Aptos')
    expect(WORKBOOK_FONT_SANS).not.toContain('Calibri')
  })

  test('keeps the rendered cell default aligned with spreadsheet toolbar sizing', () => {
    expect(WORKBOOK_DEFAULT_FONT_SIZE).toBe(10)
  })

  test('renders spreadsheet point sizes as CSS pixels instead of tiny raw pixels', () => {
    expect(WORKBOOK_FONT_POINT_TO_CSS_PX).toBe(4 / 3)
    expect(workbookFontPointSizeToCssPx(10)).toBe(13.333)
    expect(workbookFontPointSizeToCssPx(11)).toBe(14.667)
    expect(workbookFontPointSizeToCssPx(12)).toBe(16)
    expect(workbookFontPointSizeToCssPx(15)).toBe(20)
  })
})
