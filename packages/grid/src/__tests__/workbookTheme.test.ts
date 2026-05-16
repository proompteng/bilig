import { describe, expect, test } from 'vitest'
import {
  WORKBOOK_DEFAULT_FONT_SIZE,
  WORKBOOK_FONT_POINT_TO_CSS_PX,
  WORKBOOK_FONT_SANS,
  workbookFontPointSizeToCssPx,
} from '../workbookTheme.js'

describe('workbookTheme', () => {
  test('uses the product font stack for rendered grid text', () => {
    expect(WORKBOOK_FONT_SANS.startsWith('"IBM Plex Sans", Inter')).toBe(true)
    expect(WORKBOOK_FONT_SANS).toContain('"SF Pro Text"')
    expect(WORKBOOK_FONT_SANS).toContain('"SF Pro Display"')
    expect(WORKBOOK_FONT_SANS).toContain('"Segoe UI"')
    expect(WORKBOOK_FONT_SANS).toContain('Arial')
    expect(WORKBOOK_FONT_SANS).not.toContain('Aptos')
    expect(WORKBOOK_FONT_SANS).not.toContain('Calibri')
  })

  test('keeps the rendered cell default aligned with spreadsheet toolbar sizing', () => {
    expect(WORKBOOK_DEFAULT_FONT_SIZE).toBe(11)
  })

  test('renders spreadsheet point sizes as CSS pixels instead of tiny raw pixels', () => {
    expect(WORKBOOK_FONT_POINT_TO_CSS_PX).toBe(4 / 3)
    expect(workbookFontPointSizeToCssPx(11)).toBe(14.667)
    expect(workbookFontPointSizeToCssPx(12)).toBe(16)
    expect(workbookFontPointSizeToCssPx(15)).toBe(20)
  })
})
