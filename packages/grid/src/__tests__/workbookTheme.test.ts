import { describe, expect, test } from 'vitest'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS } from '../workbookTheme.js'

describe('workbookTheme', () => {
  test('uses the product font stack for rendered grid text', () => {
    expect(WORKBOOK_FONT_SANS.startsWith('Arial, "Helvetica Neue", Helvetica')).toBe(true)
    expect(WORKBOOK_FONT_SANS).toContain('"Segoe UI"')
    expect(WORKBOOK_FONT_SANS).toContain('Arial')
    expect(WORKBOOK_FONT_SANS).not.toContain('IBM Plex')
    expect(WORKBOOK_FONT_SANS).not.toContain('Inter')
    expect(WORKBOOK_FONT_SANS).not.toContain('Aptos')
    expect(WORKBOOK_FONT_SANS).not.toContain('Calibri')
  })

  test('keeps the rendered cell default aligned with spreadsheet toolbar sizing', () => {
    expect(WORKBOOK_DEFAULT_FONT_SIZE).toBe(11)
  })
})
