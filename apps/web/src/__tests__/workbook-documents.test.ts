// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createBlankWorkbook, duplicateWorkbook, loadRecentWorkbooks, rememberWorkbook } from '../workbook-documents.js'

describe('workbook document actions', () => {
  const storage = new Map<string, string>()
  beforeEach(() => {
    storage.clear()
    vi.stubGlobal('localStorage', {
      getItem: (key: string) => storage.get(key) ?? null,
      setItem: (key: string, value: string) => storage.set(key, value),
    })
  })
  afterEach(() => vi.unstubAllGlobals())
  it('starts a named workbook with a clean sheet', () => {
    expect(createBlankWorkbook()).toEqual({
      version: 1,
      workbook: { name: 'Untitled workbook' },
      sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }],
    })
  })
  it('duplicates all workbook data without changing the original', () => {
    const original = createBlankWorkbook()
    original.sheets[0].cells.push({ address: 'A1', formula: 'SUM(3,4,5)', format: '0.00' })
    const copy = duplicateWorkbook(original)
    expect(copy.workbook.name).toBe('Copy of Untitled workbook')
    expect(copy.sheets).toEqual(original.sheets)
    copy.sheets[0].cells[0].formula = '1+1'
    expect(original.sheets[0].cells[0].formula).toBe('SUM(3,4,5)')
  })
  it('keeps bounded recents per user and updates renamed documents', () => {
    for (let i = 0; i < 15; i++) rememberWorkbook('alice', { documentId: String(i), name: `Book ${i}` })
    rememberWorkbook('alice', { documentId: '14', name: 'Budget' })
    expect(loadRecentWorkbooks('alice')).toHaveLength(12)
    expect(loadRecentWorkbooks('alice')[0]).toEqual({ documentId: '14', name: 'Budget' })
    expect(loadRecentWorkbooks('bob')).toEqual([])
  })
  it('ignores damaged browser storage', () => {
    localStorage.setItem('bilig:recent-workbooks:alice', '{')
    expect(loadRecentWorkbooks('alice')).toEqual([])
  })
})
