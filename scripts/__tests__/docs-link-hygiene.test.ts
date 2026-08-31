import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)
const comparisonDocPath = resolve(repoRoot, 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md')

describe('maintained documentation links', () => {
  it('keeps the SheetJS comparison guide on maintained repository paths', () => {
    const source = readFileSync(comparisonDocPath, 'utf8')

    expect(source).not.toContain('examples/workpaper-workpaper')
    expect(source).not.toContain('docs/workpaper-spreadsheet-engine-comparison.md')
    expect(source).toContain('examples/headless-workpaper')
    expect(source).toContain('docs/headless-spreadsheet-engine-comparison.md')
  })
})
