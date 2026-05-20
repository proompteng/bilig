import { describe, expect, it } from 'vitest'
import { scanWorkspaceResolution, workspaceRootDir } from '../workspace-resolution.js'

describe('workspace resolution', () => {
  it('maps exported package subpaths back to workspace source files', () => {
    const resolution = scanWorkspaceResolution(workspaceRootDir)

    expect(resolution['@bilig/benchmarks/workbook-corpus']).toEqual({
      packageDir: 'packages/benchmarks',
      sourceEntry: 'packages/benchmarks/src/workbook-corpus.ts',
    })
    expect(resolution['@bilig/formula/program-arena']).toEqual({
      packageDir: 'packages/formula',
      sourceEntry: 'packages/formula/src/program-arena.ts',
    })
    expect(resolution['@bilig/headless/xlsx']).toEqual({
      packageDir: 'packages/headless',
      sourceEntry: 'packages/excel-import/src/index.ts',
    })
    expect(resolution['@bilig/workpaper']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/index.ts',
    })
    expect(resolution['@bilig/workpaper/xlsx']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/xlsx.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc']).toEqual({
      packageDir: 'packages/bilig-xlsx-formula-recalc',
      sourceEntry: 'packages/bilig-xlsx-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc/cli-api']).toEqual({
      packageDir: 'packages/bilig-xlsx-formula-recalc',
      sourceEntry: 'packages/bilig-xlsx-formula-recalc/src/cli-api.ts',
    })
    expect(resolution['@bilig/sheetjs-formula-recalc']).toEqual({
      packageDir: 'packages/bilig-sheetjs-formula-recalc',
      sourceEntry: 'packages/bilig-sheetjs-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/exceljs-formula-recalc']).toEqual({
      packageDir: 'packages/bilig-exceljs-formula-recalc',
      sourceEntry: 'packages/bilig-exceljs-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/formula/external-function-adapter']).toEqual({
      packageDir: 'packages/formula',
      sourceEntry: 'packages/formula/src/external-function-adapter.ts',
    })
    expect(resolution['bilig-workpaper']).toEqual({
      packageDir: 'packages/bilig',
      sourceEntry: 'packages/bilig/src/index.ts',
    })
    expect(resolution['bilig-workpaper/xlsx']).toEqual({
      packageDir: 'packages/bilig',
      sourceEntry: 'packages/bilig/src/xlsx.ts',
    })
    expect(resolution['xlsx-formula-recalc']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/index.ts',
    })
    expect(resolution['xlsx-formula-recalc/cli-api']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/cli-api.ts',
    })
    expect(resolution['sheetjs-formula-recalc']).toEqual({
      packageDir: 'packages/sheetjs-formula-recalc',
      sourceEntry: 'packages/sheetjs-formula-recalc/src/index.ts',
    })
    expect(resolution['exceljs-formula-recalc']).toEqual({
      packageDir: 'packages/exceljs-formula-recalc',
      sourceEntry: 'packages/exceljs-formula-recalc/src/index.ts',
    })
  })

  it('orders Vitest aliases with package subpaths before package roots', async () => {
    const { createVitestAliasEntries } = await import('../workspace-resolution.js')

    const aliases = createVitestAliasEntries([], workspaceRootDir)
    const rootIndex = aliases.findIndex((entry) => entry.find === '@bilig/benchmarks')
    const subpathIndex = aliases.findIndex((entry) => entry.find === '@bilig/benchmarks/workbook-corpus')

    expect(subpathIndex).toBeGreaterThanOrEqual(0)
    expect(rootIndex).toBeGreaterThanOrEqual(0)
    expect(subpathIndex).toBeLessThan(rootIndex)
  })
})
