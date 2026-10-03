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
    expect(resolution['@bilig/workpaper']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/index.ts',
    })
    expect(resolution['@bilig/workpaper/ai-sdk']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/ai-sdk.ts',
    })
    expect(resolution['@bilig/workpaper/evaluator']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/evaluator.ts',
    })
    expect(resolution['@bilig/workpaper/xlsx']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/xlsx.ts',
    })
    expect(resolution['@bilig/xlsx']).toEqual({
      packageDir: 'packages/xlsx',
      sourceEntry: 'packages/xlsx/src/index.ts',
    })
    expect(resolution['@bilig/xlsx/address']).toEqual({
      packageDir: 'packages/xlsx',
      sourceEntry: 'packages/xlsx/src/address.ts',
    })
    expect(resolution['@bilig/xlsx/streaming-native-recalc']).toEqual({
      packageDir: 'packages/xlsx',
      sourceEntry: 'packages/xlsx/src/streaming-native-recalc.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc/cli-api']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/cli-api.ts',
    })
    expect(resolution['@bilig/sheetjs-formula-recalc']).toEqual({
      packageDir: 'packages/sheetjs-formula-recalc',
      sourceEntry: 'packages/sheetjs-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/exceljs-formula-recalc']).toEqual({
      packageDir: 'packages/exceljs-formula-recalc',
      sourceEntry: 'packages/exceljs-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/formula/external-function-adapter']).toEqual({
      packageDir: 'packages/formula',
      sourceEntry: 'packages/formula/src/external-function-adapter.ts',
    })
    expect(resolution['@bilig/workpaper']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/index.ts',
    })
    expect(resolution['@bilig/workpaper/evaluator']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/evaluator.ts',
    })
    expect(resolution['@bilig/workpaper/xlsx']).toEqual({
      packageDir: 'packages/workpaper',
      sourceEntry: 'packages/workpaper/src/xlsx.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/xlsx-formula-recalc/cli-api']).toEqual({
      packageDir: 'packages/xlsx-formula-recalc',
      sourceEntry: 'packages/xlsx-formula-recalc/src/cli-api.ts',
    })
    expect(resolution['@bilig/sheetjs-formula-recalc']).toEqual({
      packageDir: 'packages/sheetjs-formula-recalc',
      sourceEntry: 'packages/sheetjs-formula-recalc/src/index.ts',
    })
    expect(resolution['@bilig/exceljs-formula-recalc']).toEqual({
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
