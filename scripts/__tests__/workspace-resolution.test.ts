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
