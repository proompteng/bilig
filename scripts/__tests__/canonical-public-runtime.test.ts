import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { buildA1WorkPaper, createWorkPaperFromDocument, parseWorkPaperDocument } from '@bilig/workpaper'
import { loadRuntimeNpmPackages } from '../runtime-package-set.ts'
import { buildClaudeCodeMcpConfig } from '../agent-discovery-mcp-configs.ts'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)

describe('canonical public workbook runtime', () => {
  it('publishes one WorkPaper implementation and scoped recalc adapters', () => {
    const names = loadRuntimeNpmPackages(repoRoot).map((entry) => entry.name)
    expect(names).toEqual([
      '@bilig/protocol',
      '@bilig/formula',
      '@bilig/workbook',
      '@bilig/wasm-kernel',
      '@bilig/xlsx',
      '@bilig/core',
      '@bilig/workpaper',
      '@bilig/xlsx-formula-recalc',
      '@bilig/sheetjs-formula-recalc',
      '@bilig/exceljs-formula-recalc',
      '@bilig/create-workpaper',
    ])
  })

  it('owns the browser and CLI entrypoints without a package forwarding chain', () => {
    const manifest = JSON.parse(readFileSync(resolve(repoRoot, 'packages/workpaper/package.json'), 'utf8'))
    expect(manifest.exports).toHaveProperty('./browser')
    expect(manifest.exports).toHaveProperty('./cli')
    expect(manifest.dependencies).toHaveProperty('@bilig/core')
    expect(manifest.dependencies).toHaveProperty('@bilig/formula')
    expect(manifest.dependencies).not.toHaveProperty('@bilig/workpaper')
  })

  it('keeps editor identifiers separate from the scoped npm runtime', () => {
    const extension = JSON.parse(readFileSync(resolve(repoRoot, 'gemini-extension.json'), 'utf8'))
    expect(extension.name).toBe('bilig-workpaper')
    expect(Object.keys(extension.mcpServers)).toEqual(['bilig-workpaper'])
    expect(extension.mcpServers['bilig-workpaper'].args).toContain('@bilig/workpaper@latest')
    const config = JSON.parse(
      buildClaudeCodeMcpConfig({
        workpaperPackageSpec: '@bilig/workpaper@latest',
        remoteMcpEndpoint: 'https://bilig.proompteng.ai/mcp',
        repositoryUrl: 'https://github.com/proompteng/bilig',
        siteRoot: 'https://proompteng.github.io/bilig',
      }),
    )
    expect(Object.keys(config.mcpServers)).toEqual(['bilig-workpaper'])
    expect(config.mcpServers['bilig-workpaper'].args).toContain('@bilig/workpaper@latest')
  })

  it('recalculates and restores a formula through the public A1 API', () => {
    const book = buildA1WorkPaper({
      Inputs: [
        ['Units', 40],
        ['Price', 1200],
      ],
      Results: [['Revenue', '=Inputs!B1*Inputs!B2']],
    })
    try {
      expect(book.get('Results!B1')).toMatchObject({ value: 48_000 })
      const proof = book.setCellAndReadback('Inputs!B1', 48, { readbackRange: 'Results!B1' })
      expect(proof.verified).toBe(true)
      expect(proof.afterReadback.values).toEqual([[{ tag: 1, value: 57_600 }]])
      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(book.serialize()))
      try {
        expect(restored.getCellValue({ sheet: restored.getSheetId('Results')!, row: 0, col: 1 })).toMatchObject({ value: 57_600 })
      } finally {
        restored.dispose()
      }
    } finally {
      book.dispose()
    }
  })
})
