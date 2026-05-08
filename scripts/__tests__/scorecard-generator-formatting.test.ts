import { readFileSync, readdirSync } from 'node:fs'
import { basename, join } from 'node:path'
import { describe, expect, it } from 'vitest'

const scriptsDir = new URL('..', import.meta.url).pathname

describe('scorecard generator formatting', () => {
  it('uses the shared deterministic JSON formatter instead of generator-local oxfmt shells', () => {
    const generatorFiles = readdirSync(scriptsDir)
      .filter((fileName) => fileName.startsWith('gen-') && fileName.endsWith('.ts'))
      .map((fileName) => join(scriptsDir, fileName))

    const generatorLocalFormatters = generatorFiles
      .map((filePath) => ({
        fileName: basename(filePath),
        source: readFileSync(filePath, 'utf8'),
      }))
      .filter(({ source }) => source.includes('function formatJsonForRepo') || source.includes("node_modules', '.bin', 'oxfmt'"))
      .map(({ fileName }) => fileName)

    expect(generatorLocalFormatters).toEqual([])
  })
})
