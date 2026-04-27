import { readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, test } from 'vitest'

const GRID_SRC_ROOT = dirname(dirname(fileURLToPath(import.meta.url)))

function collectSourceFiles(path: string): string[] {
  const stat = statSync(path)
  if (stat.isFile()) {
    return path.endsWith('.ts') || path.endsWith('.tsx') ? [path] : []
  }
  return readdirSync(path).flatMap((entry) => collectSourceFiles(join(path, entry)))
}

describe('renderer v3 import boundary', () => {
  test('mounted V3 renderer paths do not import renderer-v2 modules', () => {
    const files = [
      ...collectSourceFiles(join(GRID_SRC_ROOT, 'renderer-v3')),
      join(GRID_SRC_ROOT, 'gridHeaderPanes.ts'),
      join(GRID_SRC_ROOT, 'useWorkbookGridRenderState.ts'),
    ]
    const offenders = files.flatMap((file) => {
      const source = readFileSync(file, 'utf8')
      return source.includes('renderer-v2') ? [relative(GRID_SRC_ROOT, file)] : []
    })

    expect(offenders).toEqual([])
  })
})
