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
  test('product grid paths do not import renderer-v2 modules', () => {
    const files = collectSourceFiles(GRID_SRC_ROOT).filter((file) => {
      const local = relative(GRID_SRC_ROOT, file)
      return !local.startsWith('__tests__/') && !local.startsWith('renderer-v2/')
    })
    const offenders = files.flatMap((file) => {
      const source = readFileSync(file, 'utf8')
      return source.includes('renderer-v2') ? [relative(GRID_SRC_ROOT, file)] : []
    })

    expect(offenders).toEqual([])
  })

  test('V3 pane renderer shell does not own renderer readiness in React state', () => {
    const source = readFileSync(join(GRID_SRC_ROOT, 'renderer-v3/WorkbookPaneRendererV3.tsx'), 'utf8')

    expect(source).not.toContain('useState')
  })
})
