import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join, relative } from 'node:path'
import { describe, expect, it } from 'vitest'
import { workspaceRootDir } from '../workspace-resolution.js'

const WEB_SRC_ROOT = join(workspaceRootDir, 'apps/web/src')

function collectSourceFiles(path: string): string[] {
  const stat = statSync(path)
  if (stat.isFile()) {
    return path.endsWith('.ts') || path.endsWith('.tsx') ? [path] : []
  }
  return readdirSync(path).flatMap((entry) => collectSourceFiles(join(path, entry)))
}

describe('web import boundary', () => {
  it('keeps production web runtime imports off benchmark package source paths', () => {
    const files = collectSourceFiles(WEB_SRC_ROOT).filter((file) => !relative(WEB_SRC_ROOT, file).startsWith('__tests__/'))
    const offenders = files.flatMap((file) => {
      const source = readFileSync(file, 'utf8')
      return /from\s+['"](?:\.\.\/)+packages\/benchmarks\/src\//.test(source) ||
        /import\(\s*['"](?:\.\.\/)+packages\/benchmarks\/src\//.test(source)
        ? [relative(WEB_SRC_ROOT, file)]
        : []
    })

    expect(offenders).toEqual([])
  })
})
