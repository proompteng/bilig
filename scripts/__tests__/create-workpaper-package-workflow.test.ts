import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('create-bilig-workpaper package workflow', () => {
  it('keeps the starter package on a verified npm publish path', () => {
    const source = readFileSync(resolve(repoRoot, '.github/workflows/create-workpaper-package.yml'), 'utf8')

    expect(source).toContain('packages/create-workpaper/**')
    expect(source).toContain('scripts/check-create-workpaper-package.ts')
    expect(source).toContain('pnpm create-workpaper:check')
    expect(source).toContain('node packages/create-workpaper/bin/create-bilig-workpaper.js .cache/create-workpaper-ci')
    expect(source).toContain('npm install --ignore-scripts')
    expect(source).toContain('npm run typecheck')
    expect(source).toContain('npm run smoke')
    expect(source).toContain('id-token: write')
    expect(source).toContain('publish_args=(./packages/create-workpaper --tag "$npm_tag" --access public --provenance)')
    expect(source).toContain('npm publish "${publish_args[@]}"')
    expect(source).toContain('allow_new_package')
    expect(source).toContain('Configure npm trusted publishing for this workflow')
  })

  it('keeps create-workpaper package files in the repository', () => {
    for (const path of [
      'packages/create-workpaper/package.json',
      'packages/create-workpaper/README.md',
      'packages/create-workpaper/bin/create-bilig-workpaper.js',
      'packages/create-workpaper/template/package.json',
      'packages/create-workpaper/template/src/index.ts',
      'scripts/check-create-workpaper-package.ts',
    ]) {
      expect(existsSync(resolve(repoRoot, path)), `${path} should exist`).toBe(true)
    }
  })
})
