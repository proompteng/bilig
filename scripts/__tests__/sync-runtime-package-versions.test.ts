import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { RUNTIME_PACKAGE_DIRS } from '../runtime-package-set.ts'
import { syncRuntimePackageVersions } from '../sync-runtime-package-versions.ts'

describe('syncRuntimePackageVersions', () => {
  it('aligns runtime package manifests and headless MCP metadata', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-versions-'))

    for (const packageDir of RUNTIME_PACKAGE_DIRS) {
      const absoluteDir = join(rootDir, packageDir)
      mkdirSync(absoluteDir, { recursive: true })
      writeFileSync(
        join(absoluteDir, 'package.json'),
        `${JSON.stringify({ name: packageNameForDir(packageDir), version: '0.1.95' }, null, 2)}\n`,
      )
    }

    writeFileSync(join(rootDir, '.release-please-manifest.json'), `${JSON.stringify({ 'packages/headless': '0.1.95' }, null, 2)}\n`)

    writeFileSync(
      join(rootDir, 'packages/headless/server.json'),
      `${JSON.stringify(
        {
          name: 'io.github.proompteng.bilig',
          version: '0.1.95',
          remotes: [
            {
              type: 'streamable-http',
              url: 'https://bilig.proompteng.ai/mcp',
            },
          ],
          packages: [
            {
              registryType: 'npm',
              identifier: '@bilig/headless',
              version: '0.1.95',
            },
          ],
        },
        null,
        2,
      )}\n`,
    )

    const result = syncRuntimePackageVersions({ rootDir, version: '0.14.14' })

    expect(result.updatedPackages).toEqual(RUNTIME_PACKAGE_DIRS.map(packageNameForDir))
    expect(result.updatedFiles).toHaveLength(RUNTIME_PACKAGE_DIRS.length + 2)

    for (const packageDir of RUNTIME_PACKAGE_DIRS) {
      const manifest = JSON.parse(readFileSync(join(rootDir, packageDir, 'package.json'), 'utf8'))
      expect(manifest.version).toBe('0.14.14')
    }

    const serverJson = JSON.parse(readFileSync(join(rootDir, 'packages/headless/server.json'), 'utf8'))
    expect(serverJson.version).toBe('0.14.14')
    expect(serverJson.remotes[0]).toEqual({
      type: 'streamable-http',
      url: 'https://bilig.proompteng.ai/mcp',
    })
    expect(serverJson.packages[0].version).toBe('0.14.14')

    const releasePleaseManifest = JSON.parse(readFileSync(join(rootDir, '.release-please-manifest.json'), 'utf8'))
    expect(releasePleaseManifest['packages/headless']).toBe('0.14.14')
  })

  it('rejects non-stable semver versions before writing files', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-versions-'))
    mkdirSync(join(rootDir, 'packages/protocol'), { recursive: true })

    expect(() => syncRuntimePackageVersions({ rootDir, version: '0.14.14-beta.1' })).toThrow('Expected stable semver version')
  })
})

function packageNameForDir(packageDir: string): string {
  if (packageDir === 'packages/xlsx-formula-recalc') {
    return 'xlsx-formula-recalc'
  }
  if (packageDir === 'packages/exceljs-formula-recalc') {
    return 'exceljs-formula-recalc'
  }
  return `@bilig/${packageDir.split('/').at(-1) ?? packageDir}`
}
