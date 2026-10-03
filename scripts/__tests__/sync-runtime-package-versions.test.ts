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
        `${JSON.stringify(
          {
            name: packageNameForDir(packageDir),
            version: '0.1.95',
            ...(packageDir === 'packages/workpaper' ? { mcpName: 'io.github.proompteng/bilig-workpaper' } : {}),
          },
          null,
          2,
        )}\n`,
      )
    }

    writeFileSync(join(rootDir, '.release-please-manifest.json'), `${JSON.stringify({ 'packages/workpaper': '0.1.95' }, null, 2)}\n`)
    writeFileSync(join(rootDir, 'Dockerfile'), 'ARG BILIG_WORKPAPER_VERSION=0.1.95\n')
    writeFileSync(
      join(rootDir, 'gemini-extension.json'),
      `${JSON.stringify(
        {
          name: 'bilig-workpaper',
          version: '0.1.95',
          contextFileName: 'gemini-workpaper-context.md',
        },
        null,
        2,
      )}\n`,
    )
    const xlsxCacheDoctorAction = [
      'name: XLSX Cache Doctor',
      'inputs:',
      '  workbook:',
      "    default: ''",
      '  package-version:',
      '    description: npm version or dist-tag for @bilig/xlsx-formula-recalc. Pin this for production workflows.',
      '    required: false',
      "    default: '0.1.95'",
      '',
    ].join('\n')
    writeFileSync(join(rootDir, 'action.yml'), xlsxCacheDoctorAction)
    mkdirSync(join(rootDir, 'actions/xlsx-cache-doctor'), { recursive: true })
    writeFileSync(join(rootDir, 'actions/xlsx-cache-doctor/action.yml'), xlsxCacheDoctorAction)
    mkdirSync(join(rootDir, 'docs'), { recursive: true })
    writeFileSync(join(rootDir, 'README.md'), ["          package-version: '0.1.95'", "    package-version: '0.1.95'", ''].join('\n'))
    writeFileSync(
      join(rootDir, 'docs/xlsx-cache-doctor-github-action.md'),
      [
        "          package-version: '0.1.95'",
        "    package-version: '0.1.95'",
        '| `package-version`    | 0.1.95 | npm version or dist-tag for `@bilig/xlsx-formula-recalc`. Pin this in production. |',
        '',
      ].join('\n'),
    )
    writeFileSync(
      join(rootDir, 'docs/mcp-spreadsheet-server-directory.md'),
      [
        'Latest checked result on June 1, 2026: Live, and the Registry latest marker is live but currently trails the current package.',
        'The official Registry latest-marked server `io.github.proompteng/bilig-workpaper` is version',
        '`0.1.90`, package `@bilig/workpaper` is version `0.1.90`, and the current',
        'repo package version is `0.1.95`. The latest-marked entry was updated at `2026-06-01T00:56:47.948741Z`.',
        '',
      ].join('\n'),
    )
    const agentEvaluatorDoc = ['{', '  "@bilig/workpaper": "0.1.95",', '  "@bilig/xlsx-formula-recalc": "0.1.95"', '}', ''].join('\n')
    writeFileSync(join(rootDir, 'docs/agent-adoption-kit.md'), agentEvaluatorDoc)
    writeFileSync(join(rootDir, 'docs/eval-agent-mcp.md'), agentEvaluatorDoc)
    writeFileSync(
      join(rootDir, 'docs/eval-workpaper-service.md'),
      ['This transcript was captured against `@bilig/workpaper@0.1.95`.', '{', '  "@bilig/workpaper": "0.1.95"', '}', ''].join('\n'),
    )
    writeFileSync(join(rootDir, 'packages/workpaper/README.md'), ['{', '  "@bilig/workpaper": "0.1.95"', '}', ''].join('\n'))
    writeFileSync(
      join(rootDir, 'packages/workpaper/server.json'),
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
              identifier: '@bilig/workpaper',
              version: '0.1.95',
            },
          ],
        },
        null,
        2,
      )}\n`,
    )
    writeFileSync(
      join(rootDir, 'packages/workpaper/server.json'),
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
              identifier: '@bilig/workpaper',
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
    expect(result.updatedFiles).toHaveLength(RUNTIME_PACKAGE_DIRS.length + 13)

    for (const packageDir of RUNTIME_PACKAGE_DIRS) {
      const manifest = JSON.parse(readFileSync(join(rootDir, packageDir, 'package.json'), 'utf8'))
      expect(manifest.version).toBe('0.14.14')
    }

    const serverJson = JSON.parse(readFileSync(join(rootDir, 'packages/workpaper/server.json'), 'utf8'))
    expect(serverJson.version).toBe('0.14.14')
    expect(serverJson.remotes[0]).toEqual({
      type: 'streamable-http',
      url: 'https://bilig.proompteng.ai/mcp',
    })
    expect(serverJson.packages[0].version).toBe('0.14.14')

    const workpaperServerJson = JSON.parse(readFileSync(join(rootDir, 'packages/workpaper/server.json'), 'utf8'))
    expect(workpaperServerJson.version).toBe('0.14.14')
    expect(workpaperServerJson.remotes[0]).toEqual({
      type: 'streamable-http',
      url: 'https://bilig.proompteng.ai/mcp',
    })
    expect(workpaperServerJson.packages[0].version).toBe('0.14.14')

    const releasePleaseManifest = JSON.parse(readFileSync(join(rootDir, '.release-please-manifest.json'), 'utf8'))
    expect(releasePleaseManifest['packages/workpaper']).toBe('0.14.14')
    expect(readFileSync(join(rootDir, 'Dockerfile'), 'utf8')).toBe('ARG BILIG_WORKPAPER_VERSION=0.14.14\n')
    const geminiExtension = JSON.parse(readFileSync(join(rootDir, 'gemini-extension.json'), 'utf8'))
    expect(geminiExtension.version).toBe('0.14.14')
    expect(readFileSync(join(rootDir, 'action.yml'), 'utf8')).toContain("default: '0.14.14'")
    expect(readFileSync(join(rootDir, 'actions/xlsx-cache-doctor/action.yml'), 'utf8')).toContain("default: '0.14.14'")
    expect(readFileSync(join(rootDir, 'README.md'), 'utf8')).toContain("package-version: '0.14.14'")
    expect(readFileSync(join(rootDir, 'README.md'), 'utf8')).not.toContain("package-version: '0.1.95'")
    expect(readFileSync(join(rootDir, 'docs/xlsx-cache-doctor-github-action.md'), 'utf8')).toContain("package-version: '0.14.14'")
    expect(readFileSync(join(rootDir, 'docs/xlsx-cache-doctor-github-action.md'), 'utf8')).toContain('| `package-version`    | 0.14.14 |')
    expect(readFileSync(join(rootDir, 'docs/agent-adoption-kit.md'), 'utf8')).toContain('"@bilig/workpaper": "0.14.14"')
    expect(readFileSync(join(rootDir, 'docs/eval-agent-mcp.md'), 'utf8')).toContain('"@bilig/xlsx-formula-recalc": "0.14.14"')
    expect(readFileSync(join(rootDir, 'docs/eval-workpaper-service.md'), 'utf8')).toContain('@bilig/workpaper@0.14.14')
    expect(readFileSync(join(rootDir, 'docs/eval-workpaper-service.md'), 'utf8')).toContain('"@bilig/workpaper": "0.14.14"')
    expect(readFileSync(join(rootDir, 'packages/workpaper/README.md'), 'utf8')).toContain('"@bilig/workpaper": "0.14.14"')
    const mcpDirectoryDoc = readFileSync(join(rootDir, 'docs/mcp-spreadsheet-server-directory.md'), 'utf8')
    expect(mcpDirectoryDoc).toContain('package `@bilig/workpaper` is version `0.1.90`')
    expect(mcpDirectoryDoc).toContain('the current repo package version is `0.14.14`')
  })

  it('rejects non-stable semver versions before writing files', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-versions-'))
    mkdirSync(join(rootDir, 'packages/protocol'), { recursive: true })

    expect(() => syncRuntimePackageVersions({ rootDir, version: '0.14.14-beta.1' })).toThrow('Expected stable semver version')
  })
})

function packageNameForDir(packageDir: string): string {
  if (packageDir === 'packages/workpaper') {
    return '@bilig/workpaper'
  }
  if (packageDir === 'packages/workpaper') {
    return '@bilig/workpaper'
  }
  if (packageDir === 'packages/xlsx-formula-recalc') {
    return '@bilig/xlsx-formula-recalc'
  }
  if (packageDir === 'packages/xlsx-formula-recalc') {
    return '@bilig/xlsx-formula-recalc'
  }
  if (packageDir === 'packages/sheetjs-formula-recalc') {
    return '@bilig/sheetjs-formula-recalc'
  }
  if (packageDir === 'packages/sheetjs-formula-recalc') {
    return '@bilig/sheetjs-formula-recalc'
  }
  if (packageDir === 'packages/exceljs-formula-recalc') {
    return '@bilig/exceljs-formula-recalc'
  }
  if (packageDir === 'packages/exceljs-formula-recalc') {
    return '@bilig/exceljs-formula-recalc'
  }
  return `@bilig/${packageDir.split('/').at(-1) ?? packageDir}`
}
