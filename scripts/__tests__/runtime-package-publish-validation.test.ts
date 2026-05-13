import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { syncStagedMcpServerMetadata, validateStagedMcpServerMetadata } from '../runtime-package-mcp-metadata.ts'
import { validateStagedRuntimePackageVersion } from '../runtime-package-publish-validation.ts'

const tempDirs: string[] = []

afterEach(() => {
  while (tempDirs.length > 0) {
    const tempDir = tempDirs.pop()
    if (tempDir) {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }
})

describe('runtime package publish validation', () => {
  it('accepts a staged headless package whose WorkPaper.version follows package.json', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/headless', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('rejects a staged headless package with the old hardcoded WorkPaper.version behavior', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      versionModuleSource: "export const WORKPAPER_VERSION = '0.1.95'\n",
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/headless', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/headless WorkPaper.version does not match package version',
    )
  })

  it('rewrites staged MCP server metadata to the package release version', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverVersion: '0.1.95',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    syncStagedMcpServerMetadata('@bilig/headless', stagedPackageDir, '9.9.9')

    expect(() => validateStagedMcpServerMetadata('@bilig/headless', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('rejects stale MCP server metadata in a staged headless package', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverVersion: '0.1.95',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/headless', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/headless server.json version must match package version',
    )
  })

  it('does not require WorkPaper metadata from other runtime packages', () => {
    expect(() => validateStagedRuntimePackageVersion('@bilig/core', '/missing-package-dir', '9.9.9')).not.toThrow()
  })
})

function stageHeadlessPackage(args: {
  readonly manifestVersion: string
  readonly serverVersion?: string
  readonly versionModuleSource: string
}): string {
  const stagedPackageDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-version-validation-'))
  tempDirs.push(stagedPackageDir)
  mkdirSync(join(stagedPackageDir, 'dist'), { recursive: true })
  const serverVersion = args.serverVersion ?? args.manifestVersion
  writeFileSync(
    join(stagedPackageDir, 'package.json'),
    `${JSON.stringify({
      name: '@bilig/headless',
      version: args.manifestVersion,
      mcpName: 'io.github.proompteng/bilig-workpaper',
    })}\n`,
  )
  writeFileSync(
    join(stagedPackageDir, 'server.json'),
    `${JSON.stringify(
      {
        name: 'io.github.proompteng/bilig-workpaper',
        version: serverVersion,
        packages: [
          {
            registryType: 'npm',
            identifier: '@bilig/headless',
            version: serverVersion,
            transport: {
              type: 'stdio',
            },
          },
        ],
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(join(stagedPackageDir, 'dist/work-paper-version.js'), args.versionModuleSource)
  return stagedPackageDir
}

function packageManifestVersionModuleSource(): string {
  return [
    "import { createRequire } from 'node:module'",
    'const requirePackageJson = createRequire(import.meta.url)',
    "const packageManifest = requirePackageJson('../package.json')",
    'export const WORKPAPER_VERSION = packageManifest.version',
    '',
  ].join('\n')
}
