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

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('rejects a staged headless package with the old hardcoded WorkPaper.version behavior', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      versionModuleSource: "export const WORKPAPER_VERSION = '0.1.95'\n",
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/workpaper WorkPaper.version does not match package version',
    )
  })

  it('rewrites staged MCP server metadata to the package release version', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverVersion: '0.1.95',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    syncStagedMcpServerMetadata('@bilig/workpaper', stagedPackageDir, '9.9.9')

    expect(() => validateStagedMcpServerMetadata('@bilig/workpaper', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('accepts a staged WorkPaper package with MCP metadata', () => {
    const stagedPackageDir = stageMcpPackage({
      packageName: '@bilig/workpaper',
      manifestVersion: '9.9.9',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('rejects stale MCP server metadata in a staged headless package', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverVersion: '0.1.95',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/workpaper server.json version must match package version',
    )
  })

  it('rejects MCP registry descriptions over the published registry limit', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverDescription:
        'Headless spreadsheet WorkPaper tools for formula-backed workbook readback, validated input edits, and JSON persistence.',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/workpaper server.json description must be a string no longer than 100 characters',
    )
  })

  it('rejects MCP registry metadata that overclaims workbook compatibility', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverDescription: 'Excel compatibility checker with a compatibility score for workbook risk.',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/workpaper server.json contains overclaiming MCP Registry wording',
    )
  })

  it('accepts diagnostic MCP registry workbook risk wording', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      serverDescription: 'Formula readback, input edits, JSON persistence, and workbook risk preflight for agents.',
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).not.toThrow()
  })

  it('rejects MCP metadata without the hosted remote endpoint', () => {
    const stagedPackageDir = stageHeadlessPackage({
      manifestVersion: '9.9.9',
      includeRemote: false,
      versionModuleSource: packageManifestVersionModuleSource(),
    })

    expect(() => validateStagedRuntimePackageVersion('@bilig/workpaper', stagedPackageDir, '9.9.9')).toThrow(
      'Staged @bilig/workpaper server.json must include the hosted Streamable HTTP remote endpoint',
    )
  })

  it('does not require WorkPaper metadata from other runtime packages', () => {
    expect(() => validateStagedRuntimePackageVersion('@bilig/core', '/missing-package-dir', '9.9.9')).not.toThrow()
  })
})

function stageHeadlessPackage(args: {
  readonly manifestVersion: string
  readonly includeRemote?: boolean
  readonly serverDescription?: string
  readonly serverVersion?: string
  readonly versionModuleSource: string
}): string {
  return stageMcpPackage({
    packageName: '@bilig/workpaper',
    manifestVersion: args.manifestVersion,
    includeRemote: args.includeRemote,
    serverDescription: args.serverDescription,
    serverVersion: args.serverVersion,
    versionModuleSource: args.versionModuleSource,
  })
}

function stageMcpPackage(args: {
  readonly packageName: '@bilig/workpaper'
  readonly manifestVersion: string
  readonly includeRemote?: boolean
  readonly serverDescription?: string
  readonly serverVersion?: string
  readonly versionModuleSource?: string
}): string {
  const stagedPackageDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-version-validation-'))
  tempDirs.push(stagedPackageDir)
  mkdirSync(join(stagedPackageDir, 'dist'), { recursive: true })
  const serverVersion = args.serverVersion ?? args.manifestVersion
  writeFileSync(
    join(stagedPackageDir, 'package.json'),
    `${JSON.stringify({
      name: args.packageName,
      version: args.manifestVersion,
      mcpName: 'io.github.proompteng/bilig-workpaper',
    })}\n`,
  )
  writeFileSync(
    join(stagedPackageDir, 'server.json'),
    `${JSON.stringify(
      {
        name: 'io.github.proompteng/bilig-workpaper',
        description: args.serverDescription ?? 'Formula readback, input edits, JSON persistence, and workbook risk preflight for agents.',
        version: serverVersion,
        ...(args.includeRemote === false
          ? {}
          : {
              remotes: [
                {
                  type: 'streamable-http',
                  url: 'https://bilig.proompteng.ai/mcp',
                },
              ],
            }),
        packages: [
          {
            registryType: 'npm',
            identifier: args.packageName,
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
  if (args.versionModuleSource) {
    writeFileSync(join(stagedPackageDir, 'dist/work-paper-version.js'), args.versionModuleSource)
  }
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
