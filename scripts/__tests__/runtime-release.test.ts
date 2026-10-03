import { existsSync, readFileSync, statSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  highestPublishedStableSemver,
  highestStableSemver,
  missingPublishedRuntimePackageNames,
  orderRuntimePackagesForPublish,
  parseBooleanEnv,
  planRuntimePackagePublishProvisioning,
  resolvePublishedRuntimePackageBaseline,
  RUNTIME_NPM_PACKAGE_DIRS,
  RUNTIME_PACKAGE_DIRS,
} from '../runtime-package-set.ts'
import {
  bumpVersion,
  isRuntimeAffectingPath,
  isRuntimePackageContentPath,
  parseConventionalCommit,
  releaseTypeForConventionalCommit,
} from '../runtime-release.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

interface PackageManifestWithBins {
  readonly bin?: unknown
  readonly files?: unknown
}

function readPackageManifestWithBins(packageDir: string): PackageManifestWithBins {
  const parsed: unknown = JSON.parse(readFileSync(resolve(repoRoot, packageDir, 'package.json'), 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`Invalid package manifest: ${packageDir}/package.json`)
  }
  return parsed
}

function collectBinTargets(bin: unknown): string[] {
  if (typeof bin === 'string') {
    return [bin]
  }
  if (typeof bin !== 'object' || bin === null || Array.isArray(bin)) {
    return []
  }
  return Object.values(bin).filter((target): target is string => typeof target === 'string')
}

describe('runtime release helpers', () => {
  it('parses standard conventional commits', () => {
    const parsed = parseConventionalCommit({
      subject: 'feat(core): add runtime planner',
      body: '',
    })

    expect(parsed).toEqual({
      type: 'feat',
      scope: 'core',
      description: 'add runtime planner',
      breaking: false,
    })
  })

  it('detects breaking changes from bang markers and footer markers', () => {
    const bang = parseConventionalCommit({
      subject: 'feat(core)!: replace runtime release flow',
      body: '',
    })
    const footer = parseConventionalCommit({
      subject: 'fix(core): preserve publish ordering',
      body: 'BREAKING CHANGE: old release path removed',
    })

    expect(bang?.breaking).toBe(true)
    expect(footer?.breaking).toBe(true)
  })

  it('normalizes git-generated reverts of conventional commits', () => {
    const parsed = parseConventionalCommit({
      subject: 'Revert "perf(headless): reduce formula build overhead"',
      body: 'This reverts commit dc0093f2d909674c859a4fb0fbe8f33c317e23f7.',
    })

    expect(parsed).toEqual({
      type: 'revert',
      scope: 'headless',
      description: 'perf(headless): reduce formula build overhead',
      breaking: false,
    })
    expect(parsed && releaseTypeForConventionalCommit(parsed)).toBe('patch')
  })

  it('rejects git-generated reverts of non-conventional commits', () => {
    expect(
      parseConventionalCommit({
        subject: 'Revert "reduce formula build overhead"',
        body: 'This reverts commit dc0093f2d909674c859a4fb0fbe8f33c317e23f7.',
      }),
    ).toBeNull()
  })

  it('maps conventional commit kinds to semantic release types', () => {
    expect(
      releaseTypeForConventionalCommit({
        type: 'fix',
        scope: null,
        description: 'repair package metadata',
        breaking: false,
      }),
    ).toBe('patch')

    expect(
      releaseTypeForConventionalCommit({
        type: 'feat',
        scope: null,
        description: 'add runtime release planner',
        breaking: false,
      }),
    ).toBe('minor')

    expect(
      releaseTypeForConventionalCommit({
        type: 'refactor',
        scope: null,
        description: 'shuffle internal helpers',
        breaking: false,
      }),
    ).toBe('none')

    expect(
      releaseTypeForConventionalCommit({
        type: 'chore',
        scope: null,
        description: 'drop old release flow',
        breaking: true,
      }),
    ).toBe('major')
  })

  it('bumps semantic versions correctly', () => {
    expect(bumpVersion('0.1.2', 'patch')).toBe('0.1.3')
    expect(bumpVersion('0.1.2', 'minor')).toBe('0.2.0')
    expect(bumpVersion('0.1.2', 'major')).toBe('1.0.0')
  })

  it('uses the highest known runtime version as the publish baseline', () => {
    expect(highestStableSemver(['0.7.8', '0.9.3', '0.1.95'])).toBe('0.9.3')
  })

  it('derives a published runtime baseline from partial package publishing', () => {
    expect(highestPublishedStableSemver(['0.10.1', '0.10.0', null, undefined])).toBe('0.10.1')
    expect(highestPublishedStableSemver([null, undefined])).toBeNull()
  })

  it('allows explicit same-version recovery from a partial runtime publish set', () => {
    const publishedVersions = [
      { packageName: '@bilig/protocol', version: '0.10.1' },
      { packageName: '@bilig/core', version: '0.10.1' },
      { packageName: '@bilig/future-runtime', version: null },
      { packageName: '@bilig/workpaper', version: '0.10.0' },
    ]

    expect(resolvePublishedRuntimePackageBaseline(publishedVersions, { allowPartialPublishedSet: true })).toBe('0.10.1')
    expect(() => resolvePublishedRuntimePackageBaseline(publishedVersions, { allowPartialPublishedSet: false })).toThrow(
      'Published runtime package versions are not aligned',
    )
  })

  it('keeps a new unpublished runtime package on the existing aligned baseline before first release', () => {
    expect(
      resolvePublishedRuntimePackageBaseline(
        [
          { packageName: '@bilig/protocol', version: '0.10.0' },
          { packageName: '@bilig/core', version: '0.10.0' },
          { packageName: '@bilig/future-runtime', version: null },
          { packageName: '@bilig/workpaper', version: '0.10.0' },
        ],
        { allowPartialPublishedSet: false },
      ),
    ).toBe('0.10.0')
  })

  it('identifies missing npm packages before a non-dry-run publish mutates the package set', () => {
    expect(
      missingPublishedRuntimePackageNames([
        { packageName: '@bilig/protocol', version: '0.10.1' },
        { packageName: '@bilig/future-runtime', version: null },
        { packageName: '@bilig/workpaper', version: '0.10.0' },
      ]),
    ).toEqual(['@bilig/future-runtime'])
  })

  it('blocks automatic publishing when runtime package names are not provisioned on npm', () => {
    const publishedVersions = [
      { packageName: '@bilig/protocol', version: '0.10.1' },
      { packageName: '@bilig/future-runtime', version: null },
      { packageName: '@bilig/workpaper', version: '0.10.0' },
    ]

    expect(
      planRuntimePackagePublishProvisioning({
        publishedVersions,
        allowNewNpmPackages: false,
        dryRun: false,
      }),
    ).toEqual({
      publishAllowed: false,
      missingPackageNames: ['@bilig/future-runtime'],
      reason: 'npm package name(s) are not provisioned: @bilig/future-runtime',
    })

    expect(
      planRuntimePackagePublishProvisioning({
        publishedVersions,
        allowNewNpmPackages: true,
        dryRun: false,
      }).publishAllowed,
    ).toBe(true)
  })

  it('allows explicitly skipping unprovisioned leaf packages during partial recovery', () => {
    const publishedVersions = [
      { packageName: '@bilig/protocol', version: '0.36.0' },
      { packageName: '@bilig/sheetjs-formula-recalc', version: null },
      { packageName: '@bilig/exceljs-formula-recalc', version: '0.35.1' },
    ]

    expect(
      planRuntimePackagePublishProvisioning({
        publishedVersions,
        allowNewNpmPackages: false,
        skipUnprovisionedNpmPackages: true,
        dryRun: false,
      }),
    ).toEqual({
      publishAllowed: true,
      missingPackageNames: ['@bilig/sheetjs-formula-recalc'],
      reason: 'unprovisioned npm package name(s) will be skipped: @bilig/sheetjs-formula-recalc',
    })
  })

  it('orders missing and unpublished runtime packages before already-published target versions', () => {
    const runtimePackages = [
      { dir: 'packages/protocol', name: '@bilig/protocol', version: '0.164.0' },
      { dir: 'packages/formula', name: '@bilig/formula', version: '0.164.0' },
      { dir: 'packages/xlsx', name: '@bilig/xlsx', version: '0.164.0' },
      { dir: 'packages/workpaper', name: '@bilig/workpaper', version: '0.164.0' },
    ] as const

    expect(
      orderRuntimePackagesForPublish({
        runtimePackages,
        missingPackageNames: new Set(['@bilig/xlsx']),
        targetVersionPublishedPackageNames: new Set(['@bilig/protocol', '@bilig/formula']),
      }).map((runtimePackage) => runtimePackage.name),
    ).toEqual(['@bilig/xlsx', '@bilig/workpaper', '@bilig/protocol', '@bilig/formula'])
  })

  it('uses strict boolean parsing for runtime publish controls', () => {
    expect(parseBooleanEnv(undefined)).toBe(false)
    expect(parseBooleanEnv('')).toBe(false)
    expect(parseBooleanEnv('0')).toBe(false)
    expect(parseBooleanEnv('false')).toBe(false)
    expect(parseBooleanEnv('1')).toBe(true)
    expect(parseBooleanEnv('true')).toBe(true)

    expect(() => parseBooleanEnv('TRUE')).toThrow('Expected boolean environment value "1", "true", "0", or "false", received TRUE')
    expect(() => parseBooleanEnv('yes')).toThrow('Expected boolean environment value "1", "true", "0", or "false", received yes')
  })

  it('keeps the Excel importer runtime-affecting without requiring standalone npm publication', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/excel-import')
    expect(RUNTIME_NPM_PACKAGE_DIRS).not.toContain('packages/excel-import')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/workpaper')
  })

  it('publishes @bilig/xlsx before dependent runtime packages', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/xlsx')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/xlsx')
    expect(RUNTIME_NPM_PACKAGE_DIRS.indexOf('packages/xlsx')).toBeLessThan(RUNTIME_NPM_PACKAGE_DIRS.indexOf('packages/workpaper'))
    expect(isRuntimeAffectingPath('packages/xlsx/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx/src/index.ts')).toBe(true)
  })

  it('publishes the create-workpaper starter through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/create-workpaper')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/create-workpaper')
    expect(isRuntimeAffectingPath('packages/create-workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/create-workpaper/bin/create-bilig-workpaper.js')).toBe(true)
    expect(isRuntimePackageContentPath('packages/create-workpaper/agent-overlay/.claude/skills/bilig-workpaper/SKILL.md')).toBe(true)
  })

  it('publishes the scoped WorkPaper package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/workpaper')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/workpaper')
    expect(isRuntimeAffectingPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/README.md')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped XLSX formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/xlsx-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/xlsx-formula-recalc')
    expect(isRuntimeAffectingPath('packages/xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/README.md')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped SheetJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/sheetjs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/sheetjs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/README.md')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped ExcelJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/exceljs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/exceljs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/README.md')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('keeps runtime package bins on committed files instead of generated build output', () => {
    for (const packageDir of RUNTIME_NPM_PACKAGE_DIRS) {
      const manifest = readPackageManifestWithBins(packageDir)
      const binTargets = collectBinTargets(manifest.bin)

      for (const target of binTargets) {
        expect(target.startsWith('./dist/'), `${packageDir} exposes a generated bin target: ${target}`).toBe(false)

        const binPath = resolve(repoRoot, packageDir, target)
        expect(existsSync(binPath), `${packageDir} bin target does not exist: ${target}`).toBe(true)
        expect(statSync(binPath).mode & 0o111, `${packageDir} bin target must be executable: ${target}`).not.toBe(0)

        if (target.startsWith('./bin/') && packageDir !== 'packages/create-workpaper') {
          const source = readFileSync(binPath, 'utf8')
          expect(source.startsWith('#!/usr/bin/env node\n'), `${packageDir} bin target must start with a node shebang: ${target}`).toBe(
            true,
          )
          const delegatesToBuiltOutput = source.includes("await import('../dist/")
          const delegatesToSharedEvaluator =
            target.endsWith('/bilig-evaluate.js') && source.includes("await import('@bilig/xlsx-formula-recalc/evaluator')")
          expect(
            delegatesToBuiltOutput || delegatesToSharedEvaluator,
            `${packageDir} bin wrapper must delegate to built output or the shared evaluator: ${target}`,
          ).toBe(true)
        }
      }

      if (binTargets.some((target) => target.startsWith('./bin/'))) {
        expect(manifest.files, `${packageDir} package files must include checked-in bin wrappers`).toContain('bin')
      }
    }
  })

  it('publishes XLSX import/export through the headless package subpath', () => {
    const manifest = JSON.parse(readFileSync(resolve(repoRoot, 'packages/workpaper/package.json'), 'utf8'))

    expect(manifest.exports['./xlsx']).toEqual({
      types: './dist/xlsx.d.ts',
      import: './dist/xlsx.js',
      default: './dist/xlsx.js',
    })
    expect(manifest.dependencies).not.toHaveProperty('@bilig/excel-import')
    expect(manifest.dependencies).toHaveProperty('@bilig/xlsx', 'workspace:*')
    expect(manifest.dependencies).toHaveProperty('fflate')
    expect(manifest.dependencies).not.toHaveProperty('fast-xml-parser')
    expect(manifest.dependencies).not.toHaveProperty('fflate-stream')
    expect(manifest.dependencies).not.toHaveProperty('xlsx')
    expect(manifest.dependencies).not.toHaveProperty('xlsx-js-style')
  })

  it('matches runtime-affecting publish paths', () => {
    expect(isRuntimeAffectingPath('packages/core/src/index.ts')).toBe(true)
    expect(isRuntimeAffectingPath('packages/excel-import/src/index.ts')).toBe(true)
    expect(isRuntimeAffectingPath('scripts/publish-runtime-package-set.ts')).toBe(true)
    expect(isRuntimeAffectingPath('scripts/sync-runtime-package-versions.ts')).toBe(true)
    expect(isRuntimeAffectingPath('scripts/sync-runtime-release-metadata.ts')).toBe(true)
    expect(isRuntimeAffectingPath('apps/web/src/App.tsx')).toBe(false)
  })

  it('separates package content changes from release automation changes', () => {
    expect(isRuntimePackageContentPath('packages/core/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('scripts/plan-runtime-release.ts')).toBe(false)
    expect(isRuntimePackageContentPath('.github/workflows/headless-package.yml')).toBe(false)
  })

  it('requires committed runtime package versions before staging npm packages', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/publish-runtime-package-set.ts'), 'utf8')

    expect(source).toContain('Repository runtime package manifests must be committed')
    expect(source).toContain('manifest.version !== targetVersion')
    expect(source).not.toContain('manifest.version = targetVersion')
  })

  it('isolates packed tarballs when scoped package names share unscoped npm filenames', () => {
    const publishSource = readFileSync(resolve(repoRoot, 'scripts/publish-runtime-package-set.ts'), 'utf8')

    expect(publishSource).toContain('encodeURIComponent(runtimePackage.name)')
    expect(publishSource).toContain('listTarballsRecursive(targetDir)')
    expect(publishSource).toContain('const tarballsByPackage = indexTarballs(packDir)')
    expect(existsSync(resolve(repoRoot, 'scripts', `${['workpaper', 'external', 'smoke'].join('-')}.ts`))).toBe(false)
  })

  it('keeps release-please manifest version in the runtime release sync path', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/sync-runtime-package-versions.ts'), 'utf8')

    expect(source).toContain("manifest['packages/workpaper'] = version")
    expect(source).toContain('syncReleasePleaseManifestVersion')
    expect(source).toContain("join(rootDir, 'gemini-extension.json')")
    expect(source).toContain('syncGeminiExtensionVersion')
    expect(source).toContain('syncMcpDirectoryDocVersion')
  })
})
