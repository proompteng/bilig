import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  highestPublishedStableSemver,
  highestStableSemver,
  missingPublishedRuntimePackageNames,
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
      { packageName: '@bilig/headless', version: '0.10.0' },
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
          { packageName: '@bilig/headless', version: '0.10.0' },
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
        { packageName: '@bilig/headless', version: '0.10.0' },
      ]),
    ).toEqual(['@bilig/future-runtime'])
  })

  it('blocks automatic publishing when runtime package names are not provisioned on npm', () => {
    const publishedVersions = [
      { packageName: '@bilig/protocol', version: '0.10.1' },
      { packageName: '@bilig/future-runtime', version: null },
      { packageName: '@bilig/headless', version: '0.10.0' },
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
      { packageName: 'sheetjs-formula-recalc', version: null },
      { packageName: 'exceljs-formula-recalc', version: '0.35.1' },
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
      missingPackageNames: ['sheetjs-formula-recalc'],
      reason: 'unprovisioned npm package name(s) will be skipped: sheetjs-formula-recalc',
    })
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
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/headless')
  })

  it('publishes the create-workpaper starter through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/create-workpaper')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/create-workpaper')
    expect(isRuntimeAffectingPath('packages/create-workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/create-workpaper/bin/create-bilig-workpaper.js')).toBe(true)
  })

  it('publishes the unscoped bilig-workpaper package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/bilig')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/bilig')
    expect(isRuntimeAffectingPath('packages/bilig/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped WorkPaper package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/workpaper')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/workpaper')
    expect(isRuntimeAffectingPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/README.md')).toBe(true)
  })

  it('publishes the XLSX formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/xlsx-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/xlsx-formula-recalc')
    expect(isRuntimeAffectingPath('packages/xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped XLSX formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/bilig-xlsx-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/bilig-xlsx-formula-recalc')
    expect(isRuntimeAffectingPath('packages/bilig-xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-xlsx-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-xlsx-formula-recalc/README.md')).toBe(true)
  })

  it('publishes the SheetJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/sheetjs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/sheetjs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped SheetJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/bilig-sheetjs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/bilig-sheetjs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/bilig-sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-sheetjs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-sheetjs-formula-recalc/README.md')).toBe(true)
  })

  it('publishes the ExcelJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/exceljs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/exceljs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/AGENTS.md')).toBe(true)
  })

  it('publishes the scoped ExcelJS formula recalculation package through the common runtime workflow', () => {
    expect(RUNTIME_PACKAGE_DIRS).toContain('packages/bilig-exceljs-formula-recalc')
    expect(RUNTIME_NPM_PACKAGE_DIRS).toContain('packages/bilig-exceljs-formula-recalc')
    expect(isRuntimeAffectingPath('packages/bilig-exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-exceljs-formula-recalc/src/index.ts')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-exceljs-formula-recalc/README.md')).toBe(true)
  })

  it('publishes XLSX import/export through the headless package subpath', () => {
    const manifest = JSON.parse(readFileSync(resolve(repoRoot, 'packages/headless/package.json'), 'utf8'))

    expect(manifest.exports['./xlsx']).toEqual({
      types: './dist/xlsx.d.ts',
      import: './dist/xlsx.js',
    })
    expect(manifest.dependencies).not.toHaveProperty('@bilig/excel-import')
    expect(manifest.dependencies).toMatchObject({
      'fast-xml-parser': '5.7.3',
      fflate: '0.3.11',
      xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz',
      'xlsx-js-style': '1.2.0',
    })
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
    expect(isRuntimePackageContentPath('packages/headless/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/workpaper/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-xlsx-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-sheetjs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/exceljs-formula-recalc/package.json')).toBe(true)
    expect(isRuntimePackageContentPath('packages/bilig-exceljs-formula-recalc/package.json')).toBe(true)
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
    const smokeSource = readFileSync(resolve(repoRoot, 'scripts/workpaper-external-smoke.ts'), 'utf8')

    expect(publishSource).toContain('encodeURIComponent(runtimePackage.name)')
    expect(publishSource).toContain('listTarballsRecursive(targetDir)')
    expect(publishSource).toContain('const tarballsByPackage = indexTarballs(packDir)')
    expect(smokeSource).toContain('encodeURIComponent(runtimePackage.name)')
    expect(smokeSource).toContain('const packagePackDir = join(packDir, encodeURIComponent(runtimePackage.name))')
  })

  it('keeps release-please manifest version in the runtime release sync path', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/sync-runtime-package-versions.ts'), 'utf8')

    expect(source).toContain("manifest['packages/headless'] = version")
    expect(source).toContain('syncReleasePleaseManifestVersion')
  })
})
