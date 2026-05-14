import { readFileSync } from 'node:fs'
import { join } from 'node:path'

export const RUNTIME_PACKAGE_DIRS = [
  'packages/protocol',
  'packages/workbook-domain',
  'packages/wasm-kernel',
  'packages/formula',
  'packages/core',
  'packages/excel-import',
  'packages/headless',
] as const

export type RuntimePackageDir = (typeof RUNTIME_PACKAGE_DIRS)[number]

export const RUNTIME_NPM_PACKAGE_DIRS = [
  'packages/protocol',
  'packages/workbook-domain',
  'packages/wasm-kernel',
  'packages/formula',
  'packages/core',
  'packages/headless',
] as const satisfies readonly RuntimePackageDir[]

export interface RuntimePackageManifest {
  dir: RuntimePackageDir
  name: string
  version: string
}

export interface RuntimePackagePublishedVersion {
  packageName: string
  version: string | null
}

export interface RuntimePackagePublishProvisioningPlan {
  publishAllowed: boolean
  missingPackageNames: string[]
  reason: string
}

export interface StableSemver {
  major: number
  minor: number
  patch: number
}

export function loadRuntimePackages(rootDir: string): RuntimePackageManifest[] {
  return RUNTIME_PACKAGE_DIRS.map((dir) => loadRuntimePackageManifest(rootDir, dir))
}

export function loadRuntimeNpmPackages(rootDir: string): RuntimePackageManifest[] {
  return RUNTIME_NPM_PACKAGE_DIRS.map((dir) => loadRuntimePackageManifest(rootDir, dir))
}

function loadRuntimePackageManifest(rootDir: string, dir: RuntimePackageDir): RuntimePackageManifest {
  const manifest = JSON.parse(readFileSync(join(rootDir, dir, 'package.json'), 'utf8'))
  if (typeof manifest.name !== 'string' || typeof manifest.version !== 'string') {
    throw new Error(`Invalid package manifest: ${join(rootDir, dir, 'package.json')}`)
  }
  return {
    dir,
    name: manifest.name,
    version: manifest.version,
  }
}

export function assertAlignedVersions(runtimePackages: RuntimePackageManifest[]): string {
  const versions = [...new Set(runtimePackages.map((entry) => entry.version))]
  if (versions.length !== 1) {
    throw new Error(
      `Runtime npm package versions must stay aligned (${runtimePackages.map((entry) => `${entry.name}@${entry.version}`).join(', ')})`,
    )
  }
  const [version] = versions
  if (!version) {
    throw new Error('Runtime package set is empty')
  }
  return version
}

export function parseBooleanEnv(value: string | undefined): boolean {
  switch (value) {
    case undefined:
    case '':
    case 'false':
    case 'False':
    case 'FALSE':
      return false
    case 'true':
    case 'True':
    case 'TRUE':
      return true
    default:
      throw new Error(`Expected boolean environment value, received ${value}`)
  }
}

export function determineRuntimeReleaseVersion(options: {
  autoIncrement: boolean
  manifestVersion: string
  publishedVersion: string | null
}): string {
  const { autoIncrement, manifestVersion, publishedVersion } = options

  if (!autoIncrement || publishedVersion === null) {
    return manifestVersion
  }

  const comparison = compareStableSemver(manifestVersion, publishedVersion)
  if (comparison > 0) {
    return manifestVersion
  }

  return incrementPatch(publishedVersion)
}

export function highestStableSemver(versions: readonly (string | null | undefined)[]): string {
  const parsedVersions = versions.flatMap((version) => (version ? [version] : []))
  if (parsedVersions.length === 0) {
    throw new Error('Expected at least one semantic version')
  }
  return parsedVersions.reduce((highest, version) => (compareStableSemver(version, highest) > 0 ? version : highest))
}

export function highestPublishedStableSemver(versions: readonly (string | null | undefined)[]): string | null {
  const parsedVersions = versions.flatMap((version) => (version ? [version] : []))
  return parsedVersions.length > 0 ? highestStableSemver(parsedVersions) : null
}

export function compareStableSemver(left: string, right: string): number {
  const leftVersion = parseStableSemver(left)
  const rightVersion = parseStableSemver(right)

  if (leftVersion.major !== rightVersion.major) {
    return leftVersion.major - rightVersion.major
  }
  if (leftVersion.minor !== rightVersion.minor) {
    return leftVersion.minor - rightVersion.minor
  }
  return leftVersion.patch - rightVersion.patch
}

export function incrementPatch(version: string): string {
  const parsed = parseStableSemver(version)
  return `${parsed.major}.${parsed.minor}.${parsed.patch + 1}`
}

export function incrementMinor(version: string): string {
  const parsed = parseStableSemver(version)
  return `${parsed.major}.${parsed.minor + 1}.0`
}

export function incrementMajor(version: string): string {
  const parsed = parseStableSemver(version)
  return `${parsed.major + 1}.0.0`
}

export function parseStableSemver(version: string): StableSemver {
  const match = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.exec(version)
  if (!match) {
    throw new Error(`Expected stable semver version, received ${version}`)
  }
  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
  }
}

export function formatRuntimePackagePublishedVersions(publishedVersions: readonly RuntimePackagePublishedVersion[]): string {
  return publishedVersions.map((entry) => `${entry.packageName}@${entry.version ?? 'unpublished'}`).join(', ')
}

export function missingPublishedRuntimePackageNames(publishedVersions: readonly RuntimePackagePublishedVersion[]): string[] {
  return publishedVersions.filter((entry) => entry.version === null).map((entry) => entry.packageName)
}

export function planRuntimePackagePublishProvisioning(options: {
  publishedVersions: readonly RuntimePackagePublishedVersion[]
  allowNewNpmPackages: boolean
  dryRun: boolean
}): RuntimePackagePublishProvisioningPlan {
  const missingPackageNames = missingPublishedRuntimePackageNames(options.publishedVersions)
  if (options.dryRun) {
    return {
      publishAllowed: true,
      missingPackageNames,
      reason: 'dry-run publishing does not require pre-provisioned npm package names',
    }
  }
  if (options.allowNewNpmPackages) {
    return {
      publishAllowed: true,
      missingPackageNames,
      reason: 'first-time npm package creation was explicitly enabled',
    }
  }
  if (missingPackageNames.length === 0) {
    return {
      publishAllowed: true,
      missingPackageNames,
      reason: 'all runtime package names are provisioned on npm',
    }
  }
  return {
    publishAllowed: false,
    missingPackageNames,
    reason: `npm package name(s) are not provisioned: ${missingPackageNames.join(', ')}`,
  }
}

export function resolvePublishedRuntimePackageBaseline(
  publishedVersions: readonly RuntimePackagePublishedVersion[],
  options: { allowPartialPublishedSet: boolean },
): string | null {
  const versions = publishedVersions
    .map((entry) => entry.version)
    .filter((entry): entry is string => typeof entry === 'string' && entry.length > 0)
  const versionSet = new Set(versions)
  if (versionSet.size === 0) {
    return null
  }
  if (versionSet.size > 1) {
    if (!options.allowPartialPublishedSet) {
      throw new Error(`Published runtime package versions are not aligned (${formatRuntimePackagePublishedVersions(publishedVersions)})`)
    }
    return highestStableSemver([...versionSet])
  }
  const [publishedVersion] = versionSet
  return publishedVersion ?? null
}
