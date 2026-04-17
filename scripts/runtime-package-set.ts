import { readFileSync } from 'node:fs'
import { join } from 'node:path'

export const RUNTIME_PACKAGE_DIRS = [
  'packages/protocol',
  'packages/workbook-domain',
  'packages/wasm-kernel',
  'packages/formula',
  'packages/core',
  'packages/headless',
] as const

export interface RuntimePackageManifest {
  dir: (typeof RUNTIME_PACKAGE_DIRS)[number]
  name: string
  version: string
}

export interface StableSemver {
  major: number
  minor: number
  patch: number
}

export function loadRuntimePackages(rootDir: string): RuntimePackageManifest[] {
  return RUNTIME_PACKAGE_DIRS.map((dir) => {
    const manifest = JSON.parse(readFileSync(join(rootDir, dir, 'package.json'), 'utf8'))
    if (typeof manifest.name !== 'string' || typeof manifest.version !== 'string') {
      throw new Error(`Invalid package manifest: ${join(rootDir, dir, 'package.json')}`)
    }
    return {
      dir,
      name: manifest.name,
      version: manifest.version,
    }
  })
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
