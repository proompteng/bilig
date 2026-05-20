#!/usr/bin/env bun

import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync, cpSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join, resolve } from 'node:path'

import {
  assertAlignedVersions,
  formatRuntimePackagePublishedVersions,
  loadRuntimeNpmPackages,
  loadRuntimePackages,
  parseStableSemver,
  planRuntimePackagePublishProvisioning,
  parseBooleanEnv,
  missingPublishedRuntimePackageNames,
  type RuntimePackagePublishedVersion,
  type RuntimePackageManifest,
} from './runtime-package-set.ts'
import { syncStagedMcpServerMetadata } from './runtime-package-mcp-metadata.ts'
import { validateStagedRuntimePackageVersion } from './runtime-package-publish-validation.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultPackDir = join(rootDir, 'build', 'npm-packages-runtime')
const textDecoder = new TextDecoder()
const distTag = process.env.NPM_DIST_TAG ?? 'latest'
const dryRun = parseBooleanEnv(process.env.DRY_RUN)
const allowNewNpmPackages = parseBooleanEnv(process.env.ALLOW_NEW_NPM_PACKAGES)
const skipUnprovisionedNpmPackages = parseBooleanEnv(process.env.SKIP_UNPROVISIONED_NPM_PACKAGES)
const targetVersion = process.env.TARGET_VERSION?.trim()

if (!targetVersion) {
  throw new Error('TARGET_VERSION is required')
}
parseStableSemver(targetVersion)

const allRuntimePackages = loadRuntimePackages(rootDir)
const repositoryVersion = assertAlignedVersions(allRuntimePackages)
if (repositoryVersion !== targetVersion) {
  throw new Error(
    `Repository runtime package manifests must be committed at ${targetVersion} before publishing (found ${repositoryVersion}). Run bun scripts/sync-runtime-package-versions.ts --version ${targetVersion} and commit the result.`,
  )
}
const runtimePackages = loadRuntimeNpmPackages(rootDir)

const currentPublishedRuntimeVersions = readPublishedRuntimePackageVersions(runtimePackages.map((runtimePackage) => runtimePackage.name))
assertKnownNpmPackagesBeforePublishing(currentPublishedRuntimeVersions)
const missingPackageNames = new Set(missingPublishedRuntimePackageNames(currentPublishedRuntimeVersions))
if (skipUnprovisionedNpmPackages && missingPackageNames.size > 0) {
  assertSkippedRuntimePackagesAreLeaves(runtimePackages, missingPackageNames)
}

const stageDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-package-set-'))
const packDir = resolve(process.env.PACK_DIR ?? defaultPackDir)
const internalPackageNames = new Set(allRuntimePackages.map((runtimePackage) => runtimePackage.name))

rmSync(packDir, { recursive: true, force: true })
mkdirSync(packDir, { recursive: true })

try {
  for (const runtimePackage of runtimePackages) {
    const sourceDir = join(rootDir, runtimePackage.dir)
    const stagedPackageDir = join(stageDir, basename(runtimePackage.dir))

    cpSync(sourceDir, stagedPackageDir, {
      recursive: true,
    })

    const packageJsonPath = join(stagedPackageDir, 'package.json')
    const manifest = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
    if (manifest.version !== targetVersion) {
      throw new Error(`Staged ${runtimePackage.name} package.json must already be ${targetVersion}, found ${String(manifest.version)}`)
    }
    rewriteInternalDependencyRanges(manifest, internalPackageNames, targetVersion)
    writeFileSync(packageJsonPath, `${JSON.stringify(manifest, null, 2)}\n`)
    syncStagedMcpServerMetadata(runtimePackage.name, stagedPackageDir, targetVersion)
    validateStagedRuntimePackageVersion(runtimePackage.name, stagedPackageDir, targetVersion)

    const packagePackDir = join(packDir, encodeURIComponent(runtimePackage.name))
    mkdirSync(packagePackDir, { recursive: true })

    runCommand('npm', ['pack', '--pack-destination', packagePackDir], {
      cwd: stagedPackageDir,
    })
  }

  const tarballsByPackage = indexTarballs(packDir)
  const results = []

  for (const runtimePackage of runtimePackages) {
    if (skipUnprovisionedNpmPackages && missingPackageNames.has(runtimePackage.name)) {
      results.push({
        package: runtimePackage.name,
        version: targetVersion,
        status: 'skipped',
        reason: 'npm package name is not provisioned yet',
      })
      continue
    }

    const tarballPath = tarballsByPackage.get(`${runtimePackage.name}@${targetVersion}`)
    if (!tarballPath) {
      throw new Error(`Packed tarball missing for ${runtimePackage.name}@${targetVersion}`)
    }

    if (isVersionPublished(runtimePackage.name, targetVersion)) {
      const distTags = getDistTags(runtimePackage.name)
      if (distTags[distTag] === targetVersion) {
        results.push({
          package: runtimePackage.name,
          version: targetVersion,
          status: 'skipped',
          reason: `version already published on ${distTag}`,
        })
        continue
      }

      if (dryRun) {
        results.push({
          package: runtimePackage.name,
          version: targetVersion,
          status: 'would-tag',
          tag: distTag,
        })
        continue
      }

      runCommand('npm', ['dist-tag', 'add', `${runtimePackage.name}@${targetVersion}`, distTag])
      results.push({
        package: runtimePackage.name,
        version: targetVersion,
        status: 'tagged',
        tag: distTag,
      })
      continue
    }

    const publishArgs = ['publish', tarballPath, '--tag', distTag, '--access', 'public', '--provenance']
    if (dryRun) {
      publishArgs.push('--dry-run')
    }
    runCommand('npm', publishArgs)

    results.push({
      package: runtimePackage.name,
      version: targetVersion,
      status: dryRun ? 'would-publish' : 'published',
      tag: distTag,
    })
  }

  console.log(
    JSON.stringify(
      {
        distTag,
        dryRun,
        targetVersion,
        results,
      },
      null,
      2,
    ),
  )
} finally {
  rmSync(stageDir, { recursive: true, force: true })
}

function rewriteInternalDependencyRanges(manifest: Record<string, unknown>, internalNames: Set<string>, nextVersion: string) {
  const dependencyFields = ['dependencies', 'optionalDependencies', 'peerDependencies', 'devDependencies'] as const

  for (const dependencyField of dependencyFields) {
    const dependencies = manifest[dependencyField]
    if (!dependencies || typeof dependencies !== 'object' || Array.isArray(dependencies)) {
      continue
    }
    for (const [dependencyName] of Object.entries(dependencies)) {
      if (internalNames.has(dependencyName)) {
        dependencies[dependencyName] = nextVersion
      }
    }
  }
}

function indexTarballs(targetDir: string) {
  const tarballs = listTarballsRecursive(targetDir)
  const entriesByPackage = new Map<string, string>()

  for (const tarballPath of tarballs) {
    const packedManifest = JSON.parse(runTextCommand('tar', ['-xOf', tarballPath, 'package/package.json']))
    if (typeof packedManifest.name !== 'string' || typeof packedManifest.version !== 'string') {
      throw new Error(`Packed manifest is missing name/version in ${tarballPath}`)
    }
    entriesByPackage.set(`${packedManifest.name}@${packedManifest.version}`, tarballPath)
  }

  return entriesByPackage
}

function listTarballsRecursive(targetDir: string): string[] {
  const tarballs: string[] = []
  for (const entry of readdirSync(targetDir)) {
    const entryPath = join(targetDir, entry)
    const entryStat = statSync(entryPath)
    if (entryStat.isDirectory()) {
      tarballs.push(...listTarballsRecursive(entryPath))
      continue
    }
    if (entry.endsWith('.tgz')) {
      tarballs.push(entryPath)
    }
  }
  return tarballs
}

function isVersionPublished(packageName: string, version: string) {
  const result = Bun.spawnSync(['npm', 'view', `${packageName}@${version}`, 'version', '--json'], {
    cwd: rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  return result.exitCode === 0
}

function readPublishedRuntimePackageVersions(packageNames: string[]): RuntimePackagePublishedVersion[] {
  return packageNames.map((packageName) => ({
    packageName,
    version: getPublishedVersion(packageName),
  }))
}

function getPublishedVersion(packageName: string): string | null {
  const result = Bun.spawnSync(['npm', 'view', packageName, 'dist-tags.latest', '--json'], {
    cwd: rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.exitCode !== 0) {
    return null
  }
  const output = textDecoder.decode(result.stdout).trim()
  if (!output || output === 'null') {
    return null
  }
  return JSON.parse(output)
}

function assertKnownNpmPackagesBeforePublishing(publishedVersions: readonly RuntimePackagePublishedVersion[]): void {
  const provisioningPlan = planRuntimePackagePublishProvisioning({
    publishedVersions,
    allowNewNpmPackages,
    skipUnprovisionedNpmPackages,
    dryRun,
  })
  if (provisioningPlan.publishAllowed) {
    return
  }
  throw new Error(
    [
      `Refusing to publish a partial runtime package set because npm does not know every runtime package yet (${formatRuntimePackagePublishedVersions(
        publishedVersions,
      )}).`,
      'Create the missing npm package(s) and configure trusted publishing first, explicitly set ALLOW_NEW_NPM_PACKAGES=true when using credentials that can create new public packages, or set SKIP_UNPROVISIONED_NPM_PACKAGES=true to publish only provisioned leaf packages.',
    ].join(' '),
  )
}

function assertSkippedRuntimePackagesAreLeaves(
  packageManifests: readonly RuntimePackageManifest[],
  missingNames: ReadonlySet<string>,
): void {
  for (const runtimePackage of packageManifests) {
    if (missingNames.has(runtimePackage.name)) {
      continue
    }
    const manifest = JSON.parse(readFileSync(join(rootDir, runtimePackage.dir, 'package.json'), 'utf8'))
    for (const [dependencyField, dependencyName] of readManifestDependencyNames(manifest)) {
      if (missingNames.has(dependencyName)) {
        throw new Error(
          `Cannot skip unprovisioned package ${dependencyName}: provisioned package ${runtimePackage.name} declares it in ${dependencyField}.`,
        )
      }
    }
  }
}

function readManifestDependencyNames(manifest: Record<string, unknown>): Array<readonly [string, string]> {
  const entries: Array<readonly [string, string]> = []
  for (const dependencyField of ['dependencies', 'optionalDependencies', 'peerDependencies'] as const) {
    const dependencies = manifest[dependencyField]
    if (!dependencies || typeof dependencies !== 'object' || Array.isArray(dependencies)) {
      continue
    }
    for (const dependencyName of Object.keys(dependencies)) {
      entries.push([dependencyField, dependencyName])
    }
  }
  return entries
}

function getDistTags(packageName: string): Record<string, string> {
  const result = Bun.spawnSync(['npm', 'view', packageName, 'dist-tags', '--json'], {
    cwd: rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.exitCode !== 0) {
    return {}
  }
  const output = textDecoder.decode(result.stdout).trim()
  if (output.length === 0) {
    return {}
  }
  return JSON.parse(output)
}

function runTextCommand(command: string, args: string[]): string {
  const result = Bun.spawnSync([command, ...args], {
    cwd: rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.exitCode !== 0) {
    const stderr = textDecoder.decode(result.stderr).trim()
    throw new Error(`Command failed: ${command} ${args.join(' ')}${stderr ? `\n${stderr}` : ''}`)
  }
  return textDecoder.decode(result.stdout)
}

function runCommand(command: string, args: string[], options?: { cwd?: string }) {
  const result = Bun.spawnSync([command, ...args], {
    cwd: options?.cwd ?? rootDir,
    env: process.env,
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
  })
  if (result.exitCode !== 0) {
    throw new Error(`Command failed: ${command} ${args.join(' ')}`)
  }
}
