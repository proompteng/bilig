#!/usr/bin/env bun

import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync, cpSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join, resolve } from 'node:path'

import { assertAlignedVersions, loadRuntimePackages, parseBooleanEnv } from './runtime-package-set.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultPackDir = join(rootDir, 'build', 'npm-packages-runtime')
const textDecoder = new TextDecoder()
const distTag = process.env.NPM_DIST_TAG ?? 'latest'
const dryRun = parseBooleanEnv(process.env.DRY_RUN)
const targetVersion = process.env.TARGET_VERSION?.trim()

if (!targetVersion) {
  throw new Error('TARGET_VERSION is required')
}

const runtimePackages = loadRuntimePackages(rootDir)
assertAlignedVersions(runtimePackages)

const stageDir = mkdtempSync(join(tmpdir(), 'bilig-runtime-package-set-'))
const packDir = resolve(process.env.PACK_DIR ?? defaultPackDir)
const internalPackageNames = new Set(runtimePackages.map((runtimePackage) => runtimePackage.name))

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
    manifest.version = targetVersion
    rewriteInternalDependencyRanges(manifest, internalPackageNames, targetVersion)
    writeFileSync(packageJsonPath, `${JSON.stringify(manifest, null, 2)}\n`)

    runCommand('npm', ['pack', '--pack-destination', packDir], {
      cwd: stagedPackageDir,
    })
  }

  const tarballsByPackage = indexTarballs(packDir)
  const results = []

  for (const runtimePackage of runtimePackages) {
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
  const tarballs = readdirSync(targetDir).filter((entry) => entry.endsWith('.tgz'))
  const entriesByPackage = new Map()

  for (const tarballName of tarballs) {
    const tarballPath = join(targetDir, tarballName)
    const packedManifest = JSON.parse(runTextCommand('tar', ['-xOf', tarballPath, 'package/package.json']))
    if (typeof packedManifest.name !== 'string' || typeof packedManifest.version !== 'string') {
      throw new Error(`Packed manifest is missing name/version in ${tarballPath}`)
    }
    entriesByPackage.set(`${packedManifest.name}@${packedManifest.version}`, tarballPath)
  }

  return entriesByPackage
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
