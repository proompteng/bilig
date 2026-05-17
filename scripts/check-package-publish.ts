#!/usr/bin/env bun

import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { isAbsolute, join, resolve } from 'node:path'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const packagesDir = join(rootDir, 'packages')
const packDir = join(rootDir, 'build', 'npm-packages')
const textDecoder = new TextDecoder()
const cliArgs = process.argv.slice(2)
const requireAlignedVersions = cliArgs.includes('--require-aligned')
const packageArgs = cliArgs.filter((arg) => arg !== '--require-aligned')

const packageDirs =
  packageArgs.length > 0
    ? packageArgs.map(resolvePackageDir)
    : readdirSync(packagesDir)
        .map((name) => join(packagesDir, name))
        .filter((dir) => existsSync(join(dir, 'package.json')))

rmSync(packDir, { recursive: true, force: true })
mkdirSync(packDir, { recursive: true })

const failures = []

if (requireAlignedVersions) {
  validateAlignedVersions(packageDirs, failures)
}

for (const packageDir of packageDirs) {
  const packageJsonPath = join(packageDir, 'package.json')
  const manifest = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  const packageLabel = manifest.name ?? packageDir

  validateManifestShape(packageLabel, manifest, failures)

  const tarballName = runTextCommand('pnpm', ['pack', '--pack-destination', packDir], {
    cwd: packageDir,
  })
    .trim()
    .split('\n')
    .pop()

  if (!tarballName) {
    failures.push(`${packageLabel}: pnpm pack did not return a tarball name`)
    continue
  }

  const tarballPath = isAbsolute(tarballName) ? tarballName : join(packDir, tarballName)
  const tarEntries = runTextCommand('tar', ['-tf', tarballPath]).split('\n').filter(Boolean)

  validateTarballContents(packageLabel, manifest, tarEntries, failures)

  const packedManifest = JSON.parse(runTextCommand('tar', ['-xOf', tarballPath, 'package/package.json']))
  validatePackedManifest(packageLabel, packedManifest, tarballPath, failures)
}

if (failures.length > 0) {
  throw new Error(`npm publish readiness check failed:\n- ${failures.join('\n- ')}`)
}

console.log(
  JSON.stringify(
    {
      packages: packageDirs.length,
      packDir,
    },
    null,
    2,
  ),
)

function resolvePackageDir(packageArg) {
  const resolvedArg = resolve(rootDir, packageArg)
  if (!existsSync(join(resolvedArg, 'package.json'))) {
    throw new Error(`Package directory does not contain package.json: ${packageArg}`)
  }
  return resolvedArg
}

function runTextCommand(command, args, options = {}) {
  const result = Bun.spawnSync([command, ...args], {
    cwd: options.cwd,
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

function validateManifestShape(packageLabel, manifest, failureMessages) {
  const requiredFields =
    packageLabel === '@bilig/create-workpaper'
      ? ['name', 'version', 'description', 'license', 'repository', 'homepage', 'bugs', 'bin', 'files']
      : ['name', 'version', 'description', 'license', 'repository', 'homepage', 'bugs', 'main', 'types', 'exports', 'files']
  for (const field of requiredFields) {
    if (!(field in manifest)) {
      failureMessages.push(`${packageLabel}: missing required manifest field "${field}"`)
    }
  }

  if (manifest.publishConfig?.access !== 'public') {
    failureMessages.push(`${packageLabel}: publishConfig.access must be "public"`)
  }

  if (!Array.isArray(manifest.files) || manifest.files.length === 0) {
    failureMessages.push(`${packageLabel}: files list must be present and non-empty`)
  }

  if ((packageLabel === '@bilig/grid' || packageLabel === '@bilig/renderer') && !manifest.peerDependencies?.react) {
    failureMessages.push(`${packageLabel}: react must be declared as a peer dependency`)
  }
}

function validateTarballContents(packageLabel, manifest, tarEntries, failureMessages) {
  const requiredEntries = new Set(['package/package.json', 'package/README.md', 'package/LICENSE'])

  if (typeof manifest.main === 'string') {
    requiredEntries.add(`package/${stripDotSlash(manifest.main)}`)
  }
  if (typeof manifest.types === 'string') {
    requiredEntries.add(`package/${stripDotSlash(manifest.types)}`)
  }

  collectExportTargets(manifest.exports).forEach((target) => requiredEntries.add(`package/${stripDotSlash(target)}`))
  if (packageLabel === '@bilig/wasm-kernel') {
    requiredEntries.add('package/build/release.wasm')
  }
  if (typeof manifest.mcpName === 'string') {
    requiredEntries.add('package/server.json')
  }
  if (packageLabel === '@bilig/headless') {
    requiredEntries.add('package/AGENTS.md')
  }
  collectBinTargets(manifest.bin).forEach((target) => requiredEntries.add(`package/${stripDotSlash(target)}`))

  for (const requiredEntry of requiredEntries) {
    if (!tarEntries.includes(requiredEntry)) {
      failureMessages.push(`${packageLabel}: tarball is missing ${requiredEntry}`)
    }
  }

  for (const entry of tarEntries) {
    if (entry.includes('__tests__')) {
      failureMessages.push(`${packageLabel}: tarball must not contain test artifacts (${entry})`)
    }
    if (entry.endsWith('.tsbuildinfo')) {
      failureMessages.push(`${packageLabel}: tarball must not contain tsbuildinfo (${entry})`)
    }
    if (entry.startsWith('package/src/')) {
      failureMessages.push(`${packageLabel}: tarball must not contain source files (${entry})`)
    }
  }
}

function validatePackedManifest(packageLabel, packedManifest, tarballPath, failureMessages) {
  const serialized = JSON.stringify(packedManifest)
  if (serialized.includes('workspace:*')) {
    failureMessages.push(`${packageLabel}: packed manifest still contains workspace:* dependency ranges`)
  }

  if (typeof packedManifest.mcpName === 'string') {
    validatePackedMcpMetadata(packageLabel, packedManifest, tarballPath, failureMessages)
  }
}

function validatePackedMcpMetadata(packageLabel, packedManifest, tarballPath, failureMessages) {
  const serverJson = JSON.parse(runTextCommand('tar', ['-xOf', tarballPath, 'package/server.json']))
  if (serverJson.name !== packedManifest.mcpName) {
    failureMessages.push(`${packageLabel}: package.json mcpName must match server.json name`)
  }
  if (typeof serverJson.description !== 'string' || serverJson.description.length > 100) {
    failureMessages.push(`${packageLabel}: server.json description must be a string no longer than 100 characters`)
  }
  const npmPackage = Array.isArray(serverJson.packages)
    ? serverJson.packages.find((entry) => entry?.registryType === 'npm' && entry?.identifier === packedManifest.name)
    : undefined
  if (!npmPackage) {
    failureMessages.push(`${packageLabel}: server.json must include an npm package entry for ${packedManifest.name}`)
  } else if (npmPackage.version !== packedManifest.version) {
    failureMessages.push(`${packageLabel}: server.json package version must match package.json version`)
  }
}

function validateAlignedVersions(packageDirectories, failureMessages) {
  const versions = new Map()
  for (const packageDir of packageDirectories) {
    const manifest = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'))
    const packageLabel = manifest.name ?? packageDir
    const version = manifest.version
    if (typeof version !== 'string' || version.length === 0) {
      failureMessages.push(`${packageLabel}: version must be a non-empty string`)
      continue
    }
    const packagesForVersion = versions.get(version) ?? []
    packagesForVersion.push(packageLabel)
    versions.set(version, packagesForVersion)
  }

  if (versions.size > 1) {
    const summary = [...versions.entries()].map(([version, packageLabels]) => `${version}: ${packageLabels.join(', ')}`).join(' | ')
    failureMessages.push(`release package versions must stay aligned (${summary})`)
  }
}

function collectExportTargets(exportsField) {
  const targets = new Set()
  visitExports(exportsField, targets)
  return [...targets]
}

function collectBinTargets(binField) {
  if (typeof binField === 'string') {
    return [binField]
  }
  if (!binField || typeof binField !== 'object') {
    return []
  }
  return Object.values(binField).filter((value) => typeof value === 'string')
}

function visitExports(node, targets) {
  if (typeof node === 'string') {
    targets.add(node)
    return
  }
  if (!node || typeof node !== 'object') {
    return
  }
  for (const value of Object.values(node)) {
    visitExports(value, targets)
  }
}

function stripDotSlash(value) {
  return value.startsWith('./') ? value.slice(2) : value
}
