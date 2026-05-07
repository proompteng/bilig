#!/usr/bin/env bun

import { appendFileSync } from 'node:fs'
import { resolve } from 'node:path'

import {
  formatRuntimePackagePublishedVersions,
  loadRuntimeNpmPackages,
  parseBooleanEnv,
  planRuntimePackagePublishProvisioning,
  type RuntimePackagePublishedVersion,
} from './runtime-package-set.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const textDecoder = new TextDecoder()
const allowNewNpmPackages = parseBooleanEnv(process.env.ALLOW_NEW_NPM_PACKAGES)
const dryRun = parseBooleanEnv(process.env.DRY_RUN)

const runtimePackages = loadRuntimeNpmPackages(rootDir)
const publishedVersions = readPublishedRuntimePackageVersions(runtimePackages.map((runtimePackage) => runtimePackage.name))
const provisioningPlan = planRuntimePackagePublishProvisioning({
  publishedVersions,
  allowNewNpmPackages,
  dryRun,
})

const provisioningOutput = {
  publishAllowed: provisioningPlan.publishAllowed,
  reason: provisioningPlan.reason,
  missingPackageNames: provisioningPlan.missingPackageNames,
  publishedVersions: formatRuntimePackagePublishedVersions(publishedVersions),
}

console.log(JSON.stringify(provisioningOutput, null, 2))

if (process.env.GITHUB_OUTPUT) {
  writeGithubOutput(process.env.GITHUB_OUTPUT, {
    publish_allowed: String(provisioningPlan.publishAllowed),
    reason: provisioningPlan.reason,
    missing_packages: provisioningPlan.missingPackageNames.join(','),
    published_versions: formatRuntimePackagePublishedVersions(publishedVersions),
  })
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

function writeGithubOutput(path: string, values: Record<string, string>): void {
  for (const [key, value] of Object.entries(values)) {
    appendFileSync(path, `${key}=${escapeGithubOutputValue(value)}\n`)
  }
}

function escapeGithubOutputValue(value: string): string {
  return value.replaceAll('%', '%25').replaceAll('\r', '%0D').replaceAll('\n', '%0A')
}
