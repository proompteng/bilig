#!/usr/bin/env bun

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'

import { assertAlignedVersions, compareStableSemver, loadRuntimePackages, parseStableSemver } from './runtime-package-set.ts'
import {
  bumpVersion,
  extractVersionFromRuntimeTag,
  isRuntimeAffectingPath,
  maxReleaseType,
  parseConventionalCommit,
  releaseTypeForConventionalCommit,
  RUNTIME_RELEASE_TAG_PREFIX,
  summarizeReleaseNotes,
  type ReleaseType,
  type RuntimeReleaseCommit,
} from './runtime-release.ts'

interface RuntimeReleasePlan {
  manifestVersion: string
  publishedVersion: string | null
  lastTag: string | null
  releaseNeeded: boolean
  bootstrapRequired: boolean
  manualOverride: boolean
  releaseType: ReleaseType
  reason: 'bootstrap-required' | 'manual-override' | 'no-runtime-commits' | 'no-release-commits' | 'release-required'
  targetVersion: string | null
  tagName: string | null
  notesMarkdown: string | null
  commits: RuntimeReleaseCommit[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const cliArgs = new Map<string, string | true>()
for (let index = 0; index < process.argv.length; index += 1) {
  const value = process.argv[index]
  if (!value || value === '--' || !value.startsWith('--')) {
    continue
  }
  const key = value.slice(2)
  const nextValue = process.argv[index + 1]
  if (!nextValue || nextValue.startsWith('--')) {
    cliArgs.set(key, true)
    continue
  }
  cliArgs.set(key, nextValue)
  index += 1
}

const requestedReleaseAs = readOptionalStringArg('release-as')
const requestedNotesFile = readOptionalStringArg('notes-file')
const requestedGithubOutput = readOptionalStringArg('github-output')
const allowManualBootstrap = cliArgs.has('allow-untagged-baseline')

const runtimePackages = loadRuntimePackages(rootDir)
const runtimeManifestVersion = assertAlignedVersions(runtimePackages)
const latestPublishedVersion = getAlignedPublishedVersion(runtimePackages.map((runtimePackage) => runtimePackage.name))
const latestReachableTag = getLatestReachableRuntimeTag()

const runtimeReleasePlan = buildRuntimeReleasePlan({
  manifestVersion: runtimeManifestVersion,
  publishedVersion: latestPublishedVersion,
  lastTag: latestReachableTag,
  releaseAs: requestedReleaseAs,
  allowUntaggedBaseline: allowManualBootstrap,
})

if (requestedNotesFile && runtimeReleasePlan.notesMarkdown) {
  ensureParentDir(requestedNotesFile)
  writeFileSync(requestedNotesFile, runtimeReleasePlan.notesMarkdown)
}

if (requestedGithubOutput) {
  ensureParentDir(requestedGithubOutput)
  writeFileSync(requestedGithubOutput, formatGithubOutput(runtimeReleasePlan, requestedNotesFile), { flag: 'a' })
}

console.log(JSON.stringify(runtimeReleasePlan, null, 2))

function buildRuntimeReleasePlan(input: {
  manifestVersion: string
  publishedVersion: string | null
  lastTag: string | null
  releaseAs: string | null
  allowUntaggedBaseline: boolean
}): RuntimeReleasePlan {
  const { manifestVersion, publishedVersion, lastTag, releaseAs, allowUntaggedBaseline } = input

  if (releaseAs) {
    parseStableSemver(releaseAs)
  }

  if (!lastTag && !releaseAs && !allowUntaggedBaseline) {
    return {
      manifestVersion,
      publishedVersion,
      lastTag: null,
      releaseNeeded: false,
      bootstrapRequired: true,
      manualOverride: false,
      releaseType: 'none',
      reason: 'bootstrap-required',
      targetVersion: null,
      tagName: null,
      notesMarkdown: null,
      commits: [],
    }
  }

  if (!lastTag && allowUntaggedBaseline) {
    if (!releaseAs) {
      return {
        manifestVersion,
        publishedVersion,
        lastTag: null,
        releaseNeeded: false,
        bootstrapRequired: true,
        manualOverride: false,
        releaseType: 'none',
        reason: 'bootstrap-required',
        targetVersion: null,
        tagName: null,
        notesMarkdown: null,
        commits: [],
      }
    }

    const baselineVersion = publishedVersion ?? manifestVersion
    if (publishedVersion !== null && compareStableSemver(releaseAs, baselineVersion) <= 0) {
      throw new Error(`Manual bootstrap version ${releaseAs} must be greater than published baseline ${baselineVersion}`)
    }

    const bootstrapReleaseType = inferReleaseTypeFromVersionChange(baselineVersion, releaseAs)
    return {
      manifestVersion,
      publishedVersion,
      lastTag: null,
      releaseNeeded: true,
      bootstrapRequired: false,
      manualOverride: true,
      releaseType: bootstrapReleaseType,
      reason: 'manual-override',
      targetVersion: releaseAs,
      tagName: `${RUNTIME_RELEASE_TAG_PREFIX}${releaseAs}`,
      notesMarkdown: summarizeReleaseNotes({
        targetVersion: releaseAs,
        lastTag: null,
        commits: [],
        releaseType: bootstrapReleaseType,
        manualOverride: true,
      }),
      commits: [],
    }
  }

  const commits = collectRuntimeReleaseCommits(lastTag)
  const runtimeCommits = commits.filter((commit) => commit.runtimeAffecting)
  const strongestReleaseType = runtimeCommits.reduce<ReleaseType>((highest, commit) => maxReleaseType(highest, commit.releaseType), 'none')

  const baselineVersion = extractVersionFromRuntimeTag(lastTag ?? '') ?? publishedVersion ?? manifestVersion

  if (!runtimeCommits.length && !releaseAs) {
    return {
      manifestVersion,
      publishedVersion,
      lastTag,
      releaseNeeded: false,
      bootstrapRequired: false,
      manualOverride: false,
      releaseType: 'none',
      reason: 'no-runtime-commits',
      targetVersion: null,
      tagName: null,
      notesMarkdown: null,
      commits,
    }
  }

  if (strongestReleaseType === 'none' && !releaseAs) {
    return {
      manifestVersion,
      publishedVersion,
      lastTag,
      releaseNeeded: false,
      bootstrapRequired: false,
      manualOverride: false,
      releaseType: 'none',
      reason: 'no-release-commits',
      targetVersion: null,
      tagName: null,
      notesMarkdown: null,
      commits,
    }
  }

  const targetVersion =
    releaseAs ??
    (strongestReleaseType === 'none'
      ? manifestVersion
      : lastTag || publishedVersion
        ? bumpVersion(baselineVersion, strongestReleaseType)
        : manifestVersion)

  if ((lastTag !== null || publishedVersion !== null) && compareStableSemver(targetVersion, baselineVersion) <= 0) {
    throw new Error(`Target runtime release version ${targetVersion} must be greater than baseline ${baselineVersion}`)
  }

  const effectiveReleaseType: Exclude<ReleaseType, 'none'> =
    strongestReleaseType === 'none'
      ? releaseAs
        ? inferReleaseTypeFromVersionChange(baselineVersion, targetVersion)
        : 'patch'
      : strongestReleaseType

  return {
    manifestVersion,
    publishedVersion,
    lastTag,
    releaseNeeded: true,
    bootstrapRequired: false,
    manualOverride: Boolean(releaseAs),
    releaseType: effectiveReleaseType,
    reason: releaseAs ? 'manual-override' : 'release-required',
    targetVersion,
    tagName: `${RUNTIME_RELEASE_TAG_PREFIX}${targetVersion}`,
    notesMarkdown: summarizeReleaseNotes({
      targetVersion,
      lastTag,
      commits: runtimeCommits,
      releaseType: effectiveReleaseType,
      manualOverride: Boolean(releaseAs),
    }),
    commits,
  }
}

function collectRuntimeReleaseCommits(baseTag: string | null): RuntimeReleaseCommit[] {
  const shas = listCommitShas(baseTag)
  return shas.map((sha) => {
    const { subject, body } = readCommitMessage(sha)
    const files = readCommitFiles(sha)
    const mergeCommit = isMergeCommit(sha)
    const runtimeAffecting = !mergeCommit && files.some((file) => isRuntimeAffectingPath(file))
    const conventional = parseConventionalCommit({ subject, body })
    if (runtimeAffecting && conventional === null) {
      throw new Error(`Runtime-affecting commit ${sha.slice(0, 8)} is not a Conventional Commit: ${subject}`)
    }
    const releaseType = conventional ? releaseTypeForConventionalCommit(conventional) : 'none'
    return {
      sha,
      shortSha: sha.slice(0, 8),
      subject,
      body,
      files,
      runtimeAffecting,
      conventional,
      releaseType,
    } satisfies RuntimeReleaseCommit
  })
}

function listCommitShas(baseTag: string | null): string[] {
  const rangeArgs = baseTag ? [`${baseTag}..HEAD`] : ['HEAD']
  const result = runCommand('git', ['rev-list', '--reverse', ...rangeArgs])
  return result
    .split('\n')
    .map((entry) => entry.trim())
    .filter(Boolean)
}

function readCommitMessage(sha: string): { subject: string; body: string } {
  const output = runCommand('git', ['show', '-s', '--format=%s%n%b', sha])
  const [subject = '', ...bodyLines] = output.split('\n')
  return {
    subject: subject.trim(),
    body: bodyLines.join('\n').trim(),
  }
}

function readCommitFiles(sha: string): string[] {
  const output = runCommand('git', ['diff-tree', '--no-commit-id', '--name-only', '-r', '-m', sha])
  return [
    ...new Set(
      output
        .split('\n')
        .map((entry) => entry.trim())
        .filter(Boolean),
    ),
  ]
}

function isMergeCommit(sha: string): boolean {
  const output = runCommand('git', ['rev-list', '--parents', '-n', '1', sha])
  return output.split(' ').filter(Boolean).length > 2
}

function getLatestReachableRuntimeTag(): string | null {
  const output = runCommand('git', ['tag', '--merged', 'HEAD', '--list', `${RUNTIME_RELEASE_TAG_PREFIX}*`, '--sort=-version:refname'])
  return (
    output
      .split('\n')
      .map((entry) => entry.trim())
      .find((entry) => entry.length > 0) ?? null
  )
}

function getAlignedPublishedVersion(packageNames: string[]): string | null {
  const publishedVersions = packageNames.map((packageName) => ({
    packageName,
    version: getPublishedVersion(packageName),
  }))
  const versionSet = new Set(
    publishedVersions.map((entry) => entry.version).filter((entry): entry is string => typeof entry === 'string' && entry.length > 0),
  )
  if (versionSet.size > 1) {
    throw new Error(
      `Published runtime package versions are not aligned (${publishedVersions
        .map((entry) => `${entry.packageName}@${entry.version ?? 'unpublished'}`)
        .join(', ')})`,
    )
  }
  const [firstVersion] = versionSet
  return firstVersion ?? null
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
  const output = new TextDecoder().decode(result.stdout).trim()
  if (!output || output === 'null') {
    return null
  }
  return JSON.parse(output)
}

function inferReleaseTypeFromVersionChange(previousVersion: string, nextVersion: string): Exclude<ReleaseType, 'none'> {
  const previous = parseStableSemver(previousVersion)
  const next = parseStableSemver(nextVersion)
  if (next.major > previous.major) {
    return 'major'
  }
  if (next.minor > previous.minor) {
    return 'minor'
  }
  return 'patch'
}

function runCommand(command: string, args: string[]): string {
  const result = Bun.spawnSync([command, ...args], {
    cwd: rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.exitCode !== 0) {
    const stderr = new TextDecoder().decode(result.stderr).trim()
    throw new Error(`Command failed: ${command} ${args.join(' ')}${stderr ? `\n${stderr}` : ''}`)
  }
  return new TextDecoder().decode(result.stdout).trim()
}

function formatGithubOutput(plan: RuntimeReleasePlan, notesFile: string | null): string {
  const lines = [
    `release_needed=${String(plan.releaseNeeded)}`,
    `bootstrap_required=${String(plan.bootstrapRequired)}`,
    `manual_override=${String(plan.manualOverride)}`,
    `release_type=${plan.releaseType}`,
    `reason=${plan.reason}`,
    `manifest_version=${plan.manifestVersion}`,
    `published_version=${plan.publishedVersion ?? ''}`,
    `last_tag=${plan.lastTag ?? ''}`,
    `target_version=${plan.targetVersion ?? ''}`,
    `tag_name=${plan.tagName ?? ''}`,
    `notes_file=${notesFile ?? ''}`,
  ]
  return `${lines.join('\n')}\n`
}

function ensureParentDir(path: string): void {
  mkdirSync(dirname(resolve(rootDir, path)), { recursive: true })
}

function readOptionalStringArg(name: string): string | null {
  const value = cliArgs.get(name)
  if (typeof value !== 'string') {
    return null
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : null
}
