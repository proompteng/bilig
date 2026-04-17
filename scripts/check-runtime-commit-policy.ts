#!/usr/bin/env bun

import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { isRuntimeAffectingPath, parseConventionalCommit } from './runtime-release.ts'

interface CommitPolicyFailure {
  kind: 'commit' | 'pull-request-title'
  sha?: string
  subject: string
  message: string
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

const inputEventName = readOptionalStringArg('event-name') ?? process.env.GITHUB_EVENT_NAME ?? ''
const inputEventPath = readOptionalStringArg('event-path') ?? process.env.GITHUB_EVENT_PATH ?? ''
const baseShaArg = readOptionalStringArg('base-sha')
const headShaArg = readOptionalStringArg('head-sha')
const prTitleArg = readOptionalStringArg('pr-title')

const githubEventPayload = inputEventPath ? readEventPayload(inputEventPath) : null
const baseSha = baseShaArg ?? readBaseSha(inputEventName, githubEventPayload) ?? readLocalBaseSha()
const headSha = headShaArg ?? readHeadSha(inputEventName, githubEventPayload) ?? 'HEAD'
const pullRequestTitle = prTitleArg ?? readPullRequestTitle(inputEventName, githubEventPayload)

const commitShas = listCommitShas(baseSha, headSha)
const failures: CommitPolicyFailure[] = []
const runtimeAffectingCommits = []

for (const sha of commitShas) {
  const subject = runCommand('git', ['show', '-s', '--format=%s', sha]).trim()
  const files = readCommitFiles(sha)
  const runtimeAffecting = !isMergeCommit(sha) && files.some((file) => isRuntimeAffectingPath(file))
  if (!runtimeAffecting) {
    continue
  }
  runtimeAffectingCommits.push({ sha, subject })
  if (!parseConventionalCommit({ subject, body: runCommand('git', ['show', '-s', '--format=%b', sha]).trim() })) {
    failures.push({
      kind: 'commit',
      sha,
      subject,
      message: `Runtime-affecting commit ${sha.slice(0, 8)} must use Conventional Commits`,
    })
  }
}

if (runtimeAffectingCommits.length > 0 && inputEventName === 'pull_request') {
  if (!pullRequestTitle) {
    failures.push({
      kind: 'pull-request-title',
      subject: '',
      message: 'Runtime-affecting pull requests must expose a title for squash-merge validation',
    })
  } else if (!parseConventionalCommit({ subject: pullRequestTitle, body: '' })) {
    failures.push({
      kind: 'pull-request-title',
      subject: pullRequestTitle,
      message: `Runtime-affecting pull request title must use Conventional Commits: ${pullRequestTitle}`,
    })
  }
}

if (failures.length > 0) {
  throw new Error(failures.map((failure) => failure.message).join('\n'))
}

console.log(
  JSON.stringify(
    {
      eventName: inputEventName,
      baseSha,
      headSha,
      runtimeAffectingCommitCount: runtimeAffectingCommits.length,
      checkedCommits: commitShas.length,
      status: 'ok',
    },
    null,
    2,
  ),
)

function readEventPayload(path: string): Record<string, unknown> | null {
  const resolvedPath = resolve(rootDir, path)
  try {
    return JSON.parse(readFileSync(resolvedPath, 'utf8'))
  } catch {
    return null
  }
}

function readBaseSha(triggerEventName: string, payload: Record<string, unknown> | null): string | null {
  if (triggerEventName === 'pull_request') {
    return readNestedString(payload, ['pull_request', 'base', 'sha'])
  }
  if (triggerEventName === 'push') {
    const before = readNestedString(payload, ['before'])
    return before && !/^0+$/.test(before) ? before : null
  }
  return null
}

function readHeadSha(triggerEventName: string, payload: Record<string, unknown> | null): string | null {
  if (triggerEventName === 'pull_request') {
    return readNestedString(payload, ['pull_request', 'head', 'sha'])
  }
  if (triggerEventName === 'push') {
    return readNestedString(payload, ['after']) ?? process.env.GITHUB_SHA ?? null
  }
  return process.env.GITHUB_SHA ?? null
}

function readPullRequestTitle(triggerEventName: string, payload: Record<string, unknown> | null): string | null {
  if (triggerEventName !== 'pull_request') {
    return null
  }
  return readNestedString(payload, ['pull_request', 'title'])
}

function readLocalBaseSha(): string | null {
  const result = Bun.spawnSync(['git', 'rev-parse', 'HEAD^'], {
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
  return output.length > 0 ? output : null
}

function listCommitShas(baseCommitSha: string | null, headCommitSha: string): string[] {
  const commandArgs = baseCommitSha
    ? ['rev-list', '--reverse', `${baseCommitSha}..${headCommitSha}`]
    : ['rev-list', '--reverse', headCommitSha]
  const output = runCommand('git', commandArgs)
  return output
    .split('\n')
    .map((entry) => entry.trim())
    .filter(Boolean)
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

function readNestedString(value: Record<string, unknown> | null, path: string[]): string | null {
  let current: unknown = value
  for (const segment of path) {
    if (!isRecord(current) || !(segment in current)) {
      return null
    }
    current = current[segment]
  }
  return typeof current === 'string' && current.length > 0 ? current : null
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

function readOptionalStringArg(name: string): string | null {
  const value = cliArgs.get(name)
  if (typeof value !== 'string') {
    return null
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object'
}
