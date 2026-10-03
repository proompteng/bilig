import { spawnSync } from 'node:child_process'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const allowReleasePending = process.argv.includes('--allow-release-pending')
const publicPackageSpec = '@bilig/workpaper@latest'

interface PackageManifest {
  readonly version: string
}

interface CommandResult {
  readonly status: number | null
  readonly stdout: string
  readonly stderr: string
}

interface RuleFailure {
  readonly target: string
  readonly result: CommandResult
}

function runCommand(command: string, args: readonly string[]): CommandResult {
  const result = spawnSync(command, args, {
    cwd: repoRoot,
    encoding: 'utf8',
  })

  return {
    status: result.status,
    stdout: result.stdout,
    stderr: result.stderr,
  }
}

function requireSuccessfulCommand(command: string, args: readonly string[]): string {
  const result = runCommand(command, args)
  if (result.status === 0) {
    return result.stdout.trim()
  }

  throw new Error(
    [`${command} ${args.join(' ')} failed with status ${result.status ?? 'unknown'}`, result.stdout.trim(), result.stderr.trim()]
      .filter((line) => line.length > 0)
      .join('\n'),
  )
}

function parsePackageManifest(content: string): PackageManifest {
  const parsed: unknown = JSON.parse(content)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('packages/workpaper/package.json must be an object')
  }

  const version = Reflect.get(parsed, 'version')
  if (typeof version !== 'string') {
    throw new Error('packages/workpaper/package.json must define a string version')
  }

  return { version }
}

function extractAdvertisedRuleTargets(agentStart: string): readonly string[] {
  const targets = new Set<string>()
  for (const match of agentStart.matchAll(/\bbilig-agent-start --rules ([a-z0-9-]+)/g)) {
    const target = match[1]
    if (target !== undefined) {
      targets.add(target)
    }
  }

  return [...targets].toSorted()
}

const docsAgentStart = await readFile(join(repoRoot, '.cache', 'docs-source', 'agent-start.txt'), 'utf8')
const localManifest = parsePackageManifest(await readFile(join(repoRoot, 'packages', 'workpaper', 'package.json'), 'utf8'))
const npmLatestVersion = requireSuccessfulCommand('npm', ['view', '@bilig/workpaper', 'version'])
const advertisedTargets = extractAdvertisedRuleTargets(docsAgentStart)

if (advertisedTargets.length === 0) {
  throw new Error('docs/agent-start.txt does not advertise any bilig-agent-start --rules targets')
}

const failures: RuleFailure[] = []

for (const target of advertisedTargets) {
  const result = runCommand('npm', ['exec', '--yes', '--package', publicPackageSpec, '--', 'bilig-agent-start', '--rules', target])

  if (result.status !== 0) {
    failures.push({ target, result })
  }
}

if (failures.length > 0) {
  const releasePending = npmLatestVersion !== localManifest.version
  const details = failures
    .map((failure) =>
      [
        `target=${failure.target}`,
        `status=${failure.result.status ?? 'unknown'}`,
        failure.result.stdout.trim(),
        failure.result.stderr.trim(),
      ]
        .filter((line) => line.length > 0)
        .join('\n'),
    )
    .join('\n\n')

  if (allowReleasePending && releasePending) {
    console.warn(
      [
        `Public agent-start rules are release-pending: npm latest is ${npmLatestVersion}, local package is ${localManifest.version}.`,
        details,
      ].join('\n\n'),
    )
    process.exit(0)
  }

  throw new Error(
    [
      `Public ${publicPackageSpec} does not support every docs/agent-start.txt --rules target.`,
      `npm latest: ${npmLatestVersion}`,
      `local package: ${localManifest.version}`,
      details,
    ].join('\n\n'),
  )
}

console.log(
  JSON.stringify(
    {
      schemaVersion: 'bilig-public-agent-start-rules.v1',
      package: publicPackageSpec,
      localVersion: localManifest.version,
      npmLatestVersion,
      advertisedTargets,
      verified: true,
    },
    null,
    2,
  ),
)
