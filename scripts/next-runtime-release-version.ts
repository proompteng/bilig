#!/usr/bin/env bun

const rootDir = new URL('..', import.meta.url).pathname
const passthroughArgs = process.argv.slice(2)

const result = Bun.spawnSync(['bun', 'scripts/plan-runtime-release.ts', ...passthroughArgs], {
  cwd: rootDir,
  env: process.env,
  stdin: 'ignore',
  stdout: 'pipe',
  stderr: 'pipe',
})

if (result.exitCode !== 0) {
  const stderr = new TextDecoder().decode(result.stderr).trim()
  throw new Error(stderr || 'Unable to derive the next runtime release version')
}

const rawOutput = new TextDecoder().decode(result.stdout).trim()
const parsed = JSON.parse(rawOutput)
if (!isRuntimeReleaseSummary(parsed)) {
  throw new Error('Runtime release planner returned an invalid JSON payload')
}
const plan = parsed

console.log(
  JSON.stringify(
    {
      manifestVersion: plan.manifestVersion,
      publishedVersion: plan.publishedVersion,
      lastTag: plan.lastTag,
      releaseNeeded: plan.releaseNeeded,
      bootstrapRequired: plan.bootstrapRequired,
      reason: plan.reason,
      targetVersion: plan.targetVersion,
    },
    null,
    2,
  ),
)

function isRuntimeReleaseSummary(value: unknown): value is {
  manifestVersion: string
  publishedVersion: string | null
  targetVersion: string | null
  releaseNeeded: boolean
  bootstrapRequired: boolean
  lastTag: string | null
  reason: string
} {
  if (!isRecord(value)) {
    return false
  }
  return (
    typeof value.manifestVersion === 'string' &&
    (typeof value.publishedVersion === 'string' || value.publishedVersion === null) &&
    (typeof value.targetVersion === 'string' || value.targetVersion === null) &&
    typeof value.releaseNeeded === 'boolean' &&
    typeof value.bootstrapRequired === 'boolean' &&
    (typeof value.lastTag === 'string' || value.lastTag === null) &&
    typeof value.reason === 'string'
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object'
}
