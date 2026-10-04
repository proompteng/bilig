import { readFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parse } from 'yaml'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readWorkflow(path: string): Record<string, unknown> {
  const parsed: unknown = parse(readFileSync(resolve(repoRoot, path), 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`${path} must parse to a YAML object`)
  }
  return parsed
}

function asRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`${label} must be a YAML object`)
  }
  return value
}

function workflowTriggers(workflow: Record<string, unknown>): Record<string, unknown> {
  return asRecord(workflow['on'], 'workflow.on')
}

function workflowJobs(workflow: Record<string, unknown>): Record<string, unknown> {
  return asRecord(workflow['jobs'], 'workflow.jobs')
}

function workflowConcurrency(workflow: Record<string, unknown>): Record<string, unknown> {
  return asRecord(workflow['concurrency'], 'workflow.concurrency')
}

describe('forgejo workflows', () => {
  it('publishes all image tags with the ARM variant required by Kargo', () => {
    const jobs = workflowJobs(readWorkflow('.forgejo/workflows/release-images.yml'))
    const publish = asRecord(jobs['publish-manifests'], 'publish-manifests')
    const steps = publish['steps']
    if (!Array.isArray(steps)) throw new Error('publish-manifests.steps must be an array')
    const step = asRecord(
      steps.find((entry: unknown) => isRecord(entry) && entry['name'] === 'Publish multi-arch image tags'),
      'Publish multi-arch image tags',
    )
    const run = step['run']
    if (typeof run !== 'string') throw new Error('Publish multi-arch image tags.run must be a script')
    const sha = 'a'.repeat(40)
    const script = run
      .replaceAll('${{ steps.meta.outputs.sha }}', sha)
      .replaceAll('${{ steps.meta.outputs.short_sha }}', sha.slice(0, 7))
      .replaceAll('${{ steps.head.outputs.publish_latest }}', 'true')
    const output = execFileSync('bash', ['-euc', `docker() { printf '%s\\n' "$*"; }\n${script}`], {
      env: { ...process.env, REGISTRY: 'registry.example.test', IMAGE_NAMESPACE: 'lab' },
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    })
    const image = 'registry.example.test/lab/bilig-app'
    expect(output.split('\n').filter((line) => line.startsWith('manifest annotate') && line.includes('--arch arm64'))).toEqual(
      [sha, sha.slice(0, 7), 'latest'].map(
        (tag) => `manifest annotate ${image}:${tag} ${image}:${sha}-arm64 --os linux --arch arm64 --variant v8`,
      ),
    )
  })

  it('keeps manual deep correctness out of push-triggered release images', () => {
    const releaseImages = readWorkflow('.forgejo/workflows/release-images.yml')
    const manualDeep = readWorkflow('.forgejo/workflows/forgejo-manual-deep-correctness.yml')

    expect(Object.keys(workflowTriggers(releaseImages))).toEqual(['push', 'workflow_dispatch'])
    expect(Object.keys(workflowJobs(releaseImages))).toEqual(['build-app-amd64', 'build-app-arm64', 'publish-manifests'])
    expect(readFileSync(resolve(repoRoot, '.forgejo/workflows/release-images.yml'), 'utf8')).not.toContain('manual-deep-correctness')

    expect(Object.keys(workflowTriggers(manualDeep))).toEqual(['workflow_dispatch'])
    expect(Object.keys(workflowJobs(manualDeep))).toEqual(['correctness-deep'])
    expect(readFileSync(resolve(repoRoot, '.forgejo/workflows/forgejo-manual-deep-correctness.yml'), 'utf8')).toContain('pnpm run ci:full')
  })

  it('pins Forgejo Node and Bun surfaces used by repository CI', () => {
    const workflow = readFileSync(resolve(repoRoot, '.forgejo/workflows/forgejo-ci.yml'), 'utf8')
    const manualDeep = readFileSync(resolve(repoRoot, '.forgejo/workflows/forgejo-manual-deep-correctness.yml'), 'utf8')

    expect(workflow).toContain('image: mcr.microsoft.com/playwright:v1.58.2-noble')
    expect(workflow.match(/npm install -g bun@1\.3\.10/gu)).toHaveLength(2)
    expect(workflow.match(/test "\$\(bun --version\)" = "1\.3\.10"/gu)).toHaveLength(2)
    expect(manualDeep).toContain('npm install -g bun@1.3.10')
    expect(manualDeep).toContain('test "$(bun --version)" = "1.3.10"')
  })

  it('cancels stale per-ref Forgejo runs when a newer commit is pushed', () => {
    for (const path of ['.forgejo/workflows/forgejo-ci.yml', '.forgejo/workflows/release-images.yml']) {
      const workflow = readWorkflow(path)

      expect(workflowConcurrency(workflow)).toEqual({
        group: '${{ github.workflow }}-${{ github.ref }}',
        'cancel-in-progress': true,
      })
    }
  })

  it('keeps scheduled correctness runs reproducible and singleton', () => {
    const path = '.forgejo/workflows/forgejo-nightly-fuzz.yml'
    const workflow = readWorkflow(path)
    const source = readFileSync(resolve(repoRoot, path), 'utf8')

    expect(workflowConcurrency(workflow)).toEqual({
      group: '${{ github.workflow }}-${{ github.ref }}',
      'cancel-in-progress': true,
    })
    expect(source).toContain('npm install -g bun@1.3.10')
    expect(source).toContain('test "$(bun --version)" = "1.3.10"')
    expect(source).toContain(
      'mcr.microsoft.com/playwright:v1.58.2-noble@sha256:6446946a1d9fd62d9ae501312a2d76a43ee688542b21622056a372959b65d63d',
    )
    expect(source).not.toContain('npm install -g bun\n')
  })
})
