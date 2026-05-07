import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

describe('public workbook corpus CLI resource guards', () => {
  it('refuses in-process verification unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_VERIFY

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--in-process'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--in-process is disabled for public corpus CLI runs')
  })

  it('refuses in-process fingerprinting unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_FINGERPRINT

    const result = spawnSync('bun', [corpusScriptPath(), 'fetch', '--in-process-fingerprint'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--in-process-fingerprint is disabled for public corpus CLI runs')
  })

  it('refuses parallel verification unless explicitly enabled for a sized host', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--verify-concurrency', '4'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--verify-concurrency greater than 1 is disabled for public corpus CLI runs')
  })
})

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}
