import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..')

function source(relativePath: string): string {
  return readFileSync(join(srcDir, relativePath), 'utf8')
}

function sourceLineCount(relativePath: string): number {
  return source(relativePath).split('\n').length
}

describe('runtime session module boundary', () => {
  it('keeps authoritative sync validation outside the session controller', () => {
    const sessionSource = source('runtime-session.ts')
    const syncSource = source('runtime-authoritative-sync.ts')

    expect(sourceLineCount('runtime-session.ts')).toBeLessThan(900)
    expect(sourceLineCount('runtime-authoritative-sync.ts')).toBeLessThan(180)
    expect(sessionSource).not.toContain('isAuthoritativeWorkbookEventBatchAfterRevision')
    expect(sessionSource).not.toContain('Failed to load authoritative events')
    expect(syncSource).toContain('isAuthoritativeWorkbookEventBatchAfterRevision')
    expect(syncSource).toContain('Failed to load authoritative events')
  })
})
