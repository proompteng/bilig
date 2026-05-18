import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const testDir = dirname(fileURLToPath(import.meta.url))

function source(relativePath: string): string {
  return readFileSync(join(testDir, relativePath), 'utf8')
}

function sourceLineCount(relativePath: string): number {
  return source(relativePath).split('\n').length
}

describe('workbook agent service module boundary', () => {
  it('keeps session authority and action context wiring out of the service orchestrator', () => {
    const serviceSource = source('workbook-agent-service.ts')
    const authoritySource = source('workbook-agent-session-authority.ts')

    expect(sourceLineCount('workbook-agent-service.ts')).toBeLessThan(860)
    expect(sourceLineCount('workbook-agent-session-authority.ts')).toBeLessThan(160)
    expect(sourceLineCount('workbook-agent-service-action-contexts.ts')).toBeLessThan(120)
    expect(serviceSource).not.toContain('assertWorkbookAgentSessionAccessPolicy')
    expect(authoritySource).toContain('assertWorkbookAgentSessionAccessPolicy')
  })
})
