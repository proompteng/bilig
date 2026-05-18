import { describe, expect, it } from 'vitest'
import { homepageRequiredLinks, llmsRequiredLinks } from '../check-docs-discovery-public-link-manifest.ts'

function duplicateValues(values: readonly string[]): string[] {
  const seen = new Set<string>()
  const duplicates = new Set<string>()
  for (const value of values) {
    if (seen.has(value)) {
      duplicates.add(value)
    }
    seen.add(value)
  }
  return [...duplicates].toSorted()
}

describe('docs discovery public link manifest', () => {
  it('keeps homepage and llms public-link manifests centralized and duplicate-free', () => {
    expect(homepageRequiredLinks).toContain('./why-use-bilig.html')
    expect(homepageRequiredLinks).toContain('https://github.com/proompteng/bilig/stargazers')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/.well-known/mcp/server-card.json')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/.well-known/agent.json')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/agent.json')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/skill.txt')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/llms-full.txt')
    expect(llmsRequiredLinks).toContain('https://proompteng.github.io/bilig/.well-known/agent-skills/index.json')
    expect(llmsRequiredLinks).toContain('https://github.com/proompteng/bilig/blob/main/docs/npm-provenance-package-trust.md')
    expect(duplicateValues(homepageRequiredLinks)).toEqual([])
    expect(duplicateValues(llmsRequiredLinks)).toEqual([])
  })
})
