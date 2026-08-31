import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)

function readPackageScripts(): Record<string, string> {
  const parsed: unknown = JSON.parse(readFileSync(resolve(repoRoot, 'package.json'), 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('package.json must be an object')
  }
  const scripts = Reflect.get(parsed, 'scripts')
  if (typeof scripts !== 'object' || scripts === null || Array.isArray(scripts)) {
    throw new Error('package.json scripts must be an object')
  }
  return Object.fromEntries(Object.entries(scripts).filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
}

describe('repository package scripts', () => {
  it('references existing test files in the targeted browser correctness lane', () => {
    const scripts = readPackageScripts()
    const command = scripts['test:correctness:browser']
    expect(command).toBeDefined()

    const referencedPaths = [...command.matchAll(/(?:^|\s)((?:apps|packages)\/[^\s]+\.test\.tsx?)/gu)].map((match) => match[1])

    expect(referencedPaths.length).toBeGreaterThan(0)
    expect(referencedPaths.every((path) => path !== undefined && existsSync(resolve(repoRoot, path)))).toBe(true)
  })
})
