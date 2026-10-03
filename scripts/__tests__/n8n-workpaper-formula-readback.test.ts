import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { createN8nForecastProof } from '@bilig/workpaper'

const n8nNodeDir = join(process.cwd(), 'integrations', 'n8n-nodes-workpaper')

describe('n8n WorkPaper community node', () => {
  it('keeps the community-node package metadata aligned with n8n install expectations', () => {
    const packageJson = readJsonRecord(join(n8nNodeDir, 'package.json'))
    const codexJson = readJsonRecord(join(n8nNodeDir, 'nodes', 'Workpaper', 'BiligWorkpaper.node.json'))
    const nodeSource = readFileSync(join(n8nNodeDir, 'nodes', 'Workpaper', 'BiligWorkpaper.node.ts'), 'utf8')
    const packageName = readRequiredString(packageJson['name'], 'package name')
    const keywords = readStringArray(packageJson['keywords'], 'package keywords')
    const n8n = readRecord(packageJson['n8n'], 'package n8n metadata')
    const peerDependencies = readRecord(packageJson['peerDependencies'], 'package peerDependencies')
    const devDependencies = readRecord(packageJson['devDependencies'], 'package devDependencies')
    const nodeCliVersion = readRequiredString(devDependencies['@n8n/node-cli'], '@n8n/node-cli version')

    expect(packageName).toBe('@bilig/n8n-nodes-workpaper')
    expect(packageJson['license']).toBe('MIT')
    expect(packageJson['dependencies']).toBeUndefined()
    expect(nodeCliVersion).not.toBe('*')
    expect(isAtLeastNodeCli023(nodeCliVersion)).toBe(true)
    expect(keywords).toContain('n8n-community-node-package')
    expect(n8n['strict']).toBe(true)
    expect(n8n['credentials']).toEqual([])
    expect(n8n['nodes']).toEqual(['dist/nodes/Workpaper/BiligWorkpaper.node.js'])
    expect(nodeSource).toContain("name: 'biligWorkpaper'")
    expect(nodeSource).toContain("operation: ['verifyReadback']")
    expect(nodeSource).toContain("operation: ['evaluateDocument']")
    expect(codexJson['node']).toBe(`${packageName}.biligWorkpaper`)
    expect(peerDependencies['n8n-workflow']).toBe('*')
  })

  it('keeps the compact forecast proof contract used by the n8n node', () => {
    const proof = createN8nForecastProof({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    })

    expect({
      verdict: 'verified',
      editedCell: proof.editedCell,
      beforeExpectedArr: proof.before.expectedArr,
      afterExpectedArr: proof.after.expectedArr,
      targetGap: proof.after.targetGap,
      checks: {
        formulasPersisted: proof.checks.formulasPersisted,
        restoredMatchesAfter: proof.checks.restoredMatchesAfter,
        computedOutputChanged: proof.checks.computedOutputChanged,
      },
    }).toEqual({
      verdict: 'verified',
      editedCell: 'Inputs!B3',
      beforeExpectedArr: 60000,
      afterExpectedArr: 96000,
      targetGap: 5600,
      checks: {
        formulasPersisted: true,
        restoredMatchesAfter: true,
        computedOutputChanged: true,
      },
    })
  })

  it('documents the retained package without requiring deleted workflow examples', () => {
    const nodeReadme = readFileSync(join(n8nNodeDir, 'README.md'), 'utf8')

    for (const required of [
      '@bilig/n8n-nodes-workpaper',
      'n8n community node',
      'verifyReadback',
      'evaluateDocument',
      '"expectedArr": 60000',
      '"expectedArr": 96000',
      '"computedOutputChanged": true',
    ]) {
      expect(nodeReadme).toContain(required)
    }
    expect(nodeReadme).not.toContain('examples/n8n-workpaper-formula-readback')
    expect(nodeReadme).not.toContain('bilig-workpaper-native-node.n8n.json')
  })
})

function readJsonRecord(path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  return readRecord(parsed, path)
}

function readRecord(value: unknown, context: string): Record<string, unknown> {
  if (!isObject(value)) {
    throw new Error(`${context} must be an object`)
  }
  return value
}

function readRequiredString(value: unknown, context: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${context} must be a string`)
  }
  return value
}

function readStringArray(value: unknown, context: string): string[] {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`${context} must be a string array`)
  }
  return value
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isAtLeastNodeCli023(version: string): boolean {
  const match = version.match(/\d+\.\d+\.\d+/u)
  if (!match) {
    return false
  }
  const [major = 0, minor = 0] = match[0].split('.').map((part) => Number.parseInt(part, 10))
  return major > 0 || minor >= 23
}
