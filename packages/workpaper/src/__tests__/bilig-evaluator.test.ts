import { describe, expect, it } from 'vitest'

import { buildBiligEvaluatorProof, listBiligEvaluatorDoors, runBiligEvaluatorCli } from '../evaluator.js'

describe('bilig-workpaper evaluator', () => {
  it('lists only WorkPaper evaluator doors', () => {
    expect(listBiligEvaluatorDoors().map((door) => door.door)).toEqual(['workpaper-service', 'agent-mcp'])
  })

  it('prints a verified WorkPaper service proof', async () => {
    const proof = await buildBiligEvaluatorProof('workpaper-service')

    expect(proof.verified).toBe(true)
    expect(proof.packageVersions).toHaveProperty('@bilig/workpaper')
    expect(proof.packageVersions).not.toHaveProperty('@bilig/xlsx-formula-recalc')
    expect(proof.evidence).toMatchObject({
      editedCell: 'Inputs!B2',
      dependentCell: 'Summary!B2',
      before: 24_000,
      after: 38_400,
      afterRestore: 38_400,
      checks: {
        formulaReadbackChanged: true,
        exportedWorkPaperDocument: true,
        restoredMatchesAfter: true,
      },
    })
  })

  it('prints a verified agent MCP proof', async () => {
    const proof = await buildBiligEvaluatorProof('agent-mcp')

    expect(proof.verified).toBe(true)
    expect(proof.packageVersions).toHaveProperty('@bilig/workpaper')
    expect(proof.packageVersions).not.toHaveProperty('@bilig/xlsx-formula-recalc')
    expect(proof.evidence).toMatchObject({
      editedCell: 'Inputs!B3',
      dependentCell: 'Summary!B3',
      before: 60_000,
      after: 96_000,
      afterRestore: 96_000,
      afterRestart: 96_000,
      checks: {
        listedFileBackedTools: true,
        restartReadbackMatchesAfter: true,
      },
    })
    expect(proof.evidence.tools).toContain('read_cell')
  })

  it('prints a verified provider-backed agent MCP proof from the WorkPaper package', async () => {
    let stdout = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'agent-mcp', '--scenario', 'provider-backed', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const proof = readProof(stdout)
    expect(proof.schemaVersion).toBe('bilig-evaluator.v1')
    expect(proof.door).toBe('agent-mcp')
    expect(proof.verified).toBe(true)
    expect(proof.evidence).toMatchObject({
      scenario: 'provider-backed',
      providerFunction: 'IMPORTRANGE',
      adapterSurface: 'web',
      target: 'Imports!B2',
      formula: '=IMPORTRANGE("source","Revenue!B2")',
      adapterFormula: '=IMPORTRANGE("source","Revenue!B2")+0',
      before: {
        displayValue: '#BLOCKED!',
      },
      after: {
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
      },
      afterRestart: {
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
      },
      checks: {
        blockedReadbackIsBlocked: true,
        blockedDiagnosticExplainsAdapter: true,
        adapterBackedReadbackIsFresh: true,
        adapterBackedDiagnosticsCleared: true,
        restartReadbackMatchesAfter: true,
      },
    })
  })

  it('rejects XLSX doors from the WorkPaper evaluator package', async () => {
    let stderr = ''

    const exitCode = await runBiligEvaluatorCli(['--door', 'xlsx-cache'], {
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Unknown bilig-evaluate door: xlsx-cache')
  })
})

interface EvaluatorProofForTest {
  readonly schemaVersion: string
  readonly door: string
  readonly verified: boolean
  readonly evidence: Record<string, unknown>
}

function readProof(stdout: string): EvaluatorProofForTest {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error('Expected evaluator proof object.')
  }
  const evidence = parsed.evidence
  if (!isRecord(evidence)) {
    throw new Error('Expected evaluator proof evidence object.')
  }
  return {
    schemaVersion: requireString(parsed.schemaVersion),
    door: requireString(parsed.door),
    verified: parsed.verified === true,
    evidence,
  }
}

function requireString(value: unknown): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected string, got ${JSON.stringify(value)}.`)
  }
  return value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
