import { describe, expect, it } from 'vitest'

import {
  agentWorkbookChallengeHelpText,
  buildAgentWorkbookChallengeProof,
  parseAgentWorkbookChallengeCliArgs,
  runAgentWorkbookChallengeCli,
} from '../agent-workbook-challenge-cli.js'

describe('bilig-agent-challenge', () => {
  it('builds the verified WorkPaper proof object', () => {
    const proof = buildAgentWorkbookChallengeProof()
    expect(proof).toMatchObject({
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
      verified: true,
    })
    expect(proof.limitations).toEqual([
      'This challenge proves the WorkPaper write/read/persist loop for service-owned JSON state, not every desktop spreadsheet feature.',
      'Use saved-file compatibility evaluators only when a workbook file is the integration contract.',
    ])
  })

  it('prints JSON by default', () => {
    let stdout = ''
    const exitCode = runAgentWorkbookChallengeCli({
      argv: [],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed).toMatchObject({
      editedCell: 'Inputs!B2',
      after: 38_400,
      verified: true,
    })
    expect(parsed).not.toHaveProperty('star')
    expect(parsed).not.toHaveProperty('watchReleases')
    expect(parsed).not.toHaveProperty('adoptionBlocker')
    expect(parsed).not.toHaveProperty('nextStep')
  })

  it('prints a markdown report when requested', () => {
    let stdout = ''
    const exitCode = runAgentWorkbookChallengeCli({
      argv: ['--markdown'],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig agent workbook challenge')
    expect(stdout).toContain('"verified": true')
    expect(stdout).toContain('Inputs!B2')
  })

  it('validates arguments and help', () => {
    expect(parseAgentWorkbookChallengeCliArgs(['--json'])).toEqual({
      help: false,
      outputMode: 'json',
    })
    expect(agentWorkbookChallengeHelpText()).toContain('Usage: bilig-agent-challenge')
    expect(() => parseAgentWorkbookChallengeCliArgs(['--bad'])).toThrow('Unknown bilig-agent-challenge argument')
  })
})
