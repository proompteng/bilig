import { describe, expect, it } from 'vitest'

import { buildMcpChallengeProof, mcpChallengeHelpText, parseMcpChallengeCliArgs, runMcpChallengeCli } from '../mcp-challenge-cli.js'

describe('bilig-mcp-challenge', () => {
  it('builds the verified file-backed MCP proof object', () => {
    expect(buildMcpChallengeProof()).toMatchObject({
      transport: 'stdio-json-rpc',
      serverName: 'bilig-headless-workpaper',
      tools: [
        'list_sheets',
        'read_range',
        'read_cell',
        'set_cell_contents',
        'get_cell_display_value',
        'export_workpaper_document',
        'validate_formula',
      ],
      resources: [
        'bilig://workpaper/manifest',
        'bilig://workpaper/agent-handoff',
        'bilig://workpaper/sheets',
        'bilig://workpaper/current-document',
      ],
      prompts: ['edit_and_verify_workpaper', 'debug_workpaper_formula'],
      editedCell: 'Inputs!B3',
      dependentCell: 'Summary!B3',
      before: 60_000,
      after: 96_000,
      afterRestart: 96_000,
      displayValue: '96000',
      persistence: {
        persisted: true,
      },
      checks: {
        listedFileBackedTools: true,
        listedResourcesAndPrompts: true,
        formulaValidationPassed: true,
        dependentCellChanged: true,
        persistedToDisk: true,
        exportContainsWorkPaperDocument: true,
        restartReadbackMatchesAfter: true,
        displayValueRead: true,
      },
      verified: true,
    })
  })

  it('prints JSON by default', () => {
    let stdout = ''
    const exitCode = runMcpChallengeCli({
      argv: [],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed).toMatchObject({
      editedCell: 'Inputs!B3',
      after: 96_000,
      afterRestart: 96_000,
      verified: true,
    })
    expect(parsed.workpaperPath).toBeUndefined()
  })

  it('prints a markdown report when requested', () => {
    let stdout = ''
    const exitCode = runMcpChallengeCli({
      argv: ['--markdown'],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig MCP challenge')
    expect(stdout).toContain('"verified": true')
    expect(stdout).toContain('Inputs!B3')
    expect(stdout).toContain('Summary!B3')
  })

  it('can keep the temporary WorkPaper path for debugging', () => {
    let stdout = ''
    const exitCode = runMcpChallengeCli({
      argv: ['--keep-temp'],
      writeStdout(text) {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const parsed = JSON.parse(stdout)
    expect(parsed.workpaperPath).toMatch(/pricing\.workpaper\.json$/)
  })

  it('validates arguments and help', () => {
    expect(parseMcpChallengeCliArgs(['--json', '--keep-temp'])).toEqual({
      help: false,
      keepTemp: true,
      outputMode: 'json',
    })
    expect(mcpChallengeHelpText()).toContain('Usage: bilig-mcp-challenge')
    expect(() => parseMcpChallengeCliArgs(['--bad'])).toThrow('Unknown bilig-mcp-challenge argument')
  })
})
