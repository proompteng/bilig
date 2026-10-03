import { describe, expect, it } from 'vitest'

import {
  buildMcpChallengeProof,
  buildMcpProviderBackedChallengeProof,
  buildMcpRevenuePlanChallengeProof,
  mcpChallengeHelpText,
  parseMcpChallengeCliArgs,
  runMcpChallengeCli,
} from '../mcp-challenge-cli.js'

describe('bilig-mcp-challenge', () => {
  it('builds the verified file-backed MCP proof object', () => {
    const proof = buildMcpChallengeProof()
    expect(proof).toMatchObject({
      transport: 'stdio-json-rpc',
      serverName: 'bilig-headless-workpaper',
      tools: [
        'list_sheets',
        'read_range',
        'read_cell',
        'set_cell_contents',
        'set_cell_contents_and_readback',
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
      afterRestore: 96_000,
      afterRestart: 96_000,
      persistedDocumentBytes: expect.any(Number),
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
    expect(proof.limitations).toEqual([
      'This challenge proves the file-backed MCP WorkPaper tool surface, not every desktop spreadsheet feature.',
      'Use saved-file compatibility evaluators only when a workbook file is the integration contract.',
    ])
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
      afterRestore: 96_000,
      afterRestart: 96_000,
      persistedDocumentBytes: expect.any(Number),
      verified: true,
    })
    expect(parsed).not.toHaveProperty('star')
    expect(parsed).not.toHaveProperty('watchReleases')
    expect(parsed).not.toHaveProperty('adoptionBlocker')
    expect(parsed).not.toHaveProperty('nextStep')
    expect(parsed.persistedDocumentBytes).toBe(parsed.persistence.serializedBytes)
    expect(parsed.workpaperPath).toBeUndefined()
  })

  it('builds a verified revenue-plan MCP scenario with multiple formula families', () => {
    const proof = buildMcpRevenuePlanChallengeProof()
    expect(proof).toMatchObject({
      transport: 'stdio-json-rpc',
      scenario: 'revenue-plan',
      serverName: 'bilig-headless-workpaper',
      formulaFamilies: ['SUM', 'SUMIF', 'XLOOKUP', 'FILTER', 'named-expression', 'persistence', 'restart'],
      editedCell: 'Deals!C2',
      readbackRange: 'Summary!B2:B8',
      before: {
        totalRevenue: 27_300,
        westCustomers: 30,
        enterpriseArpa: 1_200,
        targetRevenue: 30_576,
        qualifiedCustomerCounts: [30, 18],
      },
      after: {
        totalRevenue: 36_900,
        westCustomers: 38,
        enterpriseArpa: 1_200,
        targetRevenue: 41_328,
        qualifiedCustomerCounts: [20, 30, 18],
      },
      afterRestart: {
        totalRevenue: 36_900,
        westCustomers: 38,
        enterpriseArpa: 1_200,
        targetRevenue: 41_328,
        qualifiedCustomerCounts: [20, 30, 18],
      },
      checks: {
        listedFileBackedTools: true,
        listedResourcesAndPrompts: true,
        formulaValidationPassed: true,
        totalRevenueRecalculated: true,
        sumifReadbackChanged: true,
        xlookupReadbackStable: true,
        filterSpillUpdated: true,
        namedExpressionApplied: true,
        persistedToDisk: true,
        restoredReadbackMatchesAfter: true,
        exportContainsWorkPaperDocument: true,
        restartReadbackMatchesAfter: true,
      },
      verified: true,
    })
    expect(proof.limitations).toEqual([
      'This revenue-plan challenge proves a realistic file-backed MCP WorkPaper path, not every desktop spreadsheet feature.',
      'It covers formula readback for SUM, SUMIF, XLOOKUP, FILTER spills, and a named expression; use real workbook fixtures for domain-specific parity.',
    ])
  })

  it('builds a verified provider-backed MCP scenario with blocked diagnostics and adapter readback', () => {
    expect(buildMcpProviderBackedChallengeProof()).toMatchObject({
      transport: 'stdio-json-rpc',
      scenario: 'provider-backed',
      providerFunction: 'IMPORTRANGE',
      adapterSurface: 'web',
      serverName: 'bilig-headless-workpaper',
      target: 'Imports!B2',
      formula: '=IMPORTRANGE("source","Revenue!B2")',
      adapterFormula: '=IMPORTRANGE("source","Revenue!B2")+0',
      before: {
        address: 'Imports!B2',
        displayValue: '#BLOCKED!',
        formulaDiagnostics: [
          {
            code: 'provider-backed-adapter-missing',
            functionName: 'IMPORTRANGE',
            adapterSurface: 'web',
            errorText: '#BLOCKED!',
          },
        ],
      },
      after: {
        address: 'Imports!B2',
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
        formulaDiagnostics: [],
      },
      afterRestart: {
        address: 'Imports!B2',
        serialized: '=IMPORTRANGE("source","Revenue!B2")+0',
        displayValue: '96000',
        formulaDiagnostics: [],
      },
      checks: {
        listedFileBackedTools: true,
        listedResourcesAndPrompts: true,
        formulaValidationPassed: true,
        blockedReadbackIsBlocked: true,
        blockedDiagnosticExplainsAdapter: true,
        adapterBackedReadbackIsFresh: true,
        adapterBackedDiagnosticsCleared: true,
        exportContainsWorkPaperDocument: true,
        restartReadbackMatchesAfter: true,
      },
      verified: true,
    })
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
