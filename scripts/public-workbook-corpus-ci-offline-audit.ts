export interface PublicWorkbookCorpusCiOfflineCachedModeAudit {
  readonly passed: boolean
  readonly evidence: readonly string[]
  readonly gaps: readonly string[]
}

interface PublicWorkbookCorpusCiPackageScriptPolicy {
  readonly name: string
  readonly requiredTokens: readonly string[]
  readonly requiredSubstrings?: readonly string[]
  readonly forbiddenTokens: readonly string[]
}

const ciOfflineCachedCorpusScriptNames = [
  'public-workbook-corpus:check:offline',
  'public-workbook-corpus:resume-plan:check',
  'public-workbook-corpus:resume-financial:check',
  'public-workbook-corpus:completion-audit:check',
  'test:correctness:corpus',
] as const
const ciUnsafeCorpusScriptTokens = [
  'discover',
  'discover-financial-ckan',
  'fetch',
  'fetch-source',
  'verify',
  'verify-missing',
  'verify-artifact',
  'refresh-scorecard-from-checkpoint',
  'add-link',
  '--allow-active-stop-marker',
  '--require-target',
  '--require-complete',
  'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1',
] as const
const ciOfflineCachedCorpusScriptPolicies: readonly PublicWorkbookCorpusCiPackageScriptPolicy[] = [
  {
    name: 'public-workbook-corpus:check:offline',
    requiredTokens: ['bun', 'scripts/public-workbook-corpus.ts', 'check', '--skip-manifest-check'],
    forbiddenTokens: ciUnsafeCorpusScriptTokens,
  },
  {
    name: 'public-workbook-corpus:resume-plan:check',
    requiredTokens: ['bun', 'scripts/public-workbook-corpus-resume-plan.ts', '--check'],
    forbiddenTokens: ciUnsafeCorpusScriptTokens,
  },
  {
    name: 'public-workbook-corpus:resume-financial:check',
    requiredTokens: [
      'bun',
      'scripts/public-workbook-corpus-resume-plan.ts',
      '--check',
      '--manifest',
      '.cache/public-workbook-corpus-financial/manifest.json',
      '--cache-dir',
      '.cache/public-workbook-corpus-financial',
      '--scorecard',
      '.cache/public-workbook-corpus-financial/scorecard.json',
      '--verify-checkpoint',
      '.cache/public-workbook-corpus-financial/verification-checkpoint.json',
      '--fetch-limit',
      '5000',
      '--fetch-batch-size',
      '6',
    ],
    forbiddenTokens: ciUnsafeCorpusScriptTokens,
  },
  {
    name: 'public-workbook-corpus:completion-audit:check',
    requiredTokens: ['bun', 'scripts/public-workbook-corpus-completion-audit.ts', '--check'],
    forbiddenTokens: ciUnsafeCorpusScriptTokens,
  },
  {
    name: 'test:correctness:corpus',
    requiredTokens: ['bun', 'scripts/run-vitest.ts', '--run'],
    requiredSubstrings: [
      'scripts/__tests__/public-workbook-corpus.test.ts',
      'scripts/__tests__/public-workbook-corpus-cli.test.ts',
      'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
      'scripts/__tests__/public-workbook-corpus-links.test.ts',
      'scripts/__tests__/public-workbook-corpus-verify-checkpoint.test.ts',
      'scripts/__tests__/public-workbook-corpus-workbook.test.ts',
      'packages/excel-import/src/__tests__/excel-import.test.ts',
      'packages/excel-import/src/__tests__/xlsx-export-large-simple.test.ts',
    ],
    forbiddenTokens: ciUnsafeCorpusScriptTokens,
  },
] as const

export function auditPublicWorkbookCorpusCiOfflineCachedMode(args: {
  readonly scripts: ReadonlyMap<string, string>
  readonly ciSource: string
}): PublicWorkbookCorpusCiOfflineCachedModeAudit {
  const evidence: string[] = []
  const gaps: string[] = []
  for (const policy of ciOfflineCachedCorpusScriptPolicies) {
    const command = args.scripts.get(policy.name)
    if (!command) {
      gaps.push(`missing package script: ${policy.name}`)
      continue
    }
    const tokens = splitPackageScriptCommand(command)
    const missingRequiredTokens = policy.requiredTokens.filter((token) => !tokens.includes(token))
    if (missingRequiredTokens.length > 0) {
      gaps.push(`package script ${policy.name} missing required tokens: ${missingRequiredTokens.join(', ')}`)
    }
    const missingRequiredSubstrings = (policy.requiredSubstrings ?? []).filter((substring) => !command.includes(substring))
    if (missingRequiredSubstrings.length > 0) {
      gaps.push(`package script ${policy.name} missing required coverage files: ${missingRequiredSubstrings.join(', ')}`)
    }
    const forbiddenTokens = [
      ...new Set(tokens.filter((token) => policy.forbiddenTokens.some((forbidden) => matchesPackageScriptToken(token, forbidden)))),
    ]
    if (forbiddenTokens.length > 0) {
      gaps.push(`package script ${policy.name} uses CI-unsafe corpus tokens: ${forbiddenTokens.join(', ')}`)
    }
    evidence.push(`package script ${policy.name}: ${command}`)
  }
  for (const scriptName of ciOfflineCachedCorpusScriptNames) {
    const policy = ciOfflineCachedCorpusScriptPolicies.find((entry) => entry.name === scriptName)
    if (ciSourceInvokesPackageScript(args.ciSource, scriptName)) {
      evidence.push(`CI invokes package script: ${scriptName}`)
    } else if (policy && ciSourceInvokesEquivalentDirectGate(args.ciSource, policy)) {
      evidence.push(`CI invokes equivalent direct gate: ${scriptName}`)
    } else {
      gaps.push(`CI does not invoke package script or equivalent direct gate: ${scriptName}`)
    }
  }
  return {
    passed: gaps.length === 0,
    evidence,
    gaps,
  }
}

function splitPackageScriptCommand(command: string): readonly string[] {
  return command
    .trim()
    .split(/\s+/u)
    .filter((token) => token.length > 0)
}

function matchesPackageScriptToken(token: string, forbidden: string): boolean {
  return forbidden.endsWith('*') ? token.startsWith(forbidden.slice(0, -1)) : token === forbidden
}

function ciSourceInvokesPackageScript(ciSource: string, scriptName: string): boolean {
  const escapedScriptName = escapeRegExp(scriptName)
  return new RegExp(`['"]${escapedScriptName}['"]`, 'u').test(ciSource)
}

function ciSourceInvokesEquivalentDirectGate(ciSource: string, policy: PublicWorkbookCorpusCiPackageScriptPolicy): boolean {
  const requiredSourceLiterals = policy.requiredTokens.filter((token) => token !== 'bun')
  if (requiredSourceLiterals.length === 0) {
    return false
  }
  return sourceContainsOrderedStringLiterals(ciSource, requiredSourceLiterals)
}

function sourceContainsOrderedStringLiterals(source: string, literals: readonly string[]): boolean {
  const orderedLiteralPattern = literals.map((literal) => `['"]${escapeRegExp(literal)}['"]`).join('[\\s\\S]{0,240}')
  return new RegExp(orderedLiteralPattern, 'u').test(source)
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
