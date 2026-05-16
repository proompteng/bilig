import type { BenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'

type RequireIncludes = (haystack: string, needle: string, context: string) => void

export function requireFormulaProofDiscovery({
  benchmarkEvidence,
  communityLaunchPack,
  formulaWorkbooksProof,
  headlessReadme,
  index,
  llms,
  readme,
  requireIncludes,
  showHnFormulaWorkbooksProof,
}: {
  readonly benchmarkEvidence: BenchmarkDiscoveryEvidence
  readonly communityLaunchPack: string
  readonly formulaWorkbooksProof: string
  readonly headlessReadme: string
  readonly index: string
  readonly llms: string
  readonly readme: string
  readonly requireIncludes: RequireIncludes
  readonly showHnFormulaWorkbooksProof: string
}): void {
  for (const required of [
    'title: Formula workbooks for Node services and agent tools',
    'npm install @bilig/headless',
    'quote-approval-api.ts',
    '"restoredMatchesAfter": true',
    'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable',
    'Use HyperFormula first when you need a mature, broad formula engine',
    'Use SheetJS or ExcelJS first when the primary job is reading, writing, styling',
    'Use Google Sheets API first when a shared hosted spreadsheet',
    `The current checked benchmark artifact records \`${benchmarkEvidence.meanWinHeadline}\` comparable`,
    benchmarkEvidence.p95HoldoutWorkload,
    'https://github.com/proompteng/bilig/stargazers',
    'https://github.com/proompteng/bilig/discussions/new?category=general',
    'adoption-blocker form',
  ] as const) {
    requireIncludes(formulaWorkbooksProof, required, 'docs/formula-workbooks-node-services-agent-tools.md')
  }

  requireIncludes(readme, 'formula workbooks proof page', 'README.md')
  requireIncludes(readme, 'docs/formula-workbooks-node-services-agent-tools.md', 'README.md')
  requireIncludes(headlessReadme, 'formula workbooks for Node services and agent tools', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'docs/formula-workbooks-node-services-agent-tools.md', 'packages/headless/README.md')
  requireIncludes(
    communityLaunchPack,
    'https://proompteng.github.io/bilig/formula-workbooks-node-services-agent-tools.html',
    'docs/community-launch-pack.md',
  )
  requireIncludes(communityLaunchPack, 'Hacker News Submission After The Formula Workbooks Page', 'docs/community-launch-pack.md')

  for (const required of [
    'title: Show HN: Bilig runs small formula workbooks in Node',
    `\`@bilig/headless@${benchmarkEvidence.packageVersion}\``,
    'curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts',
    '"verified": true',
    `wins \`${benchmarkEvidence.meanWinHeadline}\` comparable`,
    `\`${benchmarkEvidence.meanAndP95Headline}\` on both mean and p95`,
    `\`${benchmarkEvidence.p95HoldoutWorkload}\` is slower at`,
    'Show HN: Bilig runs small formula workbooks in Node',
    'https://github.com/proompteng/bilig/stargazers',
  ] as const) {
    requireIncludes(showHnFormulaWorkbooksProof, required, 'docs/show-hn-formula-workbooks-node-services.md')
  }

  requireIncludes(index, './show-hn-formula-workbooks-node-services.html', 'docs/index.html')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/show-hn-formula-workbooks-node-services.md', 'docs/llms.txt')
}
