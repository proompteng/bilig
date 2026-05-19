import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireAiSdkDiscovery } from './check-docs-discovery-ai-sdk.ts'
import { requireDocumentedScriptsExist, requireFile, requireIncludes } from './check-docs-discovery-core.ts'
import { requireNpmEvalDiscovery } from './check-docs-discovery-npm-eval.ts'
import { requireOpenAiResponsesDiscovery } from './check-docs-discovery-openai-responses.ts'
import { requireServerlessWorkPaperApiDiscovery } from './check-docs-discovery-serverless.ts'

function parsePackageVersion(packageJson: string): string {
  const parsed: unknown = JSON.parse(packageJson)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('packages/headless/package.json must be an object')
  }
  const version = Reflect.get(parsed, 'version')
  if (typeof version !== 'string') {
    throw new Error('packages/headless/package.json must define a string version')
  }
  return version
}

export async function requireHeadlessExampleDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  index,
  llms,
  agentToolCallingDoc,
  aiSdkLangChainDoc,
}: {
  repoRoot: string
  docsRoot: string
  readme: string
  headlessReadme: string
  index: string
  llms: string
  agentToolCallingDoc: string
  aiSdkLangChainDoc: string
}): Promise<void> {
  const [headlessExampleReadme, headlessExamplePackage, headlessPackageManifest, headlessServerJson] = await Promise.all([
    readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'examples', 'headless-workpaper', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'server.json'), 'utf8'),
  ])
  const headlessPackageSpec = `@bilig/headless@${parsePackageVersion(headlessPackageManifest)}`
  await requireNpmEvalDiscovery(repoRoot, docsRoot, readme, headlessReadme, headlessExampleReadme)
  await requireOpenAiResponsesDiscovery({
    repoRoot,
    docsRoot,
    readme,
    headlessReadme,
    index,
    llms,
    agentToolCallingDoc,
    headlessExampleReadme,
    headlessExamplePackage,
  })
  await requireAiSdkDiscovery({
    repoRoot,
    docsRoot,
    readme,
    headlessReadme,
    index,
    llms,
    agentToolCallingDoc,
    aiSdkLangChainDoc,
    headlessExampleReadme,
    headlessExamplePackage,
  })
  await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'agent-framework-adapters.ts'))
  await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-tool-server.ts'))
  await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-stdio-server.ts'))
  requireDocumentedScriptsExist(headlessExampleReadme, headlessExamplePackage, 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '90-second npm-only check', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run invoice-totals', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Invoice Totals', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run csv-shaped', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## CSV Shaped Input', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessReadme, 'npm run csv-shaped', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'CSV shaped input', 'packages/headless/README.md')
  requireIncludes(headlessExampleReadme, 'npm run budget-variance', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Budget Variance Alerts', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run fulfillment-capacity', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Fulfillment Capacity Plan', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run quote-approval', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Quote Approval Threshold', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run subscription-mrr', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Subscription MRR Forecast', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run agent:framework-adapters', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## Agent Framework Adapters', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run agent:mcp-tools', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'npm run agent:mcp-stdio', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessReadme, 'npm run --silent agent:mcp-transcript', 'packages/headless/README.md')
  requireIncludes(headlessExampleReadme, '## MCP Tool Server Shape', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## MCP Stdio Server', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'annotations.', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, 'read tool is annotated as read-only', 'examples/headless-workpaper/README.md')
  requireIncludes(
    headlessExamplePackage,
    '"agent:framework-adapters": "node --disable-warning=DEP0205 --import tsx agent-framework-adapters.ts"',
    'examples/headless-workpaper/package.json',
  )
  requireIncludes(
    headlessExamplePackage,
    '"agent:mcp-tools": "node --disable-warning=DEP0205 --import tsx mcp-tool-server.ts"',
    'examples/headless-workpaper/package.json',
  )
  requireIncludes(
    headlessExamplePackage,
    '"agent:mcp-stdio": "node --disable-warning=DEP0205 --import tsx mcp-stdio-server.ts"',
    'examples/headless-workpaper/package.json',
  )
  await requireServerlessWorkPaperApiDiscovery({
    repoRoot,
    docsRoot,
    readme,
    headlessReadme,
    llms,
  })
  requireIncludes(headlessPackageManifest, '"mcpName": "io.github.proompteng/bilig-workpaper"', 'packages/headless/package.json')
  requireIncludes(headlessPackageManifest, '"bilig-formula-clinic": "./dist/formula-clinic-bin.js"', 'packages/headless/package.json')
  requireIncludes(headlessPackageManifest, '"bilig-mcp-challenge": "./dist/mcp-challenge-bin.js"', 'packages/headless/package.json')
  requireIncludes(headlessPackageManifest, '"bilig-workpaper-mcp": "./dist/work-paper-mcp-stdio-bin.js"', 'packages/headless/package.json')
  requireIncludes(
    headlessReadme,
    `npm exec --package ${headlessPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"`,
    'packages/headless/README.md',
  )
  requireIncludes(
    headlessReadme,
    'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
    'packages/headless/README.md',
  )
  requireIncludes(headlessReadme, '`set_cell_contents` edits back to the same file', 'packages/headless/README.md')
  requireIncludes(
    readme,
    `npm exec --package ${headlessPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"`,
    'README.md',
  )
  requireIncludes(readme, 'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable', 'README.md')
  requireIncludes(readme, '`export_workpaper_document`, and `validate_formula`', 'README.md')
  requireIncludes(headlessServerJson, '"name": "io.github.proompteng/bilig-workpaper"', 'packages/headless/server.json')
  requireIncludes(headlessServerJson, '"identifier": "@bilig/headless"', 'packages/headless/server.json')
  requireIncludes(headlessReadme, 'npm run invoice-totals', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#invoice-totals', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run budget-variance', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#budget-variance-alerts', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run fulfillment-capacity', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#fulfillment-capacity-plan', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run quote-approval', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#quote-approval-threshold', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run subscription-mrr', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#subscription-mrr-forecast', 'packages/headless/README.md')

  for (const required of [
    '## Clean npm Sanity Check',
    `npm exec --package ${headlessPackageSpec} -- bilig-agent-challenge`,
    `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge`,
    'https://proompteng.github.io/bilig/npm-eval.ts',
    'examples/headless-workpaper/npm-eval.ts',
    'afterRestore',
    '`checks.restoredMatchesAfter`',
    'matching `after`/`afterRestore` values are',
  ]) {
    requireIncludes(headlessReadme, required, 'packages/headless/README.md')
  }
}
