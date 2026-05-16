import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

type ServerlessDiscoveryInput = {
  docsRoot: string
  headlessReadme: string
  llms: string
  readme: string
  repoRoot: string
}

export async function requireServerlessWorkPaperApiDiscovery({
  docsRoot,
  headlessReadme,
  llms,
  readme,
  repoRoot,
}: ServerlessDiscoveryInput): Promise<void> {
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'framework-adapters.ts'))
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'next-route-handler.ts'))
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'next-server-action.ts'))
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'next-server-action-formdata.ts'))
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'next-server-action-validation.ts'))
  await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'persistence-adapters.ts'))

  const [
    serverlessExampleReadme,
    serverlessExamplePackage,
    serverlessFrameworkAdapters,
    serverlessWorkPaperApiRouteDoc,
    nodeFrameworkAdapterDoc,
    persistenceDoc,
  ] = await Promise.all([
    readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'framework-adapters.ts'), 'utf8'),
    readFile(join(docsRoot, 'serverless-workpaper-api-route.md'), 'utf8'),
    readFile(join(docsRoot, 'node-framework-workpaper-adapters.md'), 'utf8'),
    readFile(join(docsRoot, 'persisting-formula-backed-workpaper-documents-in-node.md'), 'utf8'),
  ])

  requireIncludes(serverlessExampleReadme, 'npm run next-route-handler', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Next.js App Router Smoke', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, 'npm run next-route-handler', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(serverlessExampleReadme, 'npm run next-server-action', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Next.js Server Action Smoke', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, 'npm run next-server-action', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, '## Next.js Server Action Adapter', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(llms, 'Next.js Server Action WorkPaper smoke', 'docs/llms.txt')
  requireIncludes(llms, 'npm run next-server-action', 'docs/llms.txt')
  requireIncludes(readme, 'npm run next-server-action', 'README.md')
  requireIncludes(serverlessExampleReadme, 'npm run next-server-action-formdata', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Next.js Server Action FormData Smoke', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, 'npm run next-server-action-formdata', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, '## Next.js Server Action FormData Adapter', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(llms, 'Next.js Server Action FormData WorkPaper smoke', 'docs/llms.txt')
  requireIncludes(llms, 'npm run next-server-action-formdata', 'docs/llms.txt')
  requireIncludes(readme, 'npm run next-server-action-formdata', 'README.md')
  requireIncludes(serverlessExampleReadme, 'npm run next-server-action-validation', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Next.js Server Action Validation Smoke', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, 'npm run next-server-action-validation', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(serverlessWorkPaperApiRouteDoc, '## Next.js Server Action Validation Adapter', 'docs/serverless-workpaper-api-route.md')
  requireIncludes(llms, 'Next.js Server Action validation-error WorkPaper smoke', 'docs/llms.txt')
  requireIncludes(llms, 'npm run next-server-action-validation', 'docs/llms.txt')

  requireIncludes(serverlessExampleReadme, 'npm run framework-adapters', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Framework Adapters', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, 'Oak-style `context.request.source`', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, 'AdonisJS-style `HttpContext`', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, 'Hapi-style `request` plus `h.response()`', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(
    serverlessExampleReadme,
    '"adapters": ["fetch", "hono", "oak", "adonis", "hapi", "express", "fastify"]',
    'examples/serverless-workpaper-api/README.md',
  )
  requireIncludes(serverlessFrameworkAdapters, 'createOakWorkPaperHandler', 'examples/serverless-workpaper-api/framework-adapters.ts')
  requireIncludes(serverlessFrameworkAdapters, 'createOakWorkPaperRoutes', 'examples/serverless-workpaper-api/framework-adapters.ts')
  requireIncludes(serverlessFrameworkAdapters, 'createAdonisWorkPaperHandler', 'examples/serverless-workpaper-api/framework-adapters.ts')
  requireIncludes(serverlessFrameworkAdapters, 'createAdonisWorkPaperRoutes', 'examples/serverless-workpaper-api/framework-adapters.ts')
  requireIncludes(serverlessFrameworkAdapters, 'createHapiWorkPaperRoutes', 'examples/serverless-workpaper-api/framework-adapters.ts')
  requireIncludes(
    serverlessFrameworkAdapters,
    "adapters: ['fetch', 'hono', 'oak', 'adonis', 'hapi', 'express', 'fastify']",
    'examples/serverless-workpaper-api/framework-adapters.ts',
  )
  requireIncludes(nodeFrameworkAdapterDoc, '## Oak', 'docs/node-framework-workpaper-adapters.md')
  requireIncludes(nodeFrameworkAdapterDoc, 'createOakWorkPaperRoutes', 'docs/node-framework-workpaper-adapters.md')
  requireIncludes(nodeFrameworkAdapterDoc, '## AdonisJS', 'docs/node-framework-workpaper-adapters.md')
  requireIncludes(nodeFrameworkAdapterDoc, 'createAdonisWorkPaperRoutes', 'docs/node-framework-workpaper-adapters.md')
  requireIncludes(nodeFrameworkAdapterDoc, '## Hapi', 'docs/node-framework-workpaper-adapters.md')
  requireIncludes(nodeFrameworkAdapterDoc, 'createHapiWorkPaperRoutes', 'docs/node-framework-workpaper-adapters.md')

  requireIncludes(serverlessExampleReadme, 'npm run persistence-adapters', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, '## Persistence Adapters', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(serverlessExampleReadme, 'Postgres JSONB', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(
    serverlessExampleReadme,
    'SQLite, using parameterized SQL and a workbook id.',
    'examples/serverless-workpaper-api/README.md',
  )
  requireIncludes(serverlessExampleReadme, 'not an XLSX file cache', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(
    serverlessExampleReadme,
    '"adapters": ["postgres-jsonb", "sqlite", "redis", "object-storage"]',
    'examples/serverless-workpaper-api/README.md',
  )
  requireIncludes(serverlessExampleReadme, '"sqlite": {', 'examples/serverless-workpaper-api/README.md')
  requireIncludes(
    await readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'persistence-adapters.ts'), 'utf8'),
    'createSQLiteWorkPaperStorage',
    'examples/serverless-workpaper-api/persistence-adapters.ts',
  )
  requireIncludes(
    persistenceDoc,
    'examples/serverless-workpaper-api/persistence-adapters.ts',
    'docs/persisting-formula-backed-workpaper-documents-in-node.md',
  )
  requireIncludes(persistenceDoc, 'npm run persistence-adapters', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
  requireIncludes(persistenceDoc, 'Postgres JSONB', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
  requireIncludes(persistenceDoc, 'SQLite adapter', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
  requireIncludes(persistenceDoc, 'XLSX file cache', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
  requireIncludes(persistenceDoc, 'Redis or string-KV adapter', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
  requireIncludes(llms, 'Postgres JSONB, SQLite, Redis/KV, and object storage', 'docs/llms.txt')

  requireIncludes(
    serverlessExamplePackage,
    '"next-route-handler": "tsx next-route-handler.ts"',
    'examples/serverless-workpaper-api/package.json',
  )
  requireIncludes(
    serverlessExamplePackage,
    '"next-server-action": "tsx next-server-action.ts"',
    'examples/serverless-workpaper-api/package.json',
  )
  requireIncludes(
    serverlessExamplePackage,
    '"next-server-action-formdata": "tsx next-server-action-formdata.ts"',
    'examples/serverless-workpaper-api/package.json',
  )
  requireIncludes(
    serverlessExamplePackage,
    '"next-server-action-validation": "tsx next-server-action-validation.ts"',
    'examples/serverless-workpaper-api/package.json',
  )
  requireIncludes(
    serverlessExamplePackage,
    '"framework-adapters": "tsx framework-adapters.ts"',
    'examples/serverless-workpaper-api/package.json',
  )
  requireIncludes(
    serverlessExamplePackage,
    '"persistence-adapters": "tsx persistence-adapters.ts"',
    'examples/serverless-workpaper-api/package.json',
  )

  requireIncludes(headlessReadme, 'npm run next-route-handler', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run next-server-action', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#nextjs-server-action-smoke', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run next-server-action-formdata', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#nextjs-server-action-formdata-smoke', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run framework-adapters', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'npm run persistence-adapters', 'packages/headless/README.md')
  requireIncludes(headlessReadme, '#persistence-adapters', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'node-framework-workpaper-adapters.html', 'packages/headless/README.md')
}
