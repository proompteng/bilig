import { readFile, stat } from 'node:fs/promises'
import { join } from 'node:path'

type AiSdkStreamTextDiscoveryInput = {
  repoRoot: string
  docsRoot: string
  readme: string
  headlessReadme: string
  index: string
  llms: string
  agentToolCallingDoc: string
  aiSdkLangChainDoc: string
  headlessExampleReadme: string
  headlessExamplePackage: string
}

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

async function requireFile(path: string): Promise<void> {
  const info = await stat(path)
  if (!info.isFile()) {
    throw new Error(`${path} is not a file`)
  }
}

export async function requireAiSdkStreamTextDiscovery({
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
}: AiSdkStreamTextDiscoveryInput): Promise<void> {
  const scriptPath = join(repoRoot, 'examples', 'headless-workpaper', 'ai-sdk-stream-text-tool-smoke.ts')
  const sharedScriptPath = join(repoRoot, 'examples', 'headless-workpaper', 'ai-sdk-workpaper-tool-smoke-shared.ts')
  const script = await readFile(scriptPath, 'utf8')
  const sharedScript = await readFile(sharedScriptPath, 'utf8')
  const aiSdkDoc = await readFile(join(docsRoot, 'vercel-ai-sdk-langchain-spreadsheet-tool.md'), 'utf8')

  await requireFile(scriptPath)
  await requireFile(sharedScriptPath)

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['docs/vercel-ai-sdk-langchain-spreadsheet-tool.md', aiSdkLangChainDoc],
    ['examples/headless-workpaper/README.md', headlessExampleReadme],
  ] as const) {
    requireIncludes(content, 'npm run agent:ai-sdk-stream-text', path)
  }

  for (const [path, content] of [
    ['docs/llms.txt', llms],
    ['docs/vercel-ai-sdk-langchain-spreadsheet-tool.md', aiSdkDoc],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['examples/headless-workpaper/README.md', headlessExampleReadme],
  ] as const) {
    requireIncludes(content, 'ai-sdk-stream-text-tool-smoke.ts', path)
  }

  for (const required of [
    'streamText',
    'simulateReadableStream',
    'stepCountIs',
    'MockLanguageModelV3',
    'readWorkPaperSummary',
    'setWorkPaperInputCell',
    'AI SDK streamText -> tool -> execute',
    'result.steps',
    'tool-call',
    'tool-result',
    'text-delta',
  ]) {
    requireIncludes(script, required, 'examples/headless-workpaper/ai-sdk-stream-text-tool-smoke.ts')
  }

  for (const required of [
    'createAiSdkWorkPaperTools',
    'assertAiSdkWorkPaperSmokeProof',
    'tool(',
    'inputSchema',
    'formulasPersisted',
    'restoredMatchesAfter',
  ]) {
    requireIncludes(sharedScript, required, 'examples/headless-workpaper/ai-sdk-workpaper-tool-smoke-shared.ts')
  }

  for (const required of [
    'Real AI SDK `streamText()` Smoke',
    'simulateReadableStream',
    'https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling',
    'https://ai-sdk.dev/docs/reference/ai-sdk-core/tool',
    'https://ai-sdk.dev/docs/reference/ai-sdk-core/stream-text',
  ]) {
    requireIncludes(aiSdkDoc, required, 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  }

  requireIncludes(
    headlessExamplePackage,
    '"agent:ai-sdk-stream-text": "tsx ai-sdk-stream-text-tool-smoke.ts"',
    'examples/headless-workpaper/package.json',
  )
  requireIncludes(headlessExampleReadme, '## AI SDK StreamText Tool Smoke', 'examples/headless-workpaper/README.md')
}
