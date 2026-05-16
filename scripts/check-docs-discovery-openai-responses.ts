import { readFile, stat } from 'node:fs/promises'
import { join } from 'node:path'

type OpenAiResponsesDiscoveryInput = {
  repoRoot: string
  docsRoot: string
  readme: string
  headlessReadme: string
  index: string
  llms: string
  agentToolCallingDoc: string
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

export async function requireOpenAiResponsesDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  index,
  llms,
  agentToolCallingDoc,
  headlessExampleReadme,
  headlessExamplePackage,
}: OpenAiResponsesDiscoveryInput): Promise<void> {
  const openAiDoc = await readFile(join(docsRoot, 'openai-responses-workpaper-tool-call.md'), 'utf8')

  await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'openai-responses-tool-wrapper.ts'))

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['docs/openai-responses-workpaper-tool-call.md', openAiDoc],
    ['examples/headless-workpaper/README.md', headlessExampleReadme],
  ] as const) {
    requireIncludes(content, 'npm run agent:openai-responses', path)
  }

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
  ] as const) {
    requireIncludes(content, 'openai-responses-workpaper-tool-call', path)
  }

  for (const required of [
    'title: OpenAI Responses WorkPaper tool calls',
    'description: Run @bilig/headless behind OpenAI Responses function calls',
    'image: /assets/github-social-preview.png',
    'function_call',
    'function_call_output',
    'examples/headless-workpaper/openai-responses-tool-wrapper.ts',
    'https://platform.openai.com/docs/guides/function-calling?api-mode=responses',
  ]) {
    requireIncludes(openAiDoc, required, 'docs/openai-responses-workpaper-tool-call.md')
  }

  requireIncludes(
    headlessExamplePackage,
    '"agent:openai-responses": "node --disable-warning=DEP0205 --import tsx openai-responses-tool-wrapper.ts"',
    'examples/headless-workpaper/package.json',
  )
  requireIncludes(headlessExampleReadme, '## OpenAI Responses Tool Wrapper', 'examples/headless-workpaper/README.md')
}
