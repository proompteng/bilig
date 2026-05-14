import { requireAiSdkGenerateTextDiscovery } from './check-docs-discovery-ai-sdk-generate-text.ts'
import { requireAiSdkStreamTextDiscovery } from './check-docs-discovery-ai-sdk-stream-text.ts'

type AiSdkDiscoveryInput = {
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

export async function requireAiSdkDiscovery(input: AiSdkDiscoveryInput): Promise<void> {
  await Promise.all([requireAiSdkGenerateTextDiscovery(input), requireAiSdkStreamTextDiscovery(input)])
}
