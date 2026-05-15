export const agentFrameworkLlmsRequiredLinks = [
  'https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html',
  'https://proompteng.github.io/bilig/copilotkit-workpaper-spreadsheet-action.html',
  'https://proompteng.github.io/bilig/cloudflare-agents-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/crewai-workpaper-spreadsheet-tool.html',
] as const

export const agentFrameworkDocRequirements = [
  {
    path: 'docs/mastra-workpaper-spreadsheet-tool.md',
    includes: ['Mastra WorkPaper spreadsheet tool', 'createTool', 'npm run agent:framework-adapters'],
  },
  {
    path: 'docs/llamaindex-workpaper-spreadsheet-tool.md',
    includes: ['LlamaIndex.TS WorkPaper spreadsheet tool', 'tool(fn, { parameters })', 'npm run agent:framework-adapters'],
  },
  {
    path: 'docs/langgraph-workpaper-toolnode-spreadsheet.md',
    includes: ['LangGraph.js WorkPaper ToolNode spreadsheet tool', 'ToolNode', 'npm run agent:framework-adapters'],
  },
  {
    path: 'docs/copilotkit-workpaper-spreadsheet-action.md',
    includes: ['CopilotKit WorkPaper spreadsheet action', 'useCopilotAction', 'npm run agent:framework-adapters'],
  },
  {
    path: 'docs/cloudflare-agents-workpaper-spreadsheet-tool.md',
    includes: ['Cloudflare Agents WorkPaper spreadsheet tool', 'agentTool', 'npm run agent:framework-adapters'],
  },
  {
    path: 'docs/crewai-workpaper-spreadsheet-tool.md',
    includes: ['CrewAI WorkPaper spreadsheet tool', 'JSON contract', 'npm run agent:framework-adapters'],
  },
] as const
