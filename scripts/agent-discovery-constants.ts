export const skillTags = ['agents', 'workbook-runtime', 'formulas', 'workpaper', 'mcp', 'typescript'] as const

export const mcpPromptNames = ['edit_and_verify_workpaper', 'debug_workpaper_formula'] as const

export const agentNotAFitBoundaries = [
  'manual spreadsheet editing as the main product',
  'Office macros or desktop Excel automation',
  'one-off arithmetic',
] as const

export const versionedStaticReferenceRoots = [
  'README.md',
  'packages/workpaper/README.md',
  'packages/workpaper/AGENTS.md',
  'packages/workpaper/SKILL.md',
  'skills/bilig-workpaper/SKILL.md',
  'docs/agent-workbook-challenge.md',
  'docs/agent-xlsx-risk-preflight.md',
  'docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
  'docs/claude-desktop-mcpb-workpaper.md',
  'docs/formula-bug-clinic.md',
  'docs/formula-workbooks-node-services-agent-tools.md',
  'docs/ai-agent-spreadsheet-tool-node.md',
  'docs/agent-framework-workbook-tools.md',
  'docs/headless-workpaper-agent-handbook.md',
  'docs/index.html',
  'docs/llms.txt',
  'docs/mcp-client-setup.md',
  'docs/open-webui-workpaper-mcp.md',
  'docs/openhands-workpaper-mcp.md',
  'docs/opencode-workpaper-mcp.md',
  'docs/open-multi-agent-workpaper-mcp.md',
  'docs/lobehub-workpaper-mcp.md',
  'docs/anythingllm-workpaper-mcp.md',
  'docs/sim-workpaper-mcp.md',
  'docs/mcp-spreadsheet-server-directory.md',
  'docs/mcp-workpaper-tool-server.md',
  'docs/spreadsheet-mcp-server-comparison.md',
  'docs/why-agents-need-workbook-apis.md',
] as const

export const versionedStaticReferenceExtensions = new Set(['.html', '.json', '.md', '.txt'])
