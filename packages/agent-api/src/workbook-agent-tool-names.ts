export const WORKBOOK_AGENT_TOOL_NAMES = {
  getContext: "bilig_get_context",
  readWorkbook: "bilig_read_workbook",
  readRange: "bilig_read_range",
  readSelection: "bilig_read_selection",
  readVisibleRange: "bilig_read_visible_range",
  inspectCell: "bilig_inspect_cell",
  findFormulaIssues: "bilig_find_formula_issues",
  searchWorkbook: "bilig_search_workbook",
  traceDependencies: "bilig_trace_dependencies",
  writeRange: "bilig_write_range",
  clearRange: "bilig_clear_range",
  formatRange: "bilig_format_range",
  fillRange: "bilig_fill_range",
  copyRange: "bilig_copy_range",
  moveRange: "bilig_move_range",
  createSheet: "bilig_create_sheet",
  renameSheet: "bilig_rename_sheet",
} as const;

export type WorkbookAgentToolName =
  (typeof WORKBOOK_AGENT_TOOL_NAMES)[keyof typeof WORKBOOK_AGENT_TOOL_NAMES];

const LEGACY_WORKBOOK_AGENT_TOOL_NAME_MAP: Record<string, WorkbookAgentToolName> = {
  "bilig.get_context": WORKBOOK_AGENT_TOOL_NAMES.getContext,
  "bilig.read_workbook": WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
  "bilig.read_range": WORKBOOK_AGENT_TOOL_NAMES.readRange,
  "bilig.read_selection": WORKBOOK_AGENT_TOOL_NAMES.readSelection,
  "bilig.read_visible_range": WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
  "bilig.inspect_cell": WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
  "bilig.find_formula_issues": WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
  "bilig.search_workbook": WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
  "bilig.trace_dependencies": WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
  "bilig.write_range": WORKBOOK_AGENT_TOOL_NAMES.writeRange,
  "bilig.clear_range": WORKBOOK_AGENT_TOOL_NAMES.clearRange,
  "bilig.format_range": WORKBOOK_AGENT_TOOL_NAMES.formatRange,
  "bilig.fill_range": WORKBOOK_AGENT_TOOL_NAMES.fillRange,
  "bilig.copy_range": WORKBOOK_AGENT_TOOL_NAMES.copyRange,
  "bilig.move_range": WORKBOOK_AGENT_TOOL_NAMES.moveRange,
  "bilig.create_sheet": WORKBOOK_AGENT_TOOL_NAMES.createSheet,
  "bilig.rename_sheet": WORKBOOK_AGENT_TOOL_NAMES.renameSheet,
};

const workbookAgentToolNameSet = new Set<string>(Object.values(WORKBOOK_AGENT_TOOL_NAMES));

export function normalizeWorkbookAgentToolName(toolName: string): string {
  return LEGACY_WORKBOOK_AGENT_TOOL_NAME_MAP[toolName] ?? toolName;
}

export function isWorkbookAgentToolName(toolName: string): toolName is WorkbookAgentToolName {
  return workbookAgentToolNameSet.has(normalizeWorkbookAgentToolName(toolName));
}
