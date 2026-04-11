export const WORKBOOK_AGENT_TOOL_NAMES = {
  getContext: "get_context",
  readWorkbook: "read_workbook",
  readRange: "read_range",
  readSelection: "read_selection",
  readVisibleRange: "read_visible_range",
  readRecentChanges: "read_recent_changes",
  startWorkflow: "start_workflow",
  inspectCell: "inspect_cell",
  findFormulaIssues: "find_formula_issues",
  searchWorkbook: "search_workbook",
  traceDependencies: "trace_dependencies",
  writeRange: "write_range",
  clearRange: "clear_range",
  formatRange: "format_range",
  fillRange: "fill_range",
  copyRange: "copy_range",
  moveRange: "move_range",
  createSheet: "create_sheet",
  renameSheet: "rename_sheet",
  updateRowMetadata: "update_row_metadata",
  updateColumnMetadata: "update_column_metadata",
} as const;

export type WorkbookAgentToolName =
  (typeof WORKBOOK_AGENT_TOOL_NAMES)[keyof typeof WORKBOOK_AGENT_TOOL_NAMES];

const LEGACY_WORKBOOK_AGENT_TOOL_NAME_MAP: Record<string, WorkbookAgentToolName> = {
  bilig_get_context: WORKBOOK_AGENT_TOOL_NAMES.getContext,
  "bilig.get_context": WORKBOOK_AGENT_TOOL_NAMES.getContext,
  bilig_read_workbook: WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
  "bilig.read_workbook": WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
  bilig_read_range: WORKBOOK_AGENT_TOOL_NAMES.readRange,
  "bilig.read_range": WORKBOOK_AGENT_TOOL_NAMES.readRange,
  bilig_read_selection: WORKBOOK_AGENT_TOOL_NAMES.readSelection,
  "bilig.read_selection": WORKBOOK_AGENT_TOOL_NAMES.readSelection,
  bilig_read_visible_range: WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
  "bilig.read_visible_range": WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
  bilig_read_recent_changes: WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
  "bilig.read_recent_changes": WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
  bilig_start_workflow: WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
  "bilig.start_workflow": WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
  bilig_inspect_cell: WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
  "bilig.inspect_cell": WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
  bilig_find_formula_issues: WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
  "bilig.find_formula_issues": WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
  bilig_search_workbook: WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
  "bilig.search_workbook": WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
  bilig_trace_dependencies: WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
  "bilig.trace_dependencies": WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
  bilig_write_range: WORKBOOK_AGENT_TOOL_NAMES.writeRange,
  "bilig.write_range": WORKBOOK_AGENT_TOOL_NAMES.writeRange,
  bilig_clear_range: WORKBOOK_AGENT_TOOL_NAMES.clearRange,
  "bilig.clear_range": WORKBOOK_AGENT_TOOL_NAMES.clearRange,
  bilig_format_range: WORKBOOK_AGENT_TOOL_NAMES.formatRange,
  "bilig.format_range": WORKBOOK_AGENT_TOOL_NAMES.formatRange,
  bilig_fill_range: WORKBOOK_AGENT_TOOL_NAMES.fillRange,
  "bilig.fill_range": WORKBOOK_AGENT_TOOL_NAMES.fillRange,
  bilig_copy_range: WORKBOOK_AGENT_TOOL_NAMES.copyRange,
  "bilig.copy_range": WORKBOOK_AGENT_TOOL_NAMES.copyRange,
  bilig_move_range: WORKBOOK_AGENT_TOOL_NAMES.moveRange,
  "bilig.move_range": WORKBOOK_AGENT_TOOL_NAMES.moveRange,
  bilig_create_sheet: WORKBOOK_AGENT_TOOL_NAMES.createSheet,
  "bilig.create_sheet": WORKBOOK_AGENT_TOOL_NAMES.createSheet,
  bilig_rename_sheet: WORKBOOK_AGENT_TOOL_NAMES.renameSheet,
  "bilig.rename_sheet": WORKBOOK_AGENT_TOOL_NAMES.renameSheet,
  bilig_update_row_metadata: WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata,
  "bilig.update_row_metadata": WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata,
  bilig_update_column_metadata: WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata,
  "bilig.update_column_metadata": WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata,
};

const workbookAgentToolNameSet = new Set<string>(Object.values(WORKBOOK_AGENT_TOOL_NAMES));

export function normalizeWorkbookAgentToolName(toolName: string): string {
  return LEGACY_WORKBOOK_AGENT_TOOL_NAME_MAP[toolName] ?? toolName;
}

export function isWorkbookAgentToolName(toolName: string): toolName is WorkbookAgentToolName {
  return workbookAgentToolNameSet.has(normalizeWorkbookAgentToolName(toolName));
}
