import { WORKBOOK_AGENT_TOOL_NAMES } from "./workbook-agent-tool-names.js";

export type WorkbookAgentSkillFocus = "read" | "analyze" | "edit";

export interface WorkbookAgentSkillDescriptor {
  readonly id: string;
  readonly label: string;
  readonly focus: WorkbookAgentSkillFocus;
  readonly description: string;
  readonly prompt: string;
  readonly toolNames: readonly string[];
}

export const workbookAgentSkillDescriptors: readonly WorkbookAgentSkillDescriptor[] = [
  {
    id: "summarize-workbook",
    label: "Summarize Workbook",
    focus: "read",
    description: "Read the workbook structure and explain what each sheet appears to do.",
    prompt:
      "Summarize this workbook, explain what each sheet appears to do, and call out any obvious hotspots or risks. Prefer the built-in durable workflow when it matches.",
    toolNames: [WORKBOOK_AGENT_TOOL_NAMES.startWorkflow, WORKBOOK_AGENT_TOOL_NAMES.readWorkbook],
  },
  {
    id: "inspect-selection",
    label: "Inspect Selection",
    focus: "analyze",
    description: "Read the current cell, explain its value or formula, and trace direct links.",
    prompt:
      "Inspect the current cell selection, explain its value or formula, and trace its direct precedents and dependents.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.getContext,
      WORKBOOK_AGENT_TOOL_NAMES.readSelection,
      WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
    ],
  },
  {
    id: "find-formula-issues",
    label: "Find Formula Issues",
    focus: "analyze",
    description: "Scan the workbook for broken formulas, cycles, and JS-only fallback formulas.",
    prompt:
      "Scan this workbook for broken formulas, cycles, and JS-only fallback formulas, then summarize the highest-risk issues first. Prefer the built-in durable workflows when the user wants either a saved report or a staged highlight preview in the thread.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
      WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
    ],
  },
  {
    id: "search-workbook",
    label: "Search Workbook",
    focus: "analyze",
    description: "Search workbook structure, formulas, inputs, and visible values for a concept.",
    prompt:
      "Search this workbook for the concept I mention, use workbook search before broader explanation, and cite the strongest matching cells or sheets.",
    toolNames: [WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook, WORKBOOK_AGENT_TOOL_NAMES.inspectCell],
  },
  {
    id: "trace-dependencies",
    label: "Trace Dependencies",
    focus: "analyze",
    description: "Trace upstream and downstream workbook links from the current selection.",
    prompt:
      "Trace the dependency graph around the current selection for multiple hops, then explain the most important upstream and downstream cells. Prefer the built-in durable workflow when the user wants that trace saved in the thread.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      WORKBOOK_AGENT_TOOL_NAMES.getContext,
      WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
      WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
    ],
  },
  {
    id: "review-visible-range",
    label: "Review Visible Range",
    focus: "read",
    description: "Read the current viewport and summarize headers, patterns, and issues.",
    prompt:
      "Read the currently visible range and summarize its structure, headers, patterns, and any obvious issues.",
    toolNames: [WORKBOOK_AGENT_TOOL_NAMES.getContext, WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange],
  },
  {
    id: "describe-recent-changes",
    label: "Describe Recent Changes",
    focus: "read",
    description: "Read the latest durable workbook revisions and summarize what changed recently.",
    prompt:
      "Describe the most recent workbook changes, highlight the highest-impact revisions first, and cite the affected sheets or ranges when possible. Prefer the built-in durable workflow when it matches.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
    ],
  },
  {
    id: "edit-selection",
    label: "Edit Selection",
    focus: "edit",
    description: "Use the current selection to stage a spreadsheet-native preview bundle.",
    prompt:
      "Use the current selection context to stage the right spreadsheet edit as one coherent preview bundle.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.getContext,
      WORKBOOK_AGENT_TOOL_NAMES.readSelection,
      WORKBOOK_AGENT_TOOL_NAMES.writeRange,
      WORKBOOK_AGENT_TOOL_NAMES.formatRange,
    ],
  },
  {
    id: "reshape-sheet",
    label: "Reshape Sheet",
    focus: "edit",
    description: "Reorganize sheet structure with semantic copy, move, fill, and sheet tools.",
    prompt:
      "Restructure or clean up this sheet using semantic range and sheet tools. Prefer the built-in durable workflows for create-sheet, rename-current-sheet, row/column hide or unhide, and current-sheet header, number-format, or whitespace normalization when they match, otherwise stage one coherent preview bundle.",
    toolNames: [
      WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
      WORKBOOK_AGENT_TOOL_NAMES.readRange,
      WORKBOOK_AGENT_TOOL_NAMES.fillRange,
      WORKBOOK_AGENT_TOOL_NAMES.copyRange,
      WORKBOOK_AGENT_TOOL_NAMES.moveRange,
      WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata,
      WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata,
      WORKBOOK_AGENT_TOOL_NAMES.createSheet,
      WORKBOOK_AGENT_TOOL_NAMES.renameSheet,
    ],
  },
];

export function renderWorkbookAgentSkillInstructions(): string {
  return workbookAgentSkillDescriptors
    .map(
      (skill) =>
        `${skill.label} (${skill.focus}): ${skill.description} Tools: ${skill.toolNames.join(", ")}.`,
    )
    .join(" ");
}
