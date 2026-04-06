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
      "Summarize this workbook, explain what each sheet appears to do, and call out any obvious hotspots or risks.",
    toolNames: ["bilig.read_workbook"],
  },
  {
    id: "inspect-selection",
    label: "Inspect Selection",
    focus: "analyze",
    description: "Read the current cell, explain its value or formula, and trace direct links.",
    prompt:
      "Inspect the current cell selection, explain its value or formula, and trace its direct precedents and dependents.",
    toolNames: ["bilig.get_context", "bilig.read_selection", "bilig.inspect_cell"],
  },
  {
    id: "review-visible-range",
    label: "Review Visible Range",
    focus: "read",
    description: "Read the current viewport and summarize headers, patterns, and issues.",
    prompt:
      "Read the currently visible range and summarize its structure, headers, patterns, and any obvious issues.",
    toolNames: ["bilig.get_context", "bilig.read_visible_range"],
  },
  {
    id: "edit-selection",
    label: "Edit Selection",
    focus: "edit",
    description: "Use the current selection to stage a spreadsheet-native preview bundle.",
    prompt:
      "Use the current selection context to stage the right spreadsheet edit as one coherent preview bundle.",
    toolNames: [
      "bilig.get_context",
      "bilig.read_selection",
      "bilig.write_range",
      "bilig.format_range",
    ],
  },
  {
    id: "reshape-sheet",
    label: "Reshape Sheet",
    focus: "edit",
    description: "Reorganize sheet structure with semantic copy, move, fill, and sheet tools.",
    prompt:
      "Restructure this sheet using semantic range and sheet tools, then stage one coherent preview bundle.",
    toolNames: [
      "bilig.read_workbook",
      "bilig.read_range",
      "bilig.fill_range",
      "bilig.copy_range",
      "bilig.move_range",
      "bilig.create_sheet",
      "bilig.rename_sheet",
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
