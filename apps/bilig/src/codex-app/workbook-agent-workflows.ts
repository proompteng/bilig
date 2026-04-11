import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowRun,
  WorkbookAgentWorkflowStep,
  WorkbookAgentWorkflowTemplate,
} from "@bilig/contracts";
import { ValueTag, formatErrorCode } from "@bilig/protocol";
import type { ZeroSyncService } from "../zero/service.js";
import {
  findWorkbookFormulaIssues,
  searchWorkbook,
  type WorkbookSearchReport,
  summarizeWorkbookStructure,
  traceWorkbookDependencies,
} from "./workbook-agent-comprehension.js";

export interface WorkbookAgentWorkflowExecutionInput {
  readonly query?: string;
  readonly sheetName?: string;
  readonly limit?: number;
}

interface WorkbookAgentWorkflowStepPlan {
  readonly stepId: string;
  readonly label: string;
  readonly runningSummary: string;
  readonly pendingSummary: string;
}

interface WorkbookAgentWorkflowStepResult {
  readonly stepId: string;
  readonly label: string;
  readonly summary: string;
}

export interface WorkbookAgentWorkflowExecutionResult {
  readonly title: string;
  readonly summary: string;
  readonly artifact: WorkbookAgentWorkflowArtifact;
  readonly citations: readonly WorkbookAgentTimelineCitation[];
  readonly steps: readonly WorkbookAgentWorkflowStepResult[];
}

interface WorkbookAgentWorkflowTemplateMetadata {
  readonly title: string;
  readonly runningSummary: string;
  readonly stepPlans: readonly WorkbookAgentWorkflowStepPlan[];
}

function getWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null,
): WorkbookAgentWorkflowTemplateMetadata {
  switch (workflowTemplate) {
    case "summarizeWorkbook":
      return {
        title: "Summarize Workbook",
        runningSummary: "Running workbook summary workflow.",
        stepPlans: [
          {
            stepId: "inspect-workbook",
            label: "Inspect workbook structure",
            runningSummary: "Reading durable workbook structure and layout metadata.",
            pendingSummary: "Waiting to inspect workbook structure and layout metadata.",
          },
          {
            stepId: "draft-summary",
            label: "Draft summary artifact",
            runningSummary: "Drafting the durable workbook summary artifact.",
            pendingSummary: "Waiting to assemble the durable workbook summary artifact.",
          },
        ],
      };
    case "summarizeCurrentSheet":
      return {
        title: "Summarize Current Sheet",
        runningSummary: "Running current sheet summary workflow.",
        stepPlans: [
          {
            stepId: "inspect-current-sheet",
            label: "Inspect current sheet",
            runningSummary: "Reading durable metadata for the active sheet.",
            pendingSummary: "Waiting to inspect durable metadata for the active sheet.",
          },
          {
            stepId: "draft-sheet-summary",
            label: "Draft current sheet summary",
            runningSummary: "Drafting the durable current-sheet summary artifact.",
            pendingSummary: "Waiting to assemble the durable current-sheet summary artifact.",
          },
        ],
      };
    case "describeRecentChanges":
      return {
        title: "Describe Recent Changes",
        runningSummary: "Running recent change report workflow.",
        stepPlans: [
          {
            stepId: "load-revisions",
            label: "Load durable revisions",
            runningSummary: "Reading the latest durable workbook revisions.",
            pendingSummary: "Waiting to read the latest durable workbook revisions.",
          },
          {
            stepId: "draft-change-report",
            label: "Draft change report",
            runningSummary: "Drafting the durable recent change report.",
            pendingSummary: "Waiting to assemble the durable recent change report.",
          },
        ],
      };
    case "findFormulaIssues": {
      const scopeLabel = workflowInput?.sheetName ?? "the workbook";
      return {
        title: "Find Formula Issues",
        runningSummary: `Running formula issue scan workflow for ${scopeLabel}.`,
        stepPlans: [
          {
            stepId: "scan-formula-cells",
            label: "Scan formula cells",
            runningSummary: workflowInput?.sheetName
              ? `Scanning ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
              : "Scanning workbook formulas for errors, cycles, and JS-only fallbacks.",
            pendingSummary: workflowInput?.sheetName
              ? `Waiting to scan ${workflowInput.sheetName} formulas for errors, cycles, and JS-only fallbacks.`
              : "Waiting to scan workbook formulas for errors, cycles, and JS-only fallbacks.",
          },
          {
            stepId: "draft-issue-report",
            label: "Draft issue report",
            runningSummary: "Drafting the durable formula issue report.",
            pendingSummary: "Waiting to assemble the durable formula issue report.",
          },
        ],
      };
    }
    case "traceSelectionDependencies":
      return {
        title: "Trace Selection Dependencies",
        runningSummary: "Running dependency trace workflow.",
        stepPlans: [
          {
            stepId: "inspect-selection",
            label: "Inspect current selection",
            runningSummary: "Reading the current workbook selection context.",
            pendingSummary: "Waiting to read the current workbook selection context.",
          },
          {
            stepId: "trace-links",
            label: "Trace workbook links",
            runningSummary: "Tracing direct precedents and dependents from the selection.",
            pendingSummary: "Waiting to trace direct precedents and dependents from the selection.",
          },
          {
            stepId: "draft-trace-report",
            label: "Draft trace report",
            runningSummary: "Drafting the durable dependency trace report.",
            pendingSummary: "Waiting to assemble the durable dependency trace report.",
          },
        ],
      };
    case "explainSelectionCell":
      return {
        title: "Explain Current Cell",
        runningSummary: "Running current cell explanation workflow.",
        stepPlans: [
          {
            stepId: "inspect-selection",
            label: "Inspect current selection",
            runningSummary: "Reading the current workbook selection context.",
            pendingSummary: "Waiting to read the current workbook selection context.",
          },
          {
            stepId: "explain-cell",
            label: "Explain current cell",
            runningSummary: "Loading the selected cell value, formula state, and workbook links.",
            pendingSummary:
              "Waiting to load the selected cell value, formula state, and workbook links.",
          },
          {
            stepId: "draft-explanation",
            label: "Draft explanation artifact",
            runningSummary: "Drafting the durable current-cell explanation artifact.",
            pendingSummary: "Waiting to assemble the durable current-cell explanation artifact.",
          },
        ],
      };
    case "searchWorkbookQuery": {
      const searchLabel =
        typeof workflowInput?.query === "string" && workflowInput.query.trim().length > 0
          ? `"${workflowInput.query.trim()}"`
          : "the requested query";
      return {
        title: "Search Workbook",
        runningSummary: `Searching the workbook for ${searchLabel}.`,
        stepPlans: [
          {
            stepId: "search-workbook",
            label: "Search workbook",
            runningSummary: `Searching workbook sheets, formulas, values, and addresses for ${searchLabel}.`,
            pendingSummary: `Waiting to search workbook sheets, formulas, values, and addresses for ${searchLabel}.`,
          },
          {
            stepId: "draft-search-report",
            label: "Draft search report",
            runningSummary: "Drafting the durable workbook search report.",
            pendingSummary: "Waiting to assemble the durable workbook search report.",
          },
        ],
      };
    }
  }
}

export function describeWorkbookAgentWorkflowTemplate(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null,
): Pick<WorkbookAgentWorkflowTemplateMetadata, "title" | "runningSummary"> {
  const metadata = getWorkflowTemplateMetadata(workflowTemplate, workflowInput);
  return {
    title: metadata.title,
    runningSummary: metadata.runningSummary,
  };
}

export function createRunningWorkflowSteps(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  now: number,
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null,
): WorkbookAgentWorkflowStep[] {
  return getWorkflowTemplateMetadata(workflowTemplate, workflowInput).stepPlans.map(
    (step, index) => ({
      stepId: step.stepId,
      label: step.label,
      status: index === 0 ? "running" : "pending",
      summary: index === 0 ? step.runningSummary : step.pendingSummary,
      updatedAtUnixMs: now,
    }),
  );
}

export function completeWorkflowSteps(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  stepResults: readonly WorkbookAgentWorkflowStepResult[],
  now: number,
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null,
): WorkbookAgentWorkflowStep[] {
  const resultByStepId = new Map(stepResults.map((step) => [step.stepId, step]));
  return getWorkflowTemplateMetadata(workflowTemplate, workflowInput).stepPlans.map((step) => {
    const result = resultByStepId.get(step.stepId);
    return {
      stepId: step.stepId,
      label: step.label,
      status: "completed",
      summary: result?.summary ?? step.runningSummary,
      updatedAtUnixMs: now,
    };
  });
}

export function failWorkflowSteps(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  runningSteps: readonly WorkbookAgentWorkflowStep[],
  errorMessage: string,
  now: number,
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null,
): WorkbookAgentWorkflowStep[] {
  let markedFailure = false;
  const runningStepId = runningSteps.find((step) => step.status === "running")?.stepId;
  return getWorkflowTemplateMetadata(workflowTemplate, workflowInput).stepPlans.map(
    (step, index) => {
      const current = runningSteps.find((candidate) => candidate.stepId === step.stepId);
      if (
        !markedFailure &&
        (step.stepId === runningStepId || (runningStepId === undefined && index === 0))
      ) {
        markedFailure = true;
        return {
          stepId: step.stepId,
          label: step.label,
          status: "failed",
          summary: errorMessage,
          updatedAtUnixMs: now,
        };
      }
      return {
        stepId: step.stepId,
        label: step.label,
        status: current?.status === "completed" ? "completed" : "pending",
        summary: current?.summary ?? step.pendingSummary,
        updatedAtUnixMs:
          current?.status === "completed"
            ? current.updatedAtUnixMs
            : (current?.updatedAtUnixMs ?? now),
      };
    },
  );
}

export function cancelWorkflowSteps(
  runningSteps: readonly WorkbookAgentWorkflowStep[],
  now: number,
): WorkbookAgentWorkflowStep[] {
  return runningSteps.map((step) => {
    if (step.status === "completed") {
      return step;
    }
    return {
      ...step,
      status: "cancelled",
      summary:
        step.status === "running"
          ? "Workflow cancelled before this step completed."
          : "Workflow cancelled before this step started.",
      updatedAtUnixMs: now,
    };
  });
}

function summarizeWorkbookMarkdown(summary: ReturnType<typeof summarizeWorkbookStructure>): string {
  const lines = [
    "## Workbook Summary",
    "",
    `Sheets: ${String(summary.summary.sheetCount)}`,
    `Formula cells: ${String(summary.summary.totalFormulaCellCount)}`,
    `Tables: ${String(summary.summary.tableCount)}`,
    `Pivots: ${String(summary.summary.pivotCount)}`,
    `Spills: ${String(summary.summary.spillCount)}`,
    "",
    "### Sheets",
  ];
  summary.sheets.forEach((sheet) => {
    lines.push(
      `- ${sheet.name}: ${String(sheet.cellCount)} populated cells, ${String(sheet.formulaCellCount)} formulas${sheet.usedRange ? `, used range ${sheet.usedRange.startAddress}:${sheet.usedRange.endAddress}` : ""}`,
    );
  });
  return lines.join("\n");
}

function summarizeCurrentSheetMarkdown(
  sheet: ReturnType<typeof summarizeWorkbookStructure>["sheets"][number],
): string {
  const lines = [
    "## Current Sheet Summary",
    "",
    `Sheet: ${sheet.name}`,
    `Order: ${String(sheet.order)}`,
    `Used range: ${sheet.usedRange ? `${sheet.usedRange.startAddress}:${sheet.usedRange.endAddress}` : "(empty)"}`,
    `Populated cells: ${String(sheet.cellCount)}`,
    `Formula cells: ${String(sheet.formulaCellCount)}`,
    `Tables: ${String(sheet.tableCount)}`,
    `Pivots: ${String(sheet.pivotCount)}`,
    `Spills: ${String(sheet.spillCount)}`,
    `Filters: ${String(sheet.filterCount)}`,
    `Sorts: ${String(sheet.sortCount)}`,
    `Freeze panes: ${sheet.freezePane ? `${String(sheet.freezePane.rows)} row(s), ${String(sheet.freezePane.cols)} column(s)` : "none"}`,
    `Hidden row indexes: ${String(sheet.rowMetadata.hiddenIndexCount)}`,
    `Hidden column indexes: ${String(sheet.columnMetadata.hiddenIndexCount)}`,
    `Explicit row sizes: ${String(sheet.rowMetadata.explicitSizeIndexCount)}`,
    `Explicit column sizes: ${String(sheet.columnMetadata.explicitSizeIndexCount)}`,
    "",
    "### Tables",
  ];
  if (sheet.tables.length === 0) {
    lines.push("- None");
  } else {
    for (const table of sheet.tables) {
      lines.push(
        `- ${table.name}: ${table.startAddress}:${table.endAddress} (${String(table.columnCount)} columns)`,
      );
    }
  }
  lines.push("", "### Pivots");
  if (sheet.pivots.length === 0) {
    lines.push("- None");
  } else {
    for (const pivot of sheet.pivots) {
      lines.push(
        `- ${pivot.name}: ${pivot.address} from ${pivot.source} (${String(pivot.valueCount)} values)`,
      );
    }
  }
  lines.push("", "### Spill Ranges");
  if (sheet.spills.length === 0) {
    lines.push("- None");
  } else {
    for (const spill of sheet.spills) {
      lines.push(`- ${spill.address}: ${String(spill.rows)}x${String(spill.cols)}`);
    }
  }
  return lines.join("\n");
}

function summarizeRecentChangesMarkdown(
  changes: Awaited<ReturnType<ZeroSyncService["listWorkbookChanges"]>>,
): string {
  const lines = ["## Recent Changes", ""];
  if (changes.length === 0) {
    lines.push("No durable workbook changes are available yet.");
    return lines.join("\n");
  }
  changes.forEach((record) => {
    const location =
      record.range?.sheetName && record.range?.startAddress && record.range?.endAddress
        ? ` ${record.range.sheetName}!${record.range.startAddress}:${record.range.endAddress}`
        : record.sheetName && record.anchorAddress
          ? ` ${record.sheetName}!${record.anchorAddress}`
          : "";
    lines.push(`- r${String(record.revision)}: ${record.summary}${location}`);
  });
  return lines.join("\n");
}

function summarizeFormulaIssueKinds(
  issueKinds: readonly ("error" | "cycle" | "unsupported")[],
): string {
  return issueKinds.join(", ");
}

function summarizeFormulaIssuesMarkdown(
  report: ReturnType<typeof findWorkbookFormulaIssues>,
): string {
  const lines = [
    "## Formula Issues",
    "",
    `Scanned formula cells: ${String(report.summary.scannedFormulaCells)}`,
    `Issues found: ${String(report.summary.issueCount)}`,
    `Errors: ${String(report.summary.errorCount)}`,
    `Cycles: ${String(report.summary.cycleCount)}`,
    `JS-only fallbacks: ${String(report.summary.unsupportedCount)}`,
  ];
  if (report.summary.truncated) {
    lines.push("Showing the highest-risk issues from the requested limit.");
  }
  lines.push("");
  if (report.issues.length === 0) {
    lines.push("No formula issues were detected in the current workbook.");
    return lines.join("\n");
  }
  lines.push("### Highest-Risk Issues");
  report.issues.forEach((issue) => {
    const valueSuffix = issue.valueText.length > 0 ? ` -> ${issue.valueText}` : "";
    const errorSuffix = issue.errorText ? ` (${issue.errorText})` : "";
    lines.push(
      `- ${issue.sheetName}!${issue.address}: =${issue.formula}${valueSuffix} [${summarizeFormulaIssueKinds(issue.issueKinds)}]${errorSuffix}`,
    );
  });
  return lines.join("\n");
}

function summarizeDependencyTraceMarkdown(
  report: ReturnType<typeof traceWorkbookDependencies>,
): string {
  const lines = [
    "## Dependency Trace",
    "",
    `Root: ${report.root.sheetName}!${report.root.address}`,
    `Direction: ${report.direction}`,
    `Depth: ${String(report.depth)}`,
    `Direct precedents discovered: ${String(report.summary.precedentCount)}`,
    `Direct dependents discovered: ${String(report.summary.dependentCount)}`,
  ];
  if (report.summary.truncated) {
    lines.push("Trace output was truncated to stay inside the workflow node budget.");
  }
  lines.push("");
  if (report.layers.length === 0) {
    lines.push("No workbook precedents or dependents were found from the current selection.");
    return lines.join("\n");
  }
  report.layers.forEach((layer) => {
    lines.push(`### Depth ${String(layer.depth)}`);
    if (layer.precedents.length > 0) {
      lines.push("Precedents:");
      layer.precedents.forEach((node) => {
        lines.push(`- ${node.sheetName}!${node.address}: ${node.valueText}`);
      });
    }
    if (layer.dependents.length > 0) {
      lines.push("Dependents:");
      layer.dependents.forEach((node) => {
        lines.push(`- ${node.sheetName}!${node.address}: ${node.valueText}`);
      });
    }
    lines.push("");
  });
  return lines.join("\n").trimEnd();
}

function serializeWorkflowCellValue(value: {
  tag: ValueTag;
  value?: number | boolean | string;
  code?: number;
}): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "(empty)";
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return String(value.value ?? "");
    case ValueTag.Error:
      return typeof value.code === "number" ? formatErrorCode(value.code) : "#ERROR!";
    default:
      return "(empty)";
  }
}

function summarizeCellExplanationMarkdown(explanation: {
  readonly sheetName: string;
  readonly address: string;
  readonly valueText: string;
  readonly formula: string | null;
  readonly format: string | null;
  readonly version: number;
  readonly inCycle: boolean;
  readonly mode: string | null;
  readonly topoRank: number | null;
  readonly directPrecedents: readonly string[];
  readonly directDependents: readonly string[];
}): string {
  const lines = [
    "## Current Cell",
    "",
    `Cell: ${explanation.sheetName}!${explanation.address}`,
    `Value: ${explanation.valueText}`,
    `Formula: ${explanation.formula ?? "(none)"}`,
    `Calculation mode: ${explanation.mode ?? "(unknown)"}`,
    `Version: ${String(explanation.version)}`,
    `In cycle: ${explanation.inCycle ? "yes" : "no"}`,
    `Direct precedents: ${String(explanation.directPrecedents.length)}`,
    `Direct dependents: ${String(explanation.directDependents.length)}`,
  ];
  if (explanation.topoRank !== null) {
    lines.push(`Topological rank: ${String(explanation.topoRank)}`);
  }
  if (explanation.format) {
    lines.push(`Number format: ${explanation.format}`);
  }
  lines.push("", "### Direct precedents");
  if (explanation.directPrecedents.length === 0) {
    lines.push("- None");
  } else {
    for (const precedent of explanation.directPrecedents) {
      lines.push(`- ${precedent}`);
    }
  }
  lines.push("", "### Direct dependents");
  if (explanation.directDependents.length === 0) {
    lines.push("- None");
  } else {
    for (const dependent of explanation.directDependents) {
      lines.push(`- ${dependent}`);
    }
  }
  return lines.join("\n");
}

function summarizeSearchResultsMarkdown(report: WorkbookSearchReport): string {
  const lines = [
    "## Workbook Search",
    "",
    `Query: ${report.query}`,
    `Matches: ${String(report.summary.matchCount)}`,
  ];
  if (report.summary.truncated) {
    lines.push("Results were truncated to stay inside the workflow result budget.");
  }
  lines.push("");
  if (report.matches.length === 0) {
    lines.push("No workbook matches were found for the requested query.");
    return lines.join("\n");
  }
  lines.push("### Top matches");
  for (const match of report.matches) {
    if (match.kind === "sheet") {
      lines.push(`- Sheet ${match.sheetName} [${match.reasons.join(", ")}]`);
      continue;
    }
    const location = `${match.sheetName}!${match.address ?? "?"}`;
    const snippet =
      match.formula ?? match.valueText ?? match.inputText ?? match.snippet ?? "(no snippet)";
    lines.push(`- ${location}: ${snippet} [${match.reasons.join(", ")}]`);
  }
  return lines.join("\n");
}

export function createWorkflowRunRecord(input: {
  runId: string;
  threadId: string;
  startedByUserId: string;
  workflowTemplate: WorkbookAgentWorkflowTemplate;
  title: string;
  summary: string;
  status: WorkbookAgentWorkflowRun["status"];
  now: number;
  artifact?: WorkbookAgentWorkflowArtifact | null;
  steps?: readonly WorkbookAgentWorkflowStep[] | null;
  completedAtUnixMs?: number | null;
  errorMessage?: string | null;
}): WorkbookAgentWorkflowRun {
  return {
    runId: input.runId,
    threadId: input.threadId,
    startedByUserId: input.startedByUserId,
    workflowTemplate: input.workflowTemplate,
    title: input.title,
    summary: input.summary,
    status: input.status,
    createdAtUnixMs: input.now,
    updatedAtUnixMs: input.now,
    completedAtUnixMs: input.completedAtUnixMs ?? null,
    errorMessage: input.errorMessage ?? null,
    steps: input.steps ? [...input.steps] : [],
    artifact: input.artifact ?? null,
  };
}

export async function executeWorkbookAgentWorkflow(input: {
  documentId: string;
  zeroSyncService: ZeroSyncService;
  workflowTemplate: WorkbookAgentWorkflowTemplate;
  context?: WorkbookAgentUiContext | null;
  workflowInput?: WorkbookAgentWorkflowExecutionInput | null;
}): Promise<WorkbookAgentWorkflowExecutionResult> {
  switch (input.workflowTemplate) {
    case "summarizeWorkbook": {
      const structure = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) =>
        summarizeWorkbookStructure(runtime),
      );
      return {
        title: "Summarize Workbook",
        summary: `Summarized workbook structure across ${String(structure.summary.sheetCount)} sheet${structure.summary.sheetCount === 1 ? "" : "s"}.`,
        artifact: {
          kind: "markdown",
          title: "Workbook Summary",
          text: summarizeWorkbookMarkdown(structure),
        },
        citations: [],
        steps: [
          {
            stepId: "inspect-workbook",
            label: "Inspect workbook structure",
            summary: `Read durable workbook structure across ${String(structure.summary.sheetCount)} sheet${structure.summary.sheetCount === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-summary",
            label: "Draft summary artifact",
            summary: "Prepared the durable workbook summary artifact for the thread.",
          },
        ],
      };
    }
    case "summarizeCurrentSheet": {
      const selection = input.context?.selection;
      if (!selection) {
        throw new Error("Selection context is required for current sheet summary workflows.");
      }
      const structure = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) =>
        summarizeWorkbookStructure(runtime),
      );
      const sheet = structure.sheets.find((candidate) => candidate.name === selection.sheetName);
      if (!sheet) {
        throw new Error(`Active sheet ${selection.sheetName} was not found in the workbook.`);
      }
      return {
        title: "Summarize Current Sheet",
        summary: `Summarized ${sheet.name} with ${String(sheet.cellCount)} populated cell${sheet.cellCount === 1 ? "" : "s"} and ${String(sheet.tableCount)} table${sheet.tableCount === 1 ? "" : "s"}.`,
        artifact: {
          kind: "markdown",
          title: "Current Sheet Summary",
          text: summarizeCurrentSheetMarkdown(sheet),
        },
        citations: [
          ...(sheet.usedRange
            ? [
                {
                  kind: "range" as const,
                  sheetName: sheet.name,
                  startAddress: sheet.usedRange.startAddress,
                  endAddress: sheet.usedRange.endAddress,
                  role: "target" as const,
                },
              ]
            : []),
          ...sheet.tables.map((table) => ({
            kind: "range" as const,
            sheetName: sheet.name,
            startAddress: table.startAddress,
            endAddress: table.endAddress,
            role: "source" as const,
          })),
          ...sheet.pivots.map((pivot) => ({
            kind: "range" as const,
            sheetName: sheet.name,
            startAddress: pivot.address,
            endAddress: pivot.address,
            role: "source" as const,
          })),
          ...sheet.spills.map((spill) => ({
            kind: "range" as const,
            sheetName: sheet.name,
            startAddress: spill.address,
            endAddress: spill.address,
            role: "source" as const,
          })),
        ],
        steps: [
          {
            stepId: "inspect-current-sheet",
            label: "Inspect current sheet",
            summary: `Read durable metadata for ${sheet.name}, including used range, tables, pivots, spills, and axis metadata.`,
          },
          {
            stepId: "draft-sheet-summary",
            label: "Draft current sheet summary",
            summary: "Prepared the durable current-sheet summary artifact for the thread.",
          },
        ],
      };
    }
    case "describeRecentChanges": {
      const changes = await input.zeroSyncService.listWorkbookChanges(input.documentId, 10);
      return {
        title: "Describe Recent Changes",
        summary:
          changes.length === 0
            ? "No durable workbook changes were available to summarize."
            : `Summarized ${String(changes.length)} recent workbook change${changes.length === 1 ? "" : "s"}.`,
        artifact: {
          kind: "markdown",
          title: "Recent Changes",
          text: summarizeRecentChangesMarkdown(changes),
        },
        citations: changes.map((record) => ({
          kind: "revision",
          revision: record.revision,
        })),
        steps: [
          {
            stepId: "load-revisions",
            label: "Load durable revisions",
            summary:
              changes.length === 0
                ? "Loaded the durable revision log and found no workbook changes yet."
                : `Loaded ${String(changes.length)} durable workbook revision${changes.length === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-change-report",
            label: "Draft change report",
            summary: "Prepared the durable recent change report for the thread.",
          },
        ],
      };
    }
    case "findFormulaIssues": {
      const formulaIssues = await input.zeroSyncService.inspectWorkbook(
        input.documentId,
        (runtime) =>
          findWorkbookFormulaIssues(runtime, {
            ...(input.workflowInput?.sheetName ? { sheetName: input.workflowInput.sheetName } : {}),
            ...(input.workflowInput?.limit !== undefined
              ? { limit: input.workflowInput.limit }
              : {}),
          }),
      );
      const scopeLabel = input.workflowInput?.sheetName
        ? ` on ${input.workflowInput.sheetName}`
        : "";
      return {
        title: "Find Formula Issues",
        summary:
          formulaIssues.summary.issueCount === 0
            ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}${scopeLabel} and found no issues.`
            : `Found ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? "" : "s"}${scopeLabel} across ${String(formulaIssues.summary.scannedFormulaCells)} scanned formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}.`,
        artifact: {
          kind: "markdown",
          title: "Formula Issues",
          text: summarizeFormulaIssuesMarkdown(formulaIssues),
        },
        citations: formulaIssues.issues.map((issue) => ({
          kind: "range",
          sheetName: issue.sheetName,
          startAddress: issue.address,
          endAddress: issue.address,
          role: "source",
        })),
        steps: [
          {
            stepId: "scan-formula-cells",
            label: "Scan formula cells",
            summary:
              formulaIssues.summary.issueCount === 0
                ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}${scopeLabel} and found no issues.`
                : `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}${scopeLabel} and found ${String(formulaIssues.summary.issueCount)} issue${formulaIssues.summary.issueCount === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-issue-report",
            label: "Draft issue report",
            summary: "Prepared the durable formula issue report for the thread.",
          },
        ],
      };
    }
    case "traceSelectionDependencies": {
      const selection = input.context?.selection;
      if (!selection) {
        throw new Error("Selection context is required for dependency trace workflows.");
      }
      const dependencyTrace = await input.zeroSyncService.inspectWorkbook(
        input.documentId,
        (runtime) =>
          traceWorkbookDependencies(runtime, {
            sheetName: selection.sheetName,
            address: selection.address,
          }),
      );
      const citedNodes = [
        dependencyTrace.root,
        ...dependencyTrace.layers.flatMap((layer) => [...layer.precedents, ...layer.dependents]),
      ];
      return {
        title: "Trace Selection Dependencies",
        summary:
          dependencyTrace.summary.precedentCount === 0 &&
          dependencyTrace.summary.dependentCount === 0
            ? `No workbook precedents or dependents were found from ${selection.sheetName}!${selection.address}.`
            : `Traced ${String(dependencyTrace.summary.precedentCount)} precedent${dependencyTrace.summary.precedentCount === 1 ? "" : "s"} and ${String(dependencyTrace.summary.dependentCount)} dependent${dependencyTrace.summary.dependentCount === 1 ? "" : "s"} from ${selection.sheetName}!${selection.address}.`,
        artifact: {
          kind: "markdown",
          title: "Dependency Trace",
          text: summarizeDependencyTraceMarkdown(dependencyTrace),
        },
        citations: citedNodes.map((node) => ({
          kind: "range",
          sheetName: node.sheetName,
          startAddress: node.address,
          endAddress: node.address,
          role: "source",
        })),
        steps: [
          {
            stepId: "inspect-selection",
            label: "Inspect current selection",
            summary: `Loaded workbook context for ${selection.sheetName}!${selection.address}.`,
          },
          {
            stepId: "trace-links",
            label: "Trace workbook links",
            summary:
              dependencyTrace.summary.precedentCount === 0 &&
              dependencyTrace.summary.dependentCount === 0
                ? "No workbook precedents or dependents were discovered from the current selection."
                : `Traced ${String(dependencyTrace.summary.precedentCount)} precedent${dependencyTrace.summary.precedentCount === 1 ? "" : "s"} and ${String(dependencyTrace.summary.dependentCount)} dependent${dependencyTrace.summary.dependentCount === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-trace-report",
            label: "Draft trace report",
            summary: "Prepared the durable dependency trace report for the thread.",
          },
        ],
      };
    }
    case "explainSelectionCell": {
      const selection = input.context?.selection;
      if (!selection) {
        throw new Error("Selection context is required for current cell explanation workflows.");
      }
      const explanation = await input.zeroSyncService.inspectWorkbook(
        input.documentId,
        (runtime) => {
          const cell = runtime.engine.explainCell(selection.sheetName, selection.address);
          return {
            sheetName: cell.sheetName,
            address: cell.address,
            valueText: serializeWorkflowCellValue(cell.value),
            formula: cell.formula !== undefined ? `=${cell.formula}` : null,
            format: cell.format ?? null,
            version: cell.version,
            inCycle: cell.inCycle,
            mode: cell.mode ? String(cell.mode) : null,
            topoRank: cell.topoRank ?? null,
            directPrecedents: [...cell.directPrecedents],
            directDependents: [...cell.directDependents],
          };
        },
      );
      return {
        title: "Explain Current Cell",
        summary: `Explained ${selection.sheetName}!${selection.address}, including direct precedents and dependents.`,
        artifact: {
          kind: "markdown",
          title: "Current Cell",
          text: summarizeCellExplanationMarkdown(explanation),
        },
        citations: [
          {
            kind: "range",
            sheetName: explanation.sheetName,
            startAddress: explanation.address,
            endAddress: explanation.address,
            role: "target",
          },
          ...explanation.directPrecedents.map((address) => ({
            kind: "range" as const,
            sheetName: explanation.sheetName,
            startAddress: address,
            endAddress: address,
            role: "source" as const,
          })),
          ...explanation.directDependents.map((address) => ({
            kind: "range" as const,
            sheetName: explanation.sheetName,
            startAddress: address,
            endAddress: address,
            role: "source" as const,
          })),
        ],
        steps: [
          {
            stepId: "inspect-selection",
            label: "Inspect current selection",
            summary: `Loaded workbook context for ${selection.sheetName}!${selection.address}.`,
          },
          {
            stepId: "explain-cell",
            label: "Explain current cell",
            summary:
              explanation.formula === null
                ? `Read the current value and workbook links for ${selection.sheetName}!${selection.address}.`
                : `Read the current value, formula, and workbook links for ${selection.sheetName}!${selection.address}.`,
          },
          {
            stepId: "draft-explanation",
            label: "Draft explanation artifact",
            summary: "Prepared the durable current-cell explanation artifact for the thread.",
          },
        ],
      };
    }
    case "searchWorkbookQuery": {
      const query = input.workflowInput?.query?.trim();
      if (!query) {
        throw new Error("A query is required for workbook search workflows.");
      }
      const searchReport = await input.zeroSyncService.inspectWorkbook(
        input.documentId,
        (runtime) =>
          searchWorkbook(runtime, {
            query,
            ...(input.workflowInput?.sheetName ? { sheetName: input.workflowInput.sheetName } : {}),
            ...(input.workflowInput?.limit !== undefined
              ? { limit: input.workflowInput.limit }
              : {}),
          }),
      );
      return {
        title: "Search Workbook",
        summary:
          searchReport.summary.matchCount === 0
            ? `Found no workbook matches for "${query}".`
            : `Found ${String(searchReport.summary.matchCount)} workbook match${searchReport.summary.matchCount === 1 ? "" : "es"} for "${query}".`,
        artifact: {
          kind: "markdown",
          title: "Workbook Search",
          text: summarizeSearchResultsMarkdown(searchReport),
        },
        citations: searchReport.matches.flatMap((match) =>
          match.kind === "cell" && match.address
            ? [
                {
                  kind: "range" as const,
                  sheetName: match.sheetName,
                  startAddress: match.address,
                  endAddress: match.address,
                  role: "source" as const,
                },
              ]
            : [],
        ),
        steps: [
          {
            stepId: "search-workbook",
            label: "Search workbook",
            summary:
              searchReport.summary.matchCount === 0
                ? `Searched workbook sheets, formulas, values, and addresses for "${query}" and found no matches.`
                : `Searched workbook sheets, formulas, values, and addresses for "${query}" and found ${String(searchReport.summary.matchCount)} match${searchReport.summary.matchCount === 1 ? "" : "es"}.`,
          },
          {
            stepId: "draft-search-report",
            label: "Draft search report",
            summary: "Prepared the durable workbook search report for the thread.",
          },
        ],
      };
    }
  }
}
