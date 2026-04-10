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
  summarizeWorkbookStructure,
  traceWorkbookDependencies,
} from "./workbook-agent-comprehension.js";

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
    case "findFormulaIssues":
      return {
        title: "Find Formula Issues",
        runningSummary: "Running formula issue scan workflow.",
        stepPlans: [
          {
            stepId: "scan-formula-cells",
            label: "Scan formula cells",
            runningSummary:
              "Scanning workbook formulas for errors, cycles, and JS-only fallbacks.",
            pendingSummary:
              "Waiting to scan workbook formulas for errors, cycles, and JS-only fallbacks.",
          },
          {
            stepId: "draft-issue-report",
            label: "Draft issue report",
            runningSummary: "Drafting the durable formula issue report.",
            pendingSummary: "Waiting to assemble the durable formula issue report.",
          },
        ],
      };
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
            pendingSummary:
              "Waiting to assemble the durable current-cell explanation artifact.",
          },
        ],
      };
  }
}

export function describeWorkbookAgentWorkflowTemplate(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
): Pick<WorkbookAgentWorkflowTemplateMetadata, "title" | "runningSummary"> {
  const metadata = getWorkflowTemplateMetadata(workflowTemplate);
  return {
    title: metadata.title,
    runningSummary: metadata.runningSummary,
  };
}

export function createRunningWorkflowSteps(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  now: number,
): WorkbookAgentWorkflowStep[] {
  return getWorkflowTemplateMetadata(workflowTemplate).stepPlans.map((step, index) => ({
    stepId: step.stepId,
    label: step.label,
    status: index === 0 ? "running" : "pending",
    summary: index === 0 ? step.runningSummary : step.pendingSummary,
    updatedAtUnixMs: now,
  }));
}

export function completeWorkflowSteps(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  stepResults: readonly WorkbookAgentWorkflowStepResult[],
  now: number,
): WorkbookAgentWorkflowStep[] {
  const resultByStepId = new Map(stepResults.map((step) => [step.stepId, step]));
  return getWorkflowTemplateMetadata(workflowTemplate).stepPlans.map((step) => {
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
): WorkbookAgentWorkflowStep[] {
  let markedFailure = false;
  const runningStepId = runningSteps.find((step) => step.status === "running")?.stepId;
  return getWorkflowTemplateMetadata(workflowTemplate).stepPlans.map((step, index) => {
    const current = runningSteps.find((candidate) => candidate.stepId === step.stepId);
    if (!markedFailure && (step.stepId === runningStepId || (runningStepId === undefined && index === 0))) {
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
        current?.status === "completed" ? current.updatedAtUnixMs : current?.updatedAtUnixMs ?? now,
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
      const formulaIssues = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) =>
        findWorkbookFormulaIssues(runtime),
      );
      return {
        title: "Find Formula Issues",
        summary:
          formulaIssues.summary.issueCount === 0
            ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"} and found no issues.`
            : `Found ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? "" : "s"} across ${String(formulaIssues.summary.scannedFormulaCells)} scanned formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}.`,
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
                ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"} and found no issues.`
                : `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"} and found ${String(formulaIssues.summary.issueCount)} issue${formulaIssues.summary.issueCount === 1 ? "" : "s"}.`,
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
      const explanation = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
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
      });
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
  }
}
