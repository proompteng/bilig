import type { WorkbookAgentCommand } from "@bilig/agent-api";
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import { findWorkbookFormulaIssues } from "./workbook-agent-comprehension.js";
import { throwIfWorkflowCancelled } from "./workbook-agent-workflow-abort.js";

export interface FormulaWorkflowExecutionInput {
  readonly sheetName?: string;
  readonly limit?: number;
}

interface FormulaWorkflowStepPlan {
  readonly stepId: string;
  readonly label: string;
  readonly runningSummary: string;
  readonly pendingSummary: string;
}

interface FormulaWorkflowStepResult {
  readonly stepId: string;
  readonly label: string;
  readonly summary: string;
}

interface FormulaWorkflowTemplateMetadata {
  readonly title: string;
  readonly runningSummary: string;
  readonly stepPlans: readonly FormulaWorkflowStepPlan[];
}

export interface FormulaWorkflowExecutionResult {
  readonly title: string;
  readonly summary: string;
  readonly artifact: WorkbookAgentWorkflowArtifact;
  readonly citations: readonly WorkbookAgentTimelineCitation[];
  readonly steps: readonly FormulaWorkflowStepResult[];
  readonly commands?: readonly WorkbookAgentCommand[];
  readonly goalText?: string;
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

function summarizeHighlightedFormulaIssuesMarkdown(
  report: ReturnType<typeof findWorkbookFormulaIssues>,
): string {
  const lines = [
    "## Highlighted Formula Issues",
    "",
    `Scanned formula cells: ${String(report.summary.scannedFormulaCells)}`,
    `Issues highlighted: ${String(report.issues.length)}`,
  ];
  if (report.issues.length === 0) {
    lines.push("", "No formula issues were detected, so no preview bundle was staged.");
    return lines.join("\n");
  }
  lines.push("", "### Highlighted Cells");
  report.issues.forEach((issue) => {
    lines.push(
      `- ${issue.sheetName}!${issue.address} [${summarizeFormulaIssueKinds(issue.issueKinds)}]`,
    );
  });
  lines.push(
    "",
    "The staged preview bundle applies visible formatting to the listed cells so reviewers can inspect and apply the change authoritatively from the rail.",
  );
  return lines.join("\n");
}

function createIssueCitations(
  report: ReturnType<typeof findWorkbookFormulaIssues>,
  role: "source" | "target",
): WorkbookAgentTimelineCitation[] {
  return report.issues.map((issue) => ({
    kind: "range",
    sheetName: issue.sheetName,
    startAddress: issue.address,
    endAddress: issue.address,
    role,
  }));
}

export function getFormulaWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate,
  workflowInput?: FormulaWorkflowExecutionInput | null,
): FormulaWorkflowTemplateMetadata | null {
  const scopeLabel = workflowInput?.sheetName ?? "the workbook";
  if (workflowTemplate === "findFormulaIssues") {
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
  if (workflowTemplate === "highlightFormulaIssues") {
    return {
      title: "Highlight Formula Issues",
      runningSummary: `Running formula highlight workflow for ${scopeLabel}.`,
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
          stepId: "stage-issue-highlights",
          label: "Stage issue highlights",
          runningSummary: "Staging semantic highlight commands for the detected formula issues.",
          pendingSummary:
            "Waiting to stage semantic highlight commands for the detected formula issues.",
        },
        {
          stepId: "draft-highlight-report",
          label: "Draft highlight report",
          runningSummary: "Drafting the durable formula highlight report.",
          pendingSummary: "Waiting to assemble the durable formula highlight report.",
        },
      ],
    };
  }
  return null;
}

export async function executeFormulaWorkflow(input: {
  documentId: string;
  zeroSyncService: ZeroSyncService;
  workflowTemplate: WorkbookAgentWorkflowTemplate;
  workflowInput?: FormulaWorkflowExecutionInput | null;
  signal?: AbortSignal;
}): Promise<FormulaWorkflowExecutionResult | null> {
  if (
    input.workflowTemplate !== "findFormulaIssues" &&
    input.workflowTemplate !== "highlightFormulaIssues"
  ) {
    return null;
  }
  throwIfWorkflowCancelled(input.signal);
  const formulaIssues = await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) =>
    findWorkbookFormulaIssues(runtime, {
      ...(input.workflowInput?.sheetName ? { sheetName: input.workflowInput.sheetName } : {}),
      ...(input.workflowInput?.limit !== undefined ? { limit: input.workflowInput.limit } : {}),
    }),
  );
  throwIfWorkflowCancelled(input.signal);
  const scopeLabel = input.workflowInput?.sheetName ? ` on ${input.workflowInput.sheetName}` : "";
  if (input.workflowTemplate === "findFormulaIssues") {
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
      citations: createIssueCitations(formulaIssues, "source"),
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
  return {
    title: "Highlight Formula Issues",
    summary:
      formulaIssues.summary.issueCount === 0
        ? `Scanned ${String(formulaIssues.summary.scannedFormulaCells)} formula cell${formulaIssues.summary.scannedFormulaCells === 1 ? "" : "s"}${scopeLabel} and found no issues to highlight.`
        : `Staged highlight formatting for ${String(formulaIssues.summary.issueCount)} formula issue${formulaIssues.summary.issueCount === 1 ? "" : "s"}${scopeLabel}.`,
    artifact: {
      kind: "markdown",
      title: "Formula Issue Highlights",
      text: summarizeHighlightedFormulaIssuesMarkdown(formulaIssues),
    },
    citations: createIssueCitations(formulaIssues, "target"),
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
        stepId: "stage-issue-highlights",
        label: "Stage issue highlights",
        summary:
          formulaIssues.summary.issueCount === 0
            ? "No issue highlight commands were staged because no formula issues were found."
            : `Prepared ${String(formulaIssues.issues.length)} semantic formatting command${formulaIssues.issues.length === 1 ? "" : "s"} to highlight the detected formula issues.`,
      },
      {
        stepId: "draft-highlight-report",
        label: "Draft highlight report",
        summary: "Prepared the durable formula highlight report for the thread.",
      },
    ],
    commands: formulaIssues.issues.map((issue) => ({
      kind: "formatRange" as const,
      range: {
        sheetName: issue.sheetName,
        startAddress: issue.address,
        endAddress: issue.address,
      },
      patch: {
        fill: {
          backgroundColor: "#FEE2E2",
        },
        font: {
          bold: true,
          color: "#991B1B",
        },
      },
    })),
    goalText: input.workflowInput?.sheetName
      ? `Highlight formula issues on ${input.workflowInput.sheetName}`
      : "Highlight formula issues in the workbook",
  };
}
