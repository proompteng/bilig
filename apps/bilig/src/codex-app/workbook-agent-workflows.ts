import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowRun,
  WorkbookAgentWorkflowTemplate,
} from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import { summarizeWorkbookStructure } from "./workbook-agent-comprehension.js";

export interface WorkbookAgentWorkflowExecutionResult {
  readonly title: string;
  readonly summary: string;
  readonly artifact: WorkbookAgentWorkflowArtifact;
  readonly citations: readonly WorkbookAgentTimelineCitation[];
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
    artifact: input.artifact ?? null,
  };
}

export async function executeWorkbookAgentWorkflow(input: {
  documentId: string;
  zeroSyncService: ZeroSyncService;
  workflowTemplate: WorkbookAgentWorkflowTemplate;
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
      };
    }
  }
}
