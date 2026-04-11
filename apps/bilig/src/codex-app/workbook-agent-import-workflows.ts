import type { WorkbookAgentCommand, WorkbookAgentWriteCellInput } from "@bilig/agent-api";
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from "@bilig/contracts";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { ZeroSyncService } from "../zero/service.js";

type ImportWorkflowTemplate = "normalizeCurrentSheetHeaders";

export interface ImportWorkflowExecutionInput {
  readonly sheetName?: string;
}

interface ImportWorkflowStepPlan {
  readonly stepId: string;
  readonly label: string;
  readonly runningSummary: string;
  readonly pendingSummary: string;
}

interface ImportWorkflowStepResult {
  readonly stepId: string;
  readonly label: string;
  readonly summary: string;
}

interface ImportWorkflowTemplateMetadata {
  readonly title: string;
  readonly runningSummary: string;
  readonly stepPlans: readonly ImportWorkflowStepPlan[];
}

export interface ImportWorkflowExecutionResult {
  readonly title: string;
  readonly summary: string;
  readonly artifact: WorkbookAgentWorkflowArtifact;
  readonly citations: readonly WorkbookAgentTimelineCitation[];
  readonly steps: readonly ImportWorkflowStepResult[];
  readonly commands?: readonly WorkbookAgentCommand[];
  readonly goalText?: string;
}

function resolveWorkflowSheetName(input: {
  readonly workflowInput?: ImportWorkflowExecutionInput | null;
  readonly context?: WorkbookAgentUiContext | null;
}): string | null {
  const explicitName = input.workflowInput?.sheetName?.trim();
  if (explicitName && explicitName.length > 0) {
    return explicitName;
  }
  return input.context?.selection.sheetName ?? null;
}

function createHeaderCitation(
  sheetName: string,
  startAddress: string,
  endAddress: string,
  role: "source" | "target",
): WorkbookAgentTimelineCitation {
  return {
    kind: "range",
    sheetName,
    startAddress,
    endAddress,
    role,
  };
}

function normalizeHeaderLabel(label: string): string {
  const separatedCamel = label.replace(/([a-z0-9])([A-Z])/g, "$1 $2");
  const collapsed = separatedCamel.replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (collapsed.length === 0) {
    return "";
  }
  return collapsed
    .split(" ")
    .map((part) => {
      if (/^[A-Z0-9]{2,5}$/.test(part)) {
        return part;
      }
      const lower = part.toLowerCase();
      return `${lower.slice(0, 1).toUpperCase()}${lower.slice(1)}`;
    })
    .join(" ");
}

function dedupeHeaderLabels(labels: readonly string[]): readonly string[] {
  const seen = new Map<string, number>();
  return labels.map((label) => {
    if (label.length === 0) {
      return label;
    }
    const key = label.toLowerCase();
    const nextCount = (seen.get(key) ?? 0) + 1;
    seen.set(key, nextCount);
    return nextCount === 1 ? label : `${label} ${String(nextCount)}`;
  });
}

function summarizeHeaderNormalizationMarkdown(input: {
  readonly sheetName: string;
  readonly headerStartAddress: string;
  readonly headerEndAddress: string;
  readonly totalColumns: number;
  readonly changes: readonly {
    readonly address: string;
    readonly before: string;
    readonly after: string;
  }[];
}): string {
  const lines = [
    "## Header Normalization Preview",
    "",
    `Sheet: ${input.sheetName}`,
    `Header row: ${input.headerStartAddress}:${input.headerEndAddress}`,
    `Columns inspected: ${String(input.totalColumns)}`,
    `Headers changed: ${String(input.changes.length)}`,
    "",
  ];
  if (input.changes.length === 0) {
    lines.push("No header changes were needed. The current header row is already normalized.");
    return lines.join("\n");
  }
  lines.push("### Changed headers");
  for (const change of input.changes) {
    lines.push(`- ${change.address}: ${change.before} -> ${change.after}`);
  }
  lines.push(
    "",
    "The staged preview bundle writes the normalized header row through the normal workbook mutation path.",
  );
  return lines.join("\n");
}

export function getImportWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate | ImportWorkflowTemplate,
  workflowInput?: ImportWorkflowExecutionInput | null,
): ImportWorkflowTemplateMetadata | null {
  if (workflowTemplate !== "normalizeCurrentSheetHeaders") {
    return null;
  }
  const scopeLabel = workflowInput?.sheetName ?? "the active sheet";
  return {
    title: "Normalize Current Sheet Headers",
    runningSummary: `Running header normalization workflow for ${scopeLabel}.`,
    stepPlans: [
      {
        stepId: "inspect-header-row",
        label: "Inspect header row",
        runningSummary: `Inspecting the used range and header row on ${scopeLabel}.`,
        pendingSummary: `Waiting to inspect the used range and header row on ${scopeLabel}.`,
      },
      {
        stepId: "stage-header-normalization",
        label: "Stage header normalization",
        runningSummary: "Staging the semantic preview that normalizes the current sheet header row.",
        pendingSummary:
          "Waiting to stage the semantic preview that normalizes the current sheet header row.",
      },
      {
        stepId: "draft-header-report",
        label: "Draft header report",
        runningSummary: "Drafting the durable header normalization report.",
        pendingSummary: "Waiting to assemble the durable header normalization report.",
      },
    ],
  };
}

export async function executeImportWorkflow(input: {
  readonly documentId: string;
  readonly zeroSyncService: ZeroSyncService;
  readonly workflowTemplate: WorkbookAgentWorkflowTemplate | ImportWorkflowTemplate;
  readonly context?: WorkbookAgentUiContext | null;
  readonly workflowInput?: ImportWorkflowExecutionInput | null;
}): Promise<ImportWorkflowExecutionResult | null> {
  if (input.workflowTemplate !== "normalizeCurrentSheetHeaders") {
    return null;
  }
  const sheetName = resolveWorkflowSheetName({
    ...(input.workflowInput !== undefined ? { workflowInput: input.workflowInput } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  });
  if (!sheetName) {
    throw new Error("Selection context is required for header normalization workflows.");
  }

  return await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    const snapshot = runtime.engine.exportSnapshot();
    const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName);
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} was not found in the workbook.`);
    }
    if (sheet.cells.length === 0) {
      return {
        title: "Normalize Current Sheet Headers",
        summary: `${sheetName} is empty, so there were no headers to normalize.`,
        artifact: {
          kind: "markdown",
          title: "Header Normalization Preview",
          text: [
            "## Header Normalization Preview",
            "",
            `Sheet: ${sheetName}`,
            "",
            "No header changes were needed because the sheet is empty.",
          ].join("\n"),
        },
        citations: [],
        steps: [
          {
            stepId: "inspect-header-row",
            label: "Inspect header row",
            summary: `Loaded ${sheetName} and found no populated cells.`,
          },
          {
            stepId: "stage-header-normalization",
            label: "Stage header normalization",
            summary: "No header normalization preview was staged because the sheet is empty.",
          },
          {
            stepId: "draft-header-report",
            label: "Draft header report",
            summary: "Prepared the durable empty-sheet header report for the thread.",
          },
        ],
      } satisfies ImportWorkflowExecutionResult;
    }

    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;
    const cellByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell] as const));
    for (const cell of sheet.cells) {
      const parsed = parseCellAddress(cell.address, sheet.name);
      minRow = Math.min(minRow, parsed.row);
      minCol = Math.min(minCol, parsed.col);
      maxCol = Math.max(maxCol, parsed.col);
    }

    const headerRow = minRow;
    const headerStartAddress = formatAddress(headerRow, minCol);
    const headerEndAddress = formatAddress(headerRow, maxCol);
    const headerInputs: WorkbookAgentWriteCellInput[] = [];
    const candidateLabels: string[] = [];
    const candidateColumns: number[] = [];

    for (let col = minCol; col <= maxCol; col += 1) {
      const address = formatAddress(headerRow, col);
      const cell = cellByAddress.get(address);
      if (!cell) {
        headerInputs.push(null);
        continue;
      }
      if (cell.formula) {
        headerInputs.push({ formula: `=${cell.formula}` });
        continue;
      }
      if (typeof cell.value === "string") {
        headerInputs.push(cell.value);
        candidateLabels.push(normalizeHeaderLabel(cell.value));
        candidateColumns.push(col);
        continue;
      }
      headerInputs.push(cell.value ?? null);
    }

    const dedupedLabels = dedupeHeaderLabels(candidateLabels);
    const changes: Array<{
      readonly address: string;
      readonly before: string;
      readonly after: string;
    }> = [];
    for (let index = 0; index < candidateColumns.length; index += 1) {
      const col = candidateColumns[index]!;
      const nextLabel = dedupedLabels[index]!;
      const address = formatAddress(headerRow, col);
      const currentValue = headerInputs[col - minCol];
      if (typeof currentValue !== "string" || currentValue === nextLabel) {
        continue;
      }
      headerInputs[col - minCol] = nextLabel;
      changes.push({
        address,
        before: currentValue,
        after: nextLabel,
      });
    }

    const artifact: WorkbookAgentWorkflowArtifact = {
      kind: "markdown",
      title: "Header Normalization Preview",
      text: summarizeHeaderNormalizationMarkdown({
        sheetName,
        headerStartAddress,
        headerEndAddress,
        totalColumns: maxCol - minCol + 1,
        changes,
      }),
    };
    const citationRole = changes.length > 0 ? "target" : "source";
    return {
      title: "Normalize Current Sheet Headers",
      summary:
        changes.length === 0
          ? `Checked ${sheetName} headers and found no normalization changes to stage.`
          : `Staged normalized headers for ${String(changes.length)} cell${changes.length === 1 ? "" : "s"} on ${sheetName}.`,
      artifact,
      citations: [createHeaderCitation(sheetName, headerStartAddress, headerEndAddress, citationRole)],
      steps: [
        {
          stepId: "inspect-header-row",
          label: "Inspect header row",
          summary: `Loaded the used range and current header row from ${sheetName}.`,
        },
        {
          stepId: "stage-header-normalization",
          label: "Stage header normalization",
          summary:
            changes.length === 0
              ? "No header normalization preview was staged because the header row is already normalized."
              : `Prepared the semantic write preview that normalizes ${String(changes.length)} header cell${changes.length === 1 ? "" : "s"}.`,
        },
        {
          stepId: "draft-header-report",
          label: "Draft header report",
          summary: "Prepared the durable header normalization report for the thread.",
        },
      ],
      ...(changes.length > 0
        ? {
            commands: [
              {
                kind: "writeRange" as const,
                sheetName,
                startAddress: headerStartAddress,
                values: [headerInputs],
              },
            ],
            goalText: `Normalize header row on ${sheetName}`,
          }
        : {}),
    } satisfies ImportWorkflowExecutionResult;
  });
}
