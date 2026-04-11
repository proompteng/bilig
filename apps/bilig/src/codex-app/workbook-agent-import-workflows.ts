import type { WorkbookAgentCommand, WorkbookAgentWriteCellInput } from "@bilig/agent-api";
import type {
  WorkbookAgentTimelineCitation,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowArtifact,
  WorkbookAgentWorkflowTemplate,
} from "@bilig/contracts";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  buildCellNumberFormatCode,
  type CellNumberFormatPreset,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import type { ZeroSyncService } from "../zero/service.js";
import { throwIfWorkflowCancelled } from "./workbook-agent-workflow-abort.js";

type ImportWorkflowTemplate =
  | "normalizeCurrentSheetHeaders"
  | "normalizeCurrentSheetNumberFormats"
  | "normalizeCurrentSheetWhitespace"
  | "fillCurrentSheetFormulasDown";
type SnapshotSheet = WorkbookSnapshot["sheets"][number];
type SnapshotCell = SnapshotSheet["cells"][number];

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

interface SheetExtents {
  readonly headerRow: number;
  readonly dataStartRow: number;
  readonly minCol: number;
  readonly maxCol: number;
  readonly maxRow: number;
  readonly cellByAddress: Map<string, SnapshotCell>;
}

interface NumberFormatRecommendation {
  readonly headerLabel: string;
  readonly columnLabel: string;
  readonly startAddress: string;
  readonly endAddress: string;
  readonly preset: CellNumberFormatPreset;
  readonly numericCount: number;
}

interface FormulaFillRecommendation {
  readonly columnLabel: string;
  readonly sourceAddress: string;
  readonly targetStartAddress: string;
  readonly targetEndAddress: string;
  readonly filledRowCount: number;
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

function isImportWorkflowTemplate(
  workflowTemplate: WorkbookAgentWorkflowTemplate | ImportWorkflowTemplate,
): workflowTemplate is ImportWorkflowTemplate {
  return (
    workflowTemplate === "normalizeCurrentSheetHeaders" ||
    workflowTemplate === "normalizeCurrentSheetNumberFormats" ||
    workflowTemplate === "normalizeCurrentSheetWhitespace"
  );
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

function inspectSheetExtents(sheet: SnapshotSheet): SheetExtents {
  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  const cellByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell] as const));
  for (const cell of sheet.cells) {
    const parsed = parseCellAddress(cell.address, sheet.name);
    minRow = Math.min(minRow, parsed.row);
    minCol = Math.min(minCol, parsed.col);
    maxRow = Math.max(maxRow, parsed.row);
    maxCol = Math.max(maxCol, parsed.col);
  }
  return {
    headerRow: minRow,
    dataStartRow: minRow + 1,
    minCol,
    maxCol,
    maxRow,
    cellByAddress,
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

function inferFractionDigits(values: readonly number[]): number {
  let maxDigits = 0;
  for (const value of values) {
    const rendered = value.toString();
    const decimalIndex = rendered.indexOf(".");
    if (decimalIndex === -1) {
      continue;
    }
    const digits = rendered.slice(decimalIndex + 1).replace(/0+$/u, "").length;
    maxDigits = Math.max(maxDigits, digits);
  }
  return Math.min(maxDigits, 4);
}

function normalizeHeaderKey(label: string): string {
  return label
    .toLowerCase()
    .replace(/[^a-z0-9]+/gu, " ")
    .trim();
}

function inferNumberFormatPreset(
  headerLabel: string,
  numericValues: readonly number[],
): CellNumberFormatPreset | null {
  if (numericValues.length === 0) {
    return null;
  }
  const headerKey = normalizeHeaderKey(headerLabel);
  const fractionDigits = Math.max(0, inferFractionDigits(numericValues));
  const maxAbs = Math.max(...numericValues.map((value) => Math.abs(value)));
  const allIntegers = numericValues.every((value) => Number.isInteger(value));

  if (
    /(?:date|month|quarter|week|day|year|as of|opened|closed|created|updated)/u.test(headerKey) &&
    allIntegers &&
    numericValues.every((value) => value >= 20_000 && value <= 90_000)
  ) {
    return {
      kind: "date",
      dateStyle: "short",
    };
  }

  if (/(?:percent|pct|margin|rate|growth|share|ratio)/u.test(headerKey) && maxAbs <= 1.5) {
    return {
      kind: "percent",
      decimals: Math.max(2, fractionDigits),
    };
  }

  if (
    /(?:revenue|sales|amount|price|cost|expense|profit|budget|income|total|subtotal|usd|eur|gbp|arr|mrr|gmv)/u.test(
      headerKey,
    )
  ) {
    return {
      kind: "currency",
      currency: "USD",
      decimals: Math.max(2, fractionDigits),
      useGrouping: true,
      negativeStyle: "minus",
      zeroStyle: "zero",
    };
  }

  return {
    kind: "number",
    decimals: fractionDigits,
    useGrouping: true,
  };
}

function describeNumberFormatPreset(preset: CellNumberFormatPreset): string {
  switch (preset.kind) {
    case "general":
      return "general";
    case "text":
      return "text";
    case "number":
      return `number (${String(preset.decimals ?? 0)} decimals)`;
    case "currency":
      return `currency ${preset.currency ?? "USD"} (${String(preset.decimals ?? 2)} decimals)`;
    case "accounting":
      return `accounting ${preset.currency ?? "USD"} (${String(preset.decimals ?? 2)} decimals)`;
    case "date":
      return `date (${preset.dateStyle ?? "short"})`;
    case "time":
      return `time (${preset.dateStyle ?? "short"})`;
    case "datetime":
      return `datetime (${preset.dateStyle ?? "short"})`;
    case "percent":
      return `percent (${String(preset.decimals ?? 2)} decimals)`;
  }
}

function summarizeNumberFormatNormalizationMarkdown(input: {
  readonly sheetName: string;
  readonly dataStartAddress: string;
  readonly dataEndAddress: string;
  readonly recommendations: readonly NumberFormatRecommendation[];
}): string {
  const lines = [
    "## Number Format Normalization Preview",
    "",
    `Sheet: ${input.sheetName}`,
    `Data range: ${input.dataStartAddress}:${input.dataEndAddress}`,
    `Columns staged: ${String(input.recommendations.length)}`,
    "",
  ];
  if (input.recommendations.length === 0) {
    lines.push("No number-format changes were needed on the current sheet.");
    return lines.join("\n");
  }
  lines.push("### Staged number formats");
  for (const recommendation of input.recommendations) {
    lines.push(
      `- ${recommendation.columnLabel} (${recommendation.headerLabel || "Untitled"}): ${describeNumberFormatPreset(recommendation.preset)} across ${recommendation.startAddress}:${recommendation.endAddress} (${String(recommendation.numericCount)} numeric cell${recommendation.numericCount === 1 ? "" : "s"})`,
    );
  }
  lines.push(
    "",
    "The staged preview bundle applies semantic number-format commands through the normal workbook mutation path.",
  );
  return lines.join("\n");
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

function normalizeWhitespaceText(value: string): string {
  return value.replace(/\s+/gu, " ").trim();
}

function summarizeWhitespaceNormalizationMarkdown(input: {
  readonly sheetName: string;
  readonly rangeStartAddress: string;
  readonly rangeEndAddress: string;
  readonly changes: readonly {
    readonly address: string;
    readonly before: string;
    readonly after: string;
  }[];
}): string {
  const lines = [
    "## Whitespace Normalization Preview",
    "",
    `Sheet: ${input.sheetName}`,
    `Inspected range: ${input.rangeStartAddress}:${input.rangeEndAddress}`,
    `Text cells changed: ${String(input.changes.length)}`,
    "",
  ];
  if (input.changes.length === 0) {
    lines.push("No text whitespace changes were needed on the current sheet.");
    return lines.join("\n");
  }
  lines.push("### Normalized cells");
  for (const change of input.changes) {
    lines.push(
      `- ${change.address}: ${JSON.stringify(change.before)} -> ${JSON.stringify(change.after)}`,
    );
  }
  lines.push(
    "",
    "The staged preview bundle writes the normalized text cells through the normal workbook mutation path.",
  );
  return lines.join("\n");
}

function isBlankTextCandidate(cell: SnapshotCell | undefined): boolean {
  if (!cell) {
    return true;
  }
  if (cell.formula) {
    return false;
  }
  if (cell.value === null || cell.value === undefined) {
    return true;
  }
  return typeof cell.value === "string" && cell.value.trim().length === 0;
}

function summarizeFormulaFillMarkdown(input: {
  readonly sheetName: string;
  readonly recommendations: readonly FormulaFillRecommendation[];
}): string {
  const lines = [
    "## Formula Fill-Down Preview",
    "",
    `Sheet: ${input.sheetName}`,
    `Formula regions staged: ${String(input.recommendations.length)}`,
    "",
  ];
  if (input.recommendations.length === 0) {
    lines.push("No fill-down changes were needed on the current sheet.");
    return lines.join("\n");
  }
  lines.push("### Filled ranges");
  for (const recommendation of input.recommendations) {
    lines.push(
      `- ${recommendation.columnLabel}: fill ${recommendation.sourceAddress} down through ${recommendation.targetStartAddress}:${recommendation.targetEndAddress} (${String(recommendation.filledRowCount)} row${recommendation.filledRowCount === 1 ? "" : "s"})`,
    );
  }
  lines.push(
    "",
    "The staged preview bundle applies semantic fill commands through the normal workbook mutation path.",
  );
  return lines.join("\n");
}

export function getImportWorkflowTemplateMetadata(
  workflowTemplate: WorkbookAgentWorkflowTemplate | ImportWorkflowTemplate,
  workflowInput?: ImportWorkflowExecutionInput | null,
): ImportWorkflowTemplateMetadata | null {
  if (!isImportWorkflowTemplate(workflowTemplate)) {
    return null;
  }
  const scopeLabel = workflowInput?.sheetName ?? "the active sheet";
  switch (workflowTemplate) {
    case "normalizeCurrentSheetHeaders":
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
            runningSummary:
              "Staging the semantic preview that normalizes the current sheet header row.",
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
    case "normalizeCurrentSheetNumberFormats":
      return {
        title: "Normalize Current Sheet Number Formats",
        runningSummary: `Running number-format normalization workflow for ${scopeLabel}.`,
        stepPlans: [
          {
            stepId: "inspect-number-columns",
            label: "Inspect numeric columns",
            runningSummary: `Inspecting numeric columns and header labels on ${scopeLabel}.`,
            pendingSummary: `Waiting to inspect numeric columns and header labels on ${scopeLabel}.`,
          },
          {
            stepId: "stage-number-formats",
            label: "Stage number formats",
            runningSummary:
              "Staging semantic number-format previews for the current sheet data columns.",
            pendingSummary:
              "Waiting to stage semantic number-format previews for the current sheet data columns.",
          },
          {
            stepId: "draft-number-format-report",
            label: "Draft number-format report",
            runningSummary: "Drafting the durable number-format normalization report.",
            pendingSummary: "Waiting to assemble the durable number-format normalization report.",
          },
        ],
      };
    case "normalizeCurrentSheetWhitespace":
      return {
        title: "Normalize Current Sheet Whitespace",
        runningSummary: `Running whitespace normalization workflow for ${scopeLabel}.`,
        stepPlans: [
          {
            stepId: "inspect-text-cells",
            label: "Inspect text cells",
            runningSummary: `Inspecting text cells and the used range on ${scopeLabel}.`,
            pendingSummary: `Waiting to inspect text cells and the used range on ${scopeLabel}.`,
          },
          {
            stepId: "stage-whitespace-normalization",
            label: "Stage whitespace normalization",
            runningSummary:
              "Staging the semantic preview that trims and collapses whitespace across the current sheet.",
            pendingSummary:
              "Waiting to stage the semantic preview that trims and collapses whitespace across the current sheet.",
          },
          {
            stepId: "draft-whitespace-report",
            label: "Draft whitespace report",
            runningSummary: "Drafting the durable whitespace normalization report.",
            pendingSummary: "Waiting to assemble the durable whitespace normalization report.",
          },
        ],
      };
    case "fillCurrentSheetFormulasDown":
      return {
        title: "Fill Current Sheet Formulas Down",
        runningSummary: `Running formula fill-down workflow for ${scopeLabel}.`,
        stepPlans: [
          {
            stepId: "inspect-formula-columns",
            label: "Inspect formula columns",
            runningSummary: `Inspecting formula cells and blank gaps on ${scopeLabel}.`,
            pendingSummary: `Waiting to inspect formula cells and blank gaps on ${scopeLabel}.`,
          },
          {
            stepId: "stage-formula-fill",
            label: "Stage formula fill-down",
            runningSummary: "Staging semantic fill-down previews for the detected formula gaps.",
            pendingSummary:
              "Waiting to stage semantic fill-down previews for the detected formula gaps.",
          },
          {
            stepId: "draft-formula-fill-report",
            label: "Draft formula fill-down report",
            runningSummary: "Drafting the durable formula fill-down report.",
            pendingSummary: "Waiting to assemble the durable formula fill-down report.",
          },
        ],
      };
    default:
      return null;
  }
}

export async function executeImportWorkflow(input: {
  readonly documentId: string;
  readonly zeroSyncService: ZeroSyncService;
  readonly workflowTemplate: WorkbookAgentWorkflowTemplate | ImportWorkflowTemplate;
  readonly context?: WorkbookAgentUiContext | null;
  readonly workflowInput?: ImportWorkflowExecutionInput | null;
  readonly signal?: AbortSignal;
}): Promise<ImportWorkflowExecutionResult | null> {
  if (
    input.workflowTemplate !== "normalizeCurrentSheetHeaders" &&
    input.workflowTemplate !== "normalizeCurrentSheetNumberFormats" &&
    input.workflowTemplate !== "normalizeCurrentSheetWhitespace" &&
    input.workflowTemplate !== "fillCurrentSheetFormulasDown"
  ) {
    return null;
  }
  const sheetName = resolveWorkflowSheetName({
    ...(input.workflowInput !== undefined ? { workflowInput: input.workflowInput } : {}),
    ...(input.context !== undefined ? { context: input.context } : {}),
  });
  if (!sheetName) {
    throw new Error("Selection context is required for sheet-scoped import workflows.");
  }

  throwIfWorkflowCancelled(input.signal);
  return await input.zeroSyncService.inspectWorkbook(input.documentId, (runtime) => {
    throwIfWorkflowCancelled(input.signal);
    const snapshot = runtime.engine.exportSnapshot();
    const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName);
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} was not found in the workbook.`);
    }
    if (sheet.cells.length === 0) {
      if (input.workflowTemplate === "normalizeCurrentSheetHeaders") {
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
      if (input.workflowTemplate === "normalizeCurrentSheetWhitespace") {
        return {
          title: "Normalize Current Sheet Whitespace",
          summary: `${sheetName} is empty, so there were no text cells to normalize.`,
          artifact: {
            kind: "markdown",
            title: "Whitespace Normalization Preview",
            text: [
              "## Whitespace Normalization Preview",
              "",
              `Sheet: ${sheetName}`,
              "",
              "No whitespace changes were needed because the sheet is empty.",
            ].join("\n"),
          },
          citations: [],
          steps: [
            {
              stepId: "inspect-text-cells",
              label: "Inspect text cells",
              summary: `Loaded ${sheetName} and found no populated cells.`,
            },
            {
              stepId: "stage-whitespace-normalization",
              label: "Stage whitespace normalization",
              summary: "No whitespace normalization preview was staged because the sheet is empty.",
            },
            {
              stepId: "draft-whitespace-report",
              label: "Draft whitespace report",
              summary: "Prepared the durable empty-sheet whitespace report for the thread.",
            },
          ],
        } satisfies ImportWorkflowExecutionResult;
      }
      if (input.workflowTemplate === "fillCurrentSheetFormulasDown") {
        return {
          title: "Fill Current Sheet Formulas Down",
          summary: `${sheetName} is empty, so there were no formula gaps to fill.`,
          artifact: {
            kind: "markdown",
            title: "Formula Fill-Down Preview",
            text: [
              "## Formula Fill-Down Preview",
              "",
              `Sheet: ${sheetName}`,
              "",
              "No fill-down changes were needed because the sheet is empty.",
            ].join("\n"),
          },
          citations: [],
          steps: [
            {
              stepId: "inspect-formula-columns",
              label: "Inspect formula columns",
              summary: `Loaded ${sheetName} and found no populated cells.`,
            },
            {
              stepId: "stage-formula-fill",
              label: "Stage formula fill-down",
              summary: "No formula fill-down preview was staged because the sheet is empty.",
            },
            {
              stepId: "draft-formula-fill-report",
              label: "Draft formula fill-down report",
              summary: "Prepared the durable empty-sheet formula fill-down report for the thread.",
            },
          ],
        } satisfies ImportWorkflowExecutionResult;
      }
      return {
        title: "Normalize Current Sheet Number Formats",
        summary: `${sheetName} is empty, so there were no numeric cells to format.`,
        artifact: {
          kind: "markdown",
          title: "Number Format Normalization Preview",
          text: [
            "## Number Format Normalization Preview",
            "",
            `Sheet: ${sheetName}`,
            "",
            "No number-format changes were needed because the sheet is empty.",
          ].join("\n"),
        },
        citations: [],
        steps: [
          {
            stepId: "inspect-number-columns",
            label: "Inspect numeric columns",
            summary: `Loaded ${sheetName} and found no populated cells.`,
          },
          {
            stepId: "stage-number-formats",
            label: "Stage number formats",
            summary: "No number-format preview was staged because the sheet is empty.",
          },
          {
            stepId: "draft-number-format-report",
            label: "Draft number-format report",
            summary: "Prepared the durable empty-sheet number-format report for the thread.",
          },
        ],
      } satisfies ImportWorkflowExecutionResult;
    }

    const { headerRow, dataStartRow, minCol, maxCol, maxRow, cellByAddress } =
      inspectSheetExtents(sheet);
    const headerStartAddress = formatAddress(headerRow, minCol);
    const headerEndAddress = formatAddress(headerRow, maxCol);

    if (input.workflowTemplate === "normalizeCurrentSheetNumberFormats") {
      const recommendations: NumberFormatRecommendation[] = [];
      for (let col = minCol; col <= maxCol; col += 1) {
        const headerAddress = formatAddress(headerRow, col);
        const headerCell = cellByAddress.get(headerAddress);
        const headerLabel =
          headerCell && typeof headerCell.value === "string"
            ? normalizeHeaderLabel(headerCell.value)
            : "";
        const numericCells: SnapshotCell[] = [];
        const numericValues: number[] = [];
        for (let row = dataStartRow; row <= maxRow; row += 1) {
          const cell = cellByAddress.get(formatAddress(row, col));
          if (!cell || typeof cell.value !== "number") {
            continue;
          }
          numericCells.push(cell);
          numericValues.push(cell.value);
        }
        const preset = inferNumberFormatPreset(headerLabel, numericValues);
        if (!preset || numericCells.length === 0) {
          continue;
        }
        const targetCode = buildCellNumberFormatCode(preset);
        if (numericCells.every((cell) => (cell.format ?? "general") === targetCode)) {
          continue;
        }
        recommendations.push({
          headerLabel: headerLabel || formatAddress(0, col).replace(/\d+/gu, ""),
          columnLabel: formatAddress(0, col).replace(/\d+/gu, ""),
          startAddress: formatAddress(dataStartRow, col),
          endAddress: formatAddress(maxRow, col),
          preset,
          numericCount: numericCells.length,
        });
      }

      const dataStartAddress =
        dataStartRow <= maxRow ? formatAddress(dataStartRow, minCol) : headerStartAddress;
      const dataEndAddress =
        dataStartRow <= maxRow ? formatAddress(maxRow, maxCol) : headerEndAddress;
      return {
        title: "Normalize Current Sheet Number Formats",
        summary:
          recommendations.length === 0
            ? `Checked ${sheetName} number formats and found no normalization changes to stage.`
            : `Staged normalized number formats for ${String(recommendations.length)} column${recommendations.length === 1 ? "" : "s"} on ${sheetName}.`,
        artifact: {
          kind: "markdown",
          title: "Number Format Normalization Preview",
          text: summarizeNumberFormatNormalizationMarkdown({
            sheetName,
            dataStartAddress,
            dataEndAddress,
            recommendations,
          }),
        },
        citations: [
          {
            kind: "range",
            sheetName,
            startAddress: dataStartAddress,
            endAddress: dataEndAddress,
            role: recommendations.length > 0 ? "target" : "source",
          },
        ],
        steps: [
          {
            stepId: "inspect-number-columns",
            label: "Inspect numeric columns",
            summary: `Loaded numeric cells and header labels from ${sheetName}.`,
          },
          {
            stepId: "stage-number-formats",
            label: "Stage number formats",
            summary:
              recommendations.length === 0
                ? "No number-format preview was staged because the sheet is already normalized."
                : `Prepared semantic number-format previews for ${String(recommendations.length)} numeric column${recommendations.length === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-number-format-report",
            label: "Draft number-format report",
            summary: "Prepared the durable number-format normalization report for the thread.",
          },
        ],
        ...(recommendations.length > 0
          ? {
              commands: recommendations.map(
                (recommendation) =>
                  ({
                    kind: "formatRange" as const,
                    range: {
                      sheetName,
                      startAddress: recommendation.startAddress,
                      endAddress: recommendation.endAddress,
                    },
                    numberFormat: recommendation.preset,
                  }) satisfies WorkbookAgentCommand,
              ),
              goalText: `Normalize number formats on ${sheetName}`,
            }
          : {}),
      } satisfies ImportWorkflowExecutionResult;
    }

    if (input.workflowTemplate === "normalizeCurrentSheetWhitespace") {
      const rowValues: WorkbookAgentWriteCellInput[][] = [];
      const changes: Array<{
        readonly address: string;
        readonly before: string;
        readonly after: string;
      }> = [];

      for (let row = headerRow; row <= maxRow; row += 1) {
        const rowInputs: WorkbookAgentWriteCellInput[] = [];
        for (let col = minCol; col <= maxCol; col += 1) {
          const address = formatAddress(row, col);
          const cell = cellByAddress.get(address);
          if (!cell) {
            rowInputs.push(null);
            continue;
          }
          if (cell.formula) {
            rowInputs.push({ formula: `=${cell.formula}` });
            continue;
          }
          if (typeof cell.value === "string") {
            const normalized = normalizeWhitespaceText(cell.value);
            rowInputs.push(normalized);
            if (normalized !== cell.value) {
              changes.push({
                address,
                before: cell.value,
                after: normalized,
              });
            }
            continue;
          }
          rowInputs.push(cell.value ?? null);
        }
        rowValues.push(rowInputs);
      }

      const rangeStartAddress = formatAddress(headerRow, minCol);
      const rangeEndAddress = formatAddress(maxRow, maxCol);
      return {
        title: "Normalize Current Sheet Whitespace",
        summary:
          changes.length === 0
            ? `Checked ${sheetName} text cells and found no whitespace cleanup changes to stage.`
            : `Staged normalized whitespace for ${String(changes.length)} text cell${changes.length === 1 ? "" : "s"} on ${sheetName}.`,
        artifact: {
          kind: "markdown",
          title: "Whitespace Normalization Preview",
          text: summarizeWhitespaceNormalizationMarkdown({
            sheetName,
            rangeStartAddress,
            rangeEndAddress,
            changes,
          }),
        },
        citations: [
          createHeaderCitation(
            sheetName,
            rangeStartAddress,
            rangeEndAddress,
            changes.length > 0 ? "target" : "source",
          ),
        ],
        steps: [
          {
            stepId: "inspect-text-cells",
            label: "Inspect text cells",
            summary: `Loaded the used range and string cells from ${sheetName}.`,
          },
          {
            stepId: "stage-whitespace-normalization",
            label: "Stage whitespace normalization",
            summary:
              changes.length === 0
                ? "No whitespace normalization preview was staged because the sheet is already normalized."
                : `Prepared the semantic write preview that normalizes whitespace across ${String(changes.length)} text cell${changes.length === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-whitespace-report",
            label: "Draft whitespace report",
            summary: "Prepared the durable whitespace normalization report for the thread.",
          },
        ],
        ...(changes.length > 0
          ? {
              commands: [
                {
                  kind: "writeRange" as const,
                  sheetName,
                  startAddress: rangeStartAddress,
                  values: rowValues,
                },
              ],
              goalText: `Normalize whitespace on ${sheetName}`,
            }
          : {}),
      } satisfies ImportWorkflowExecutionResult;
    }

    if (input.workflowTemplate === "fillCurrentSheetFormulasDown") {
      const recommendations: FormulaFillRecommendation[] = [];
      for (let col = minCol; col <= maxCol; col += 1) {
        const headerCell = cellByAddress.get(formatAddress(headerRow, col));
        const columnLabel =
          headerCell && typeof headerCell.value === "string"
            ? normalizeHeaderLabel(headerCell.value)
            : formatAddress(0, col).replace(/\d+/gu, "");
        let row = dataStartRow;
        while (row <= maxRow) {
          const sourceAddress = formatAddress(row, col);
          const sourceCell = cellByAddress.get(sourceAddress);
          if (!sourceCell?.formula) {
            row += 1;
            continue;
          }
          let fillEndRow = row;
          for (let probe = row + 1; probe <= maxRow; probe += 1) {
            const candidate = cellByAddress.get(formatAddress(probe, col));
            if (!isBlankTextCandidate(candidate)) {
              break;
            }
            fillEndRow = probe;
          }
          if (fillEndRow > row) {
            recommendations.push({
              columnLabel: columnLabel || formatAddress(0, col).replace(/\d+/gu, ""),
              sourceAddress,
              targetStartAddress: formatAddress(row + 1, col),
              targetEndAddress: formatAddress(fillEndRow, col),
              filledRowCount: fillEndRow - row,
            });
          }
          row = Math.max(row + 1, fillEndRow + 1);
        }
      }

      return {
        title: "Fill Current Sheet Formulas Down",
        summary:
          recommendations.length === 0
            ? `Checked ${sheetName} formulas and found no blank ranges to fill down.`
            : `Staged formula fill-down for ${String(recommendations.length)} column${recommendations.length === 1 ? "" : "s"} on ${sheetName}.`,
        artifact: {
          kind: "markdown",
          title: "Formula Fill-Down Preview",
          text: summarizeFormulaFillMarkdown({
            sheetName,
            recommendations,
          }),
        },
        citations: recommendations.map((recommendation) =>
          createHeaderCitation(
            sheetName,
            recommendation.targetStartAddress,
            recommendation.targetEndAddress,
            "target",
          ),
        ),
        steps: [
          {
            stepId: "inspect-formula-columns",
            label: "Inspect formula columns",
            summary: `Loaded formula cells and blank fill gaps from ${sheetName}.`,
          },
          {
            stepId: "stage-formula-fill",
            label: "Stage formula fill-down",
            summary:
              recommendations.length === 0
                ? "No formula fill-down preview was staged because the sheet already has no blank gaps below formula cells."
                : `Prepared semantic fill-down previews for ${String(recommendations.length)} formula region${recommendations.length === 1 ? "" : "s"}.`,
          },
          {
            stepId: "draft-formula-fill-report",
            label: "Draft formula fill-down report",
            summary: "Prepared the durable formula fill-down report for the thread.",
          },
        ],
        ...(recommendations.length > 0
          ? {
              commands: recommendations.map(
                (recommendation) =>
                  ({
                    kind: "fillRange" as const,
                    source: {
                      sheetName,
                      startAddress: recommendation.sourceAddress,
                      endAddress: recommendation.sourceAddress,
                    },
                    target: {
                      sheetName,
                      startAddress: recommendation.targetStartAddress,
                      endAddress: recommendation.targetEndAddress,
                    },
                  }) satisfies WorkbookAgentCommand,
              ),
              goalText: `Fill formulas down on ${sheetName}`,
            }
          : {}),
      } satisfies ImportWorkflowExecutionResult;
    }

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
      citations: [
        createHeaderCitation(sheetName, headerStartAddress, headerEndAddress, citationRole),
      ],
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
