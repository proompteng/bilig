import { useEffect, useRef } from "react";
import { describeWorkbookAgentCommand, workbookAgentSkillDescriptors } from "@bilig/agent-api";
import type {
  WorkbookAgentCommandBundle,
  WorkbookAgentPreviewChangeKind,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
  WorkbookAgentSkillFocus,
} from "@bilig/agent-api";
import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
} from "@bilig/contracts";
import { cn } from "./cn.js";

function contextLabel(context: WorkbookAgentUiContext | null): string {
  if (!context) {
    return "No selection context";
  }
  return `${context.selection.sheetName}!${context.selection.address}`;
}

function skillFocusClass(focus: WorkbookAgentSkillFocus): string {
  switch (focus) {
    case "read":
      return "bg-[#e0f2fe] text-[#075985]";
    case "analyze":
      return "bg-[#ede9fe] text-[#5b21b6]";
    case "edit":
      return "bg-[#dcfce7] text-[#166534]";
  }
}

function WorkbookAgentSkillStrip(props: { readonly onUseSkillPrompt: (prompt: string) => void }) {
  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3">
      <div className="flex items-center justify-between gap-3">
        <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          Local Skills
        </div>
        <div className="text-[10px] text-[var(--wb-text-subtle)]">
          Monolith app-server + semantic workbook tools
        </div>
      </div>
      <div className="mt-2 grid gap-2">
        {workbookAgentSkillDescriptors.map((skill) => (
          <button
            key={skill.id}
            className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2 text-left shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-accent-ring)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            type="button"
            onClick={() => {
              props.onUseSkillPrompt(skill.prompt);
            }}
          >
            <div className="flex items-start justify-between gap-3">
              <div className="min-w-0">
                <div className="text-[12px] font-semibold text-[var(--wb-text)]">{skill.label}</div>
                <div className="mt-1 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
                  {skill.description}
                </div>
              </div>
              <span
                className={cn(
                  "rounded-full px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em]",
                  skillFocusClass(skill.focus),
                )}
              >
                {skill.focus}
              </span>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry["toolStatus"] }) {
  const label =
    props.status === "completed" ? "Done" : props.status === "failed" ? "Failed" : "Running";
  return (
    <span
      className={cn(
        "inline-flex h-5 items-center rounded-full px-2 text-[10px] font-semibold uppercase tracking-[0.04em]",
        props.status === "completed"
          ? "bg-[#dcfce7] text-[#166534]"
          : props.status === "failed"
            ? "bg-[#fee2e2] text-[#991b1b]"
            : "bg-[#e0f2fe] text-[#075985]",
      )}
    >
      {label}
    </span>
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function readString(value: unknown, fallback = ""): string {
  return typeof value === "string" ? value : fallback;
}

function readNumber(value: unknown, fallback = 0): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function safeParseToolOutput(outputText: string | null): unknown {
  if (!outputText) {
    return null;
  }
  try {
    return JSON.parse(outputText) as unknown;
  } catch {
    return null;
  }
}

function supportsStructuredToolOutput(toolName: string | null): boolean {
  return (
    toolName === "bilig.find_formula_issues" ||
    toolName === "bilig.search_workbook" ||
    toolName === "bilig.trace_dependencies"
  );
}

function renderReasonLabel(reason: string): string {
  switch (reason) {
    case "sheet":
      return "sheet";
    case "address":
      return "address";
    case "formula":
      return "formula";
    case "input":
      return "input";
    case "value":
      return "value";
    default:
      return reason;
  }
}

function renderPreviewChangeKind(kind: WorkbookAgentPreviewChangeKind): string {
  switch (kind) {
    case "input":
      return "value";
    case "formula":
      return "formula";
    case "style":
      return "style";
    case "numberFormat":
      return "number format";
  }
}

function StructuredToolOutput(props: {
  readonly toolName: string | null;
  readonly outputText: string | null;
}) {
  const parsed = safeParseToolOutput(props.outputText);
  if (!props.toolName || !isRecord(parsed)) {
    return null;
  }

  if (props.toolName === "bilig.find_formula_issues" && Array.isArray(parsed["issues"])) {
    const summary = isRecord(parsed["summary"]) ? parsed["summary"] : null;
    const issues = parsed["issues"].flatMap((issue) => (isRecord(issue) ? [issue] : []));
    return (
      <div className="mt-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
              Formula Issues
            </div>
            <div className="mt-1 text-[12px] text-[var(--wb-text-muted)]">
              {readNumber(summary?.["issueCount"], issues.length)} issues across{" "}
              {readNumber(summary?.["scannedFormulaCells"])} formulas
            </div>
          </div>
          <div className="text-right text-[10px] text-[var(--wb-text-subtle)]">
            {readNumber(summary?.["errorCount"])} errors · {readNumber(summary?.["cycleCount"])}{" "}
            cycles · {readNumber(summary?.["unsupportedCount"])} JS-only
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {issues.slice(0, 8).map((issue) => (
            <div
              key={`${readString(issue["sheetName"])}:${readString(issue["address"])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className="text-[12px] font-semibold text-[var(--wb-text)]">
                    {readString(issue["sheetName"])}!{readString(issue["address"])}
                  </div>
                  <div className="mt-1 break-all text-[11px] leading-5 text-[var(--wb-text-subtle)]">
                    {readString(issue["formula"])}
                  </div>
                </div>
                <div className="text-right text-[10px] text-[var(--wb-text-subtle)]">
                  {readString(issue["valueText"]) || "(empty)"}
                </div>
              </div>
              <div className="mt-2 flex flex-wrap gap-1.5">
                {Array.isArray(issue["issueKinds"])
                  ? issue["issueKinds"].map((kind) => (
                      <span
                        key={readString(kind)}
                        className="rounded-full bg-[#fee2e2] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em] text-[#991b1b]"
                      >
                        {readString(kind)}
                      </span>
                    ))
                  : null}
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (props.toolName === "bilig.search_workbook" && Array.isArray(parsed["matches"])) {
    const matches = parsed["matches"].flatMap((match) => (isRecord(match) ? [match] : []));
    return (
      <div className="mt-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
              Search Matches
            </div>
            <div className="mt-1 text-[12px] text-[var(--wb-text-muted)]">
              Query: “{readString(parsed["query"])}”
            </div>
          </div>
          <div className="text-[10px] text-[var(--wb-text-subtle)]">
            {readNumber(
              isRecord(parsed["summary"]) ? parsed["summary"]["matchCount"] : undefined,
              matches.length,
            )}{" "}
            matches
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {matches.slice(0, 8).map((match) => (
            <div
              key={`${readString(match["kind"])}:${readString(match["sheetName"])}:${readString(match["address"])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className="text-[12px] font-semibold text-[var(--wb-text)]">
                    {readString(match["kind"]) === "sheet"
                      ? `Sheet ${readString(match["sheetName"])}`
                      : `${readString(match["sheetName"])}!${readString(match["address"])}`}
                  </div>
                  <div className="mt-1 break-all text-[11px] leading-5 text-[var(--wb-text-subtle)]">
                    {readString(match["snippet"])}
                  </div>
                </div>
                <div className="text-[10px] text-[var(--wb-text-subtle)]">
                  score {readNumber(match["score"])}
                </div>
              </div>
              {Array.isArray(match["reasons"]) ? (
                <div className="mt-2 flex flex-wrap gap-1.5">
                  {match["reasons"].map((reason) => (
                    <span
                      key={readString(reason)}
                      className="rounded-full bg-[#e0f2fe] px-2 py-1 text-[10px] font-medium uppercase tracking-[0.04em] text-[#075985]"
                    >
                      {renderReasonLabel(readString(reason))}
                    </span>
                  ))}
                </div>
              ) : null}
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (props.toolName === "bilig.trace_dependencies" && Array.isArray(parsed["layers"])) {
    const root = isRecord(parsed["root"]) ? parsed["root"] : null;
    const layers = parsed["layers"].flatMap((layer) => (isRecord(layer) ? [layer] : []));
    return (
      <div className="mt-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-3">
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
              Dependency Trace
            </div>
            <div className="mt-1 text-[12px] text-[var(--wb-text-muted)]">
              {readString(root?.["sheetName"])}!{readString(root?.["address"])}
            </div>
          </div>
          <div className="text-[10px] text-[var(--wb-text-subtle)]">
            {readString(parsed["direction"], "both")} · {readNumber(parsed["depth"])} hops
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {layers.map((layer) => (
            <div
              key={`trace-layer-${readNumber(layer["depth"])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
                Hop {readNumber(layer["depth"])}
              </div>
              <div className="mt-2 grid gap-2 md:grid-cols-2">
                <div>
                  <div className="text-[11px] font-semibold text-[var(--wb-text)]">Precedents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer["precedents"]) && layer["precedents"].length > 0 ? (
                      layer["precedents"]
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`precedent:${readString(node["sheetName"])}:${readString(node["address"])}`}
                            className="text-[11px] leading-5 text-[var(--wb-text-subtle)]"
                          >
                            {readString(node["sheetName"])}!{readString(node["address"])}{" "}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node["formula"]) || readString(node["valueText"])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className="text-[11px] text-[var(--wb-text-subtle)]">None</div>
                    )}
                  </div>
                </div>
                <div>
                  <div className="text-[11px] font-semibold text-[var(--wb-text)]">Dependents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer["dependents"]) && layer["dependents"].length > 0 ? (
                      layer["dependents"]
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`dependent:${readString(node["sheetName"])}:${readString(node["address"])}`}
                            className="text-[11px] leading-5 text-[var(--wb-text-subtle)]"
                          >
                            {readString(node["sheetName"])}!{readString(node["address"])}{" "}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node["formula"]) || readString(node["valueText"])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className="text-[11px] text-[var(--wb-text-subtle)]">None</div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  return null;
}

function WorkbookAgentEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const { entry } = props;
  if (entry.kind === "user") {
    return (
      <div className="flex justify-end">
        <div className="max-w-[90%] rounded-[var(--wb-radius-control)] bg-[var(--wb-accent-soft)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]">
          {entry.text}
        </div>
      </div>
    );
  }

  if (entry.kind === "assistant") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]">
        {entry.text?.trim().length ? entry.text : "Thinking..."}
      </div>
    );
  }

  if (entry.kind === "plan") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2">
        <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          Plan
        </div>
        <div className="mt-1 whitespace-pre-wrap text-[12px] leading-5 text-[var(--wb-text-muted)]">
          {entry.text?.trim().length ? entry.text : "Planning..."}
        </div>
      </div>
    );
  }

  if (entry.kind === "tool") {
    return (
      <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0">
            <div className="text-[12px] font-semibold text-[var(--wb-text)]">{entry.toolName}</div>
            <div className="text-[11px] text-[var(--wb-text-subtle)]">Workbook tool call</div>
          </div>
          <ToolStatusPill status={entry.toolStatus} />
        </div>
        {entry.argumentsText ? (
          <pre className="mt-2 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-app-bg)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
            {entry.argumentsText}
          </pre>
        ) : null}
        <StructuredToolOutput toolName={entry.toolName} outputText={entry.outputText} />
        {entry.outputText &&
        (!supportsStructuredToolOutput(entry.toolName) ||
          safeParseToolOutput(entry.outputText) === null) ? (
          <pre className="mt-2 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
            {entry.outputText}
          </pre>
        ) : null}
      </div>
    );
  }

  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[#f1b5b5] bg-[#fff7f7] px-3 py-2 text-[12px] leading-5 text-[#991b1b]">
      {entry.text}
    </div>
  );
}

function PreviewRangeList(props: {
  readonly ranges: readonly {
    sheetName: string;
    startAddress: string;
    endAddress: string;
    role: "target" | "source";
  }[];
}) {
  if (props.ranges.length === 0) {
    return null;
  }
  return (
    <div className="mt-2 flex flex-wrap gap-1.5">
      {props.ranges.map((range) => (
        <span
          key={`${range.role}:${range.sheetName}:${range.startAddress}:${range.endAddress}`}
          className={cn(
            "inline-flex items-center rounded-full px-2 py-1 text-[10px] font-medium",
            range.role === "target" ? "bg-[#e0f2fe] text-[#0c4a6e]" : "bg-[#f1f5f9] text-[#475569]",
          )}
        >
          {range.role === "target" ? "Target" : "Source"} {range.sheetName}!{range.startAddress}
          {range.startAddress === range.endAddress ? "" : `:${range.endAddress}`}
        </span>
      ))}
    </div>
  );
}

function PendingBundleCard(props: {
  readonly bundle: WorkbookAgentCommandBundle;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly selectedCommandIndexes: readonly number[];
  readonly isApplyingBundle: boolean;
  readonly onApply: () => void;
  readonly onDismiss: () => void;
  readonly onSelectAll: () => void;
  readonly onToggleCommand: (commandIndex: number) => void;
}) {
  const selectedCount = props.selectedCommandIndexes.length;
  const hasFullSelection = selectedCount === props.bundle.commands.length;
  const canApply = props.preview !== null && !props.isApplyingBundle && selectedCount > 0;
  const applyLabel =
    selectedCount > 0 && !hasFullSelection
      ? props.isApplyingBundle
        ? "Applying Selected..."
        : "Apply Selected"
      : props.bundle.approvalMode === "explicit"
        ? props.isApplyingBundle
          ? "Approving..."
          : "Approve and Apply"
        : props.bundle.approvalMode === "auto"
          ? props.isApplyingBundle
            ? "Auto-Applying..."
            : "Apply Now"
          : props.isApplyingBundle
            ? "Applying..."
            : "Apply Preview";
  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-accent)]">
            Pending Preview
          </div>
          <div className="mt-1 text-[13px] font-semibold text-[var(--wb-text)]">
            {props.bundle.summary}
          </div>
        </div>
        <span className="rounded-full bg-[var(--wb-accent-soft)] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-accent)]">
          {props.bundle.riskClass}
        </span>
      </div>
      <div className="mt-2 text-[12px] leading-5 text-[var(--wb-text-subtle)]">
        Scope: {props.bundle.scope}. Base revision: r{String(props.bundle.baseRevision)}.
        {props.bundle.estimatedAffectedCells === null
          ? ""
          : ` ${String(props.bundle.estimatedAffectedCells)} affected cell${
              props.bundle.estimatedAffectedCells === 1 ? "" : "s"
            }.`}
      </div>
      <div className="mt-2 text-[11px] font-medium text-[var(--wb-text-subtle)]">
        Approval:{" "}
        {props.bundle.approvalMode === "auto"
          ? "auto-apply after local preview"
          : props.bundle.approvalMode === "explicit"
            ? "explicit approval required"
            : "preview required before apply"}
      </div>
      <div className="mt-3 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2 py-2">
        <div className="flex items-center justify-between gap-3">
          <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
            Selected {String(selectedCount)} of {String(props.bundle.commands.length)} changes
          </div>
          {!hasFullSelection ? (
            <button
              className="text-[11px] font-medium text-[var(--wb-accent)] transition-colors hover:brightness-[0.95] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
              type="button"
              onClick={props.onSelectAll}
            >
              Select all
            </button>
          ) : null}
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {props.bundle.commands.map((command, index) => {
            const checked = props.selectedCommandIndexes.includes(index);
            const commandLabel = describeWorkbookAgentCommand(command);
            return (
              <div
                key={`${props.bundle.id}:${JSON.stringify(command)}`}
                className={cn(
                  "flex items-start gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2 transition-colors",
                  checked
                    ? "border-[var(--wb-accent-ring)] bg-[var(--wb-surface)]"
                    : "border-[var(--wb-border)] bg-[var(--wb-surface)]",
                )}
              >
                <input
                  aria-label={`Toggle staged workbook change ${String(index + 1)}: ${commandLabel}`}
                  checked={checked}
                  className="mt-0.5 h-4 w-4 rounded border-[var(--wb-border)] text-[var(--wb-accent)] focus:ring-[var(--wb-accent-ring)]"
                  data-testid={`workbook-agent-command-toggle-${String(index)}`}
                  type="checkbox"
                  onChange={() => {
                    props.onToggleCommand(index);
                  }}
                />
                <div className="min-w-0">
                  <div className="text-[11px] font-semibold text-[var(--wb-text)]">
                    Change {String(index + 1)}
                  </div>
                  <div className="mt-1 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
                    {commandLabel}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
      {props.preview ? (
        <div className="mt-2 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
          Preview effects: {String(props.preview.effectSummary.displayedCellDiffCount)} sampled cell
          diff{props.preview.effectSummary.displayedCellDiffCount === 1 ? "" : "s"} ·{" "}
          {String(props.preview.effectSummary.formulaChangeCount)} formulas ·{" "}
          {String(props.preview.effectSummary.inputChangeCount)} values ·{" "}
          {String(props.preview.effectSummary.styleChangeCount)} styles ·{" "}
          {String(props.preview.effectSummary.numberFormatChangeCount)} number formats
          {props.preview.effectSummary.truncatedCellDiffs ? " · diff list truncated" : ""}
        </div>
      ) : selectedCount === 0 ? (
        <div className="mt-2 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
          Select at least one change to preview and apply.
        </div>
      ) : null}
      <PreviewRangeList ranges={props.preview?.ranges ?? props.bundle.affectedRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div className="mt-2 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
          {props.preview.structuralChanges.join(" · ")}
        </div>
      ) : null}
      {props.preview?.cellDiffs?.length ? (
        <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
          <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2 py-1 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
            Local Preview Diff
          </div>
          <div className="max-h-44 overflow-y-auto">
            {props.preview.cellDiffs.map((diff) => (
              <div
                key={`${diff.sheetName}:${diff.address}`}
                className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-x-2 border-t border-[var(--wb-border)] px-2 py-2 text-[11px] leading-5 first:border-t-0"
              >
                <div className="col-span-2 font-medium text-[var(--wb-text)]">
                  {diff.sheetName}!{diff.address}
                </div>
                <div className="col-span-2 mt-1 flex flex-wrap gap-1">
                  {diff.changeKinds.map((kind) => (
                    <span
                      key={kind}
                      className="rounded-full bg-[var(--wb-surface-subtle)] px-2 py-0.5 text-[10px] font-medium uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]"
                    >
                      {renderPreviewChangeKind(kind)}
                    </span>
                  ))}
                </div>
                <div className="text-[var(--wb-text-subtle)]">
                  Before: {(diff.beforeFormula ?? String(diff.beforeInput ?? "")) || "(empty)"}
                </div>
                <div className="text-[var(--wb-text)]">
                  After: {(diff.afterFormula ?? String(diff.afterInput ?? "")) || "(empty)"}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onDismiss}
        >
          Dismiss
        </button>
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:brightness-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
          data-testid="workbook-agent-apply-pending"
          disabled={!canApply}
          type="button"
          onClick={props.onApply}
        >
          {applyLabel}
        </button>
      </div>
    </div>
  );
}

function ExecutionRecordRow(props: {
  readonly record: WorkbookAgentExecutionRecord;
  readonly onReplay: () => void;
}) {
  return (
    <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="text-[12px] font-semibold text-[var(--wb-text)]">
            {props.record.summary}
          </div>
          <div className="text-[11px] text-[var(--wb-text-subtle)]">
            {props.record.appliedBy === "auto" ? "Auto-applied" : "Applied"} at r
            {String(props.record.appliedRevision)} · {props.record.riskClass} risk ·{" "}
            {props.record.approvalMode} · {props.record.acceptedScope}
          </div>
        </div>
        <span className="rounded-full bg-[var(--wb-surface-subtle)] px-2 py-1 text-[10px] font-medium uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          {props.record.scope}
        </span>
      </div>
      {(props.record.planText ?? props.record.goalText).trim().length > 0 ? (
        <div className="mt-2 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
          {props.record.planText ?? props.record.goalText}
        </div>
      ) : null}
      {props.record.preview ? (
        <div className="mt-2 text-[11px] leading-5 text-[var(--wb-text-subtle)]">
          Preview effects: {String(props.record.preview.effectSummary.displayedCellDiffCount)}{" "}
          sampled cell diff
          {props.record.preview.effectSummary.displayedCellDiffCount === 1 ? "" : "s"} ·{" "}
          {String(props.record.preview.effectSummary.structuralChangeCount)} structural changes
        </div>
      ) : null}
      <PreviewRangeList ranges={props.record.preview?.ranges ?? []} />
      <div className="mt-3 flex items-center justify-end">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onReplay}
        >
          Replay as Preview
        </button>
      </div>
    </div>
  );
}

export function WorkbookAgentPanel(props: {
  readonly currentContext: WorkbookAgentUiContext | null;
  readonly snapshot: WorkbookAgentSessionSnapshot | null;
  readonly pendingBundle: WorkbookAgentCommandBundle | null;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly selectedCommandIndexes: readonly number[];
  readonly executionRecords: readonly WorkbookAgentExecutionRecord[];
  readonly draft: string;
  readonly error: string | null;
  readonly isLoading: boolean;
  readonly isApplyingBundle: boolean;
  readonly isOpen: boolean;
  readonly onApplyPendingBundle: () => void;
  readonly onClose: () => void;
  readonly onDraftChange: (value: string) => void;
  readonly onDismissPendingBundle: () => void;
  readonly onInterrupt: () => void;
  readonly onSelectAllPendingCommands: () => void;
  readonly onTogglePendingCommand: (commandIndex: number) => void;
  readonly onReplayExecutionRecord: (recordId: string) => void;
  readonly onSubmit: () => void;
  readonly onUseSkillPrompt: (prompt: string) => void;
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    const node = scrollRef.current;
    if (!node) {
      return;
    }
    node.scrollTop = node.scrollHeight;
  }, [props.snapshot?.entries.length, props.snapshot?.status]);

  if (!props.isOpen) {
    return null;
  }

  const isRunning = props.snapshot?.status === "inProgress";

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Assistant</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            {contextLabel(props.snapshot?.context ?? props.currentContext)}
          </p>
        </div>
        <button
          aria-label="Close assistant"
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          <span aria-hidden="true">×</span>
        </button>
      </div>
      <div ref={scrollRef} className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.pendingBundle ? (
          <div className="mb-3">
            <PendingBundleCard
              bundle={props.pendingBundle}
              preview={props.preview}
              selectedCommandIndexes={props.selectedCommandIndexes}
              isApplyingBundle={props.isApplyingBundle}
              onApply={props.onApplyPendingBundle}
              onDismiss={props.onDismissPendingBundle}
              onSelectAll={props.onSelectAllPendingCommands}
              onToggleCommand={props.onTogglePendingCommand}
            />
          </div>
        ) : null}
        <div className="mb-3">
          <WorkbookAgentSkillStrip
            onUseSkillPrompt={(prompt) => {
              props.onUseSkillPrompt(prompt);
              window.requestAnimationFrame(() => {
                textareaRef.current?.focus();
              });
            }}
          />
        </div>
        {props.isLoading ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Starting workbook assistant...
          </div>
        ) : props.snapshot && props.snapshot.entries.length > 0 ? (
          <div className="flex flex-col gap-2">
            {props.snapshot.entries.map((entry) => (
              <WorkbookAgentEntryRow key={entry.id} entry={entry} />
            ))}
            {props.executionRecords.length > 0 ? (
              <div className="pt-2">
                <div className="mb-2 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
                  Applied Plans
                </div>
                <div className="flex flex-col gap-2">
                  {props.executionRecords.slice(0, 5).map((record) => (
                    <ExecutionRecordRow
                      key={record.id}
                      record={record}
                      onReplay={() => {
                        props.onReplayExecutionRecord(record.id);
                      }}
                    />
                  ))}
                </div>
              </div>
            ) : null}
          </div>
        ) : (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Ask the assistant to inspect, edit, or restructure this workbook.
          </div>
        )}
      </div>
      {props.error ? (
        <div className="border-t border-[#f1b5b5] bg-[#fff7f7] px-4 py-2 text-[12px] text-[#991b1b]">
          {props.error}
        </div>
      ) : null}
      <form
        className="border-t border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3"
        onSubmit={(event) => {
          event.preventDefault();
          props.onSubmit();
        }}
      >
        <label className="sr-only" htmlFor="workbook-agent-input">
          Ask the workbook assistant
        </label>
        <textarea
          id="workbook-agent-input"
          className="min-h-24 w-full resize-none rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
          data-testid="workbook-agent-input"
          ref={textareaRef}
          placeholder="Ask the assistant to update this workbook..."
          value={props.draft}
          onChange={(event) => {
            props.onDraftChange(event.target.value);
          }}
        />
        <div className="mt-3 flex items-center justify-between gap-2">
          <div className="text-[11px] text-[var(--wb-text-subtle)]">
            Uses local workbook tools through the monolith agent runtime.
          </div>
          <div className="flex items-center gap-2">
            {isRunning ? (
              <button
                className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f1b5b5] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1"
                data-testid="workbook-agent-interrupt"
                type="button"
                onClick={props.onInterrupt}
              >
                Stop
              </button>
            ) : null}
            <button
              className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:brightness-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
              data-testid="workbook-agent-send"
              disabled={props.draft.trim().length === 0 || props.isLoading}
              type="submit"
            >
              Send
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
