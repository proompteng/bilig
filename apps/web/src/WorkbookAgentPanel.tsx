import { useEffect, useRef } from "react";
import {
  WORKBOOK_AGENT_TOOL_NAMES,
  describeWorkbookAgentCommand,
  normalizeWorkbookAgentToolName,
} from "@bilig/agent-api";
import { cva } from "class-variance-authority";
import type {
  WorkbookAgentCommandBundle,
  WorkbookAgentPreviewChangeKind,
  WorkbookAgentExecutionRecord,
  WorkbookAgentPreviewSummary,
  WorkbookAgentSharedReviewRecommendation,
} from "@bilig/agent-api";
import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentTimelineCitation,
  WorkbookAgentThreadScope,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
} from "@bilig/contracts";
import { cn } from "./cn.js";
import {
  workbookAlertClass,
  workbookButtonClass,
  workbookInsetClass,
  workbookPillClass,
  workbookSurfaceClass,
} from "./workbook-shell-chrome.js";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";
import { WorkflowActionStrip } from "./workbook-agent-panel-workflow-actions.js";
import {
  ExecutionRecordRow,
  PreviewRangeList,
  WorkflowRunRow,
} from "./workbook-agent-panel-history.js";

const toolStatusPillClass = cva(
  "inline-flex h-5 items-center rounded-full border px-2 text-[10px] font-semibold uppercase tracking-[0.04em]",
  {
    variants: {
      status: {
        completed: workbookPillClass({ tone: "neutral", weight: "strong" }),
        failed: workbookPillClass({ tone: "danger", weight: "strong" }),
        running: workbookPillClass({ tone: "accent", weight: "strong" }),
      },
    },
  },
);

function contextLabel(context: WorkbookAgentUiContext | null): string {
  if (!context) {
    return "";
  }
  return `${context.selection.sheetName}!${context.selection.address}`;
}

function formatThreadEntryCount(entryCount: number): string {
  return `${entryCount} ${entryCount === 1 ? "item" : "items"}`;
}

function summarizeThreadActivity(text: string | null): string | null {
  if (!text) {
    return null;
  }
  const normalized = text.trim().replaceAll(/\s+/g, " ");
  if (normalized.length === 0) {
    return null;
  }
  return normalized.length <= 64 ? normalized : `${normalized.slice(0, 61)}...`;
}

function ThreadSummaryStrip(props: {
  readonly activeThreadId: string | null;
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[];
  readonly onSelectThread: (threadId: string) => void;
}) {
  if (props.threadSummaries.length === 0) {
    return null;
  }

  return (
    <div className="mt-2 flex gap-1.5 overflow-x-auto pb-0.5">
      {props.threadSummaries.map((threadSummary) => {
        const isActive = threadSummary.threadId === props.activeThreadId;
        const latestActivity = summarizeThreadActivity(threadSummary.latestEntryText);
        return (
          <button
            key={threadSummary.threadId}
            aria-label={`Open ${threadSummary.scope} thread ${threadSummary.threadId}`}
            aria-pressed={isActive}
            className={cn(
              workbookButtonClass({
                size: "sm",
                tone: isActive ? "accent" : "neutral",
                weight: isActive ? "strong" : "regular",
              }),
              "h-auto min-w-[8.5rem] shrink-0 flex-col items-start gap-1 px-3 py-2 text-left",
            )}
            data-testid={`workbook-agent-thread-${threadSummary.threadId}`}
            type="button"
            onClick={() => {
              props.onSelectThread(threadSummary.threadId);
            }}
          >
            <div className="flex w-full items-center justify-between gap-2">
              <span className="text-[11px] font-semibold text-[var(--wb-text)]">
                {threadSummary.scope === "shared" ? "Shared" : "Private"}
              </span>
              {threadSummary.hasPendingBundle ? (
                <span className={workbookPillClass({ tone: "accent", weight: "strong" })}>
                  Pending
                </span>
              ) : null}
            </div>
            <div className="text-[11px] text-[var(--wb-text-subtle)]">
              {threadSummary.scope === "shared"
                ? `${formatWorkbookCollaboratorLabel(threadSummary.ownerUserId)} · ${formatThreadEntryCount(threadSummary.entryCount)}`
                : formatThreadEntryCount(threadSummary.entryCount)}
            </div>
            {latestActivity ? (
              <div className="line-clamp-2 text-[11px] leading-4 text-[var(--wb-text-muted)]">
                {latestActivity}
              </div>
            ) : null}
          </button>
        );
      })}
    </div>
  );
}

function ThreadComposerControls(props: {
  readonly threadScope: WorkbookAgentThreadScope;
  readonly onSelectThreadScope: (scope: WorkbookAgentThreadScope) => void;
  readonly onStartNewThread: () => void;
}) {
  return (
    <div className="mt-2 flex items-center justify-between gap-2">
      <div className="flex items-center gap-1">
        {(["private", "shared"] as const).map((scope) => {
          const isActive = props.threadScope === scope;
          return (
            <button
              key={scope}
              aria-pressed={isActive}
              className={cn(
                workbookButtonClass({
                  size: "sm",
                  tone: isActive ? "accent" : "neutral",
                  weight: isActive ? "strong" : "regular",
                }),
                "px-2.5",
              )}
              data-testid={`workbook-agent-scope-${scope}`}
              type="button"
              onClick={() => {
                props.onSelectThreadScope(scope);
              }}
            >
              {scope === "shared" ? "Shared" : "Private"}
            </button>
          );
        })}
      </div>
      <button
        className={workbookButtonClass({ size: "sm", tone: "neutral", weight: "strong" })}
        data-testid="workbook-agent-new-thread"
        type="button"
        onClick={props.onStartNewThread}
      >
        New thread
      </button>
    </div>
  );
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry["toolStatus"] }) {
  const label =
    props.status === "completed" ? "Done" : props.status === "failed" ? "Failed" : "Running";
  return (
    <span
      className={toolStatusPillClass({
        status:
          props.status === "completed"
            ? "completed"
            : props.status === "failed"
              ? "failed"
              : "running",
      })}
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
  const normalizedToolName = toolName ? normalizeWorkbookAgentToolName(toolName) : null;
  return (
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues ||
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook ||
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.traceDependencies
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
  const normalizedToolName = props.toolName ? normalizeWorkbookAgentToolName(props.toolName) : null;
  if (!normalizedToolName || !isRecord(parsed)) {
    return null;
  }

  if (
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues &&
    Array.isArray(parsed["issues"])
  ) {
    const summary = isRecord(parsed["summary"]) ? parsed["summary"] : null;
    const issues = parsed["issues"].flatMap((issue) => (isRecord(issue) ? [issue] : []));
    return (
      <div className={cn(workbookInsetClass(), "mt-2 px-3 py-3")}>
        <div className="flex items-start justify-between gap-3 text-[11px] text-[var(--wb-text-muted)]">
          <div>
            {readNumber(summary?.["issueCount"], issues.length)} issues ·{" "}
            {readNumber(summary?.["scannedFormulaCells"])} formulas
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
                        className={workbookPillClass({ tone: "danger", weight: "strong" })}
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

  if (
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook &&
    Array.isArray(parsed["matches"])
  ) {
    const matches = parsed["matches"].flatMap((match) => (isRecord(match) ? [match] : []));
    return (
      <div className={cn(workbookInsetClass(), "mt-2 px-3 py-3")}>
        <div className="flex items-start justify-between gap-3 text-[11px] text-[var(--wb-text-muted)]">
          <div className="truncate">“{readString(parsed["query"])}”</div>
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
                      className={workbookPillClass({ tone: "neutral" })}
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

  if (
    normalizedToolName === WORKBOOK_AGENT_TOOL_NAMES.traceDependencies &&
    Array.isArray(parsed["layers"])
  ) {
    const root = isRecord(parsed["root"]) ? parsed["root"] : null;
    const layers = parsed["layers"].flatMap((layer) => (isRecord(layer) ? [layer] : []));
    return (
      <div className={cn(workbookInsetClass(), "mt-2 px-3 py-3")}>
        <div className="flex items-start justify-between gap-3 text-[11px] text-[var(--wb-text-muted)]">
          <div>
            {readString(root?.["sheetName"])}!{readString(root?.["address"])}
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
        <div className="max-w-[90%] rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)] px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]">
          {entry.text}
          <TimelineCitationList citations={entry.citations} />
        </div>
      </div>
    );
  }

  if (entry.kind === "assistant") {
    if (!entry.text?.trim().length) {
      return null;
    }
    return (
      <div
        className={cn(
          workbookInsetClass(),
          "px-3 py-2 text-[13px] leading-5 text-[var(--wb-text)]",
        )}
      >
        {entry.text}
        <TimelineCitationList citations={entry.citations} />
      </div>
    );
  }

  if (entry.kind === "plan") {
    if (!entry.text?.trim().length) {
      return null;
    }
    return (
      <div className={cn(workbookInsetClass(), "px-3 py-2")}>
        <div className="whitespace-pre-wrap text-[12px] leading-5 text-[var(--wb-text-muted)]">
          {entry.text}
        </div>
        <TimelineCitationList citations={entry.citations} />
      </div>
    );
  }

  if (entry.kind === "tool") {
    return (
      <div className={cn(workbookInsetClass(), "px-3 py-2")}>
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0">
            <div className="text-[11px] font-medium text-[var(--wb-text-muted)]">
              {entry.toolName}
            </div>
          </div>
          <ToolStatusPill status={entry.toolStatus} />
        </div>
        <StructuredToolOutput toolName={entry.toolName} outputText={entry.outputText} />
        {entry.outputText &&
        (!supportsStructuredToolOutput(entry.toolName) ||
          safeParseToolOutput(entry.outputText) === null) ? (
          <pre className="mt-2 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]">
            {entry.outputText}
          </pre>
        ) : null}
        <TimelineCitationList citations={entry.citations} />
      </div>
    );
  }

  return (
    <div className={cn(workbookInsetClass(), "px-3 py-2")}>
      <div className="text-[11px] leading-5 text-[var(--wb-text-muted)]">{entry.text}</div>
      <TimelineCitationList citations={entry.citations} />
    </div>
  );
}

function TimelineCitationList(props: {
  readonly citations: readonly WorkbookAgentTimelineCitation[];
}) {
  if (props.citations.length === 0) {
    return null;
  }
  const occurrenceCounts = new Map<string, number>();
  return (
    <div className="mt-2 flex flex-wrap gap-1.5">
      {props.citations.map((citation) => {
        if (citation.kind === "revision") {
          const baseKey = `revision:${String(citation.revision)}`;
          const occurrence = (occurrenceCounts.get(baseKey) ?? 0) + 1;
          occurrenceCounts.set(baseKey, occurrence);
          return (
            <span
              key={`${baseKey}:${String(occurrence)}`}
              className={workbookPillClass({ tone: "neutral" })}
            >
              r{String(citation.revision)}
            </span>
          );
        }
        const baseKey = `${citation.kind}:${citation.role}:${citation.sheetName}:${citation.startAddress}:${citation.endAddress}`;
        const occurrence = (occurrenceCounts.get(baseKey) ?? 0) + 1;
        occurrenceCounts.set(baseKey, occurrence);
        return (
          <span
            key={`${baseKey}:${String(occurrence)}`}
            className={workbookPillClass({
              tone: citation.role === "target" ? "accent" : "neutral",
            })}
          >
            {citation.sheetName}!{citation.startAddress}
            {citation.startAddress === citation.endAddress ? "" : `:${citation.endAddress}`}
          </span>
        );
      })}
    </div>
  );
}

function PendingBundleCard(props: {
  readonly bundle: WorkbookAgentCommandBundle;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly sharedApprovalOwnerUserId: string | null;
  readonly sharedReviewOwnerUserId: string | null;
  readonly sharedReviewStatus: "pending" | "approved" | "rejected" | null;
  readonly sharedReviewDecidedByUserId: string | null;
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[];
  readonly currentUserSharedRecommendation: "approved" | "rejected" | null;
  readonly canFinalizeSharedBundle: boolean;
  readonly canRecommendSharedBundle: boolean;
  readonly selectedCommandIndexes: readonly number[];
  readonly isApplyingBundle: boolean;
  readonly onApply: () => void;
  readonly onDismiss: () => void;
  readonly onReview: (decision: "approved" | "rejected") => void;
  readonly onSelectAll: () => void;
  readonly onToggleCommand: (commandIndex: number) => void;
}) {
  const selectedCount = props.selectedCommandIndexes.length;
  const hasFullSelection = selectedCount === props.bundle.commands.length;
  const sharedApprovalOwnerLabel = props.sharedApprovalOwnerUserId
    ? formatWorkbookCollaboratorLabel(props.sharedApprovalOwnerUserId)
    : null;
  const sharedReviewOwnerLabel = props.sharedReviewOwnerUserId
    ? formatWorkbookCollaboratorLabel(props.sharedReviewOwnerUserId)
    : null;
  const sharedReviewDecisionLabel = props.sharedReviewDecidedByUserId
    ? formatWorkbookCollaboratorLabel(props.sharedReviewDecidedByUserId)
    : null;
  const approvalRecommendationCount = props.sharedReviewRecommendations.filter(
    (recommendation) => recommendation.decision === "approved",
  ).length;
  const rejectionRecommendationCount = props.sharedReviewRecommendations.filter(
    (recommendation) => recommendation.decision === "rejected",
  ).length;
  const recommendationSummary =
    props.sharedReviewRecommendations.length === 0
      ? null
      : `${String(approvalRecommendationCount)} approval ${approvalRecommendationCount === 1 ? "recommendation" : "recommendations"} · ${String(rejectionRecommendationCount)} rejection ${rejectionRecommendationCount === 1 ? "recommendation" : "recommendations"}`;
  const canApply =
    props.preview !== null &&
    !props.isApplyingBundle &&
    selectedCount > 0 &&
    sharedApprovalOwnerLabel === null &&
    (props.sharedReviewStatus === null || props.sharedReviewStatus === "approved");
  const applyLabel =
    selectedCount > 0 && !hasFullSelection
      ? props.isApplyingBundle
        ? "Applying…"
        : "Apply"
      : props.sharedReviewStatus === "pending"
        ? "Awaiting approval"
        : props.sharedReviewStatus === "rejected"
          ? "Rejected"
          : props.bundle.approvalMode === "explicit"
            ? props.isApplyingBundle
              ? "Applying…"
              : "Approve"
            : props.isApplyingBundle
              ? "Applying…"
              : "Apply";
  return (
    <div
      className={cn(
        workbookSurfaceClass({ emphasis: "raised" }),
        "border-[var(--wb-border-strong)] px-3 py-3",
      )}
    >
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-[13px] font-semibold text-[var(--wb-text)]">
            {props.bundle.summary}
          </div>
        </div>
        <span className={workbookPillClass({ tone: "accent", weight: "strong" })}>
          {props.bundle.riskClass}
        </span>
      </div>
      <div className={cn(workbookInsetClass(), "mt-3 px-2 py-2")}>
        <div className="flex items-center justify-between gap-3">
          <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
            {String(selectedCount)}/{String(props.bundle.commands.length)}
          </div>
          {!hasFullSelection ? (
            <button
              className="text-[11px] font-medium text-[var(--wb-accent)] transition-colors hover:brightness-[0.95] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
              type="button"
              onClick={props.onSelectAll}
            >
              All
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
                  <div className="text-[11px] leading-5 text-[var(--wb-text)]">{commandLabel}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
      <PreviewRangeList ranges={props.preview?.ranges ?? props.bundle.affectedRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div
          className={cn(
            workbookInsetClass(),
            "mt-2 border-transparent px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]",
          )}
        >
          {props.preview.structuralChanges.join(" · ")}
        </div>
      ) : null}
      {sharedApprovalOwnerLabel ? (
        <div
          className={cn(
            workbookAlertClass({ tone: "warning" }),
            "mt-2 border-[var(--wb-border-strong)] text-[11px] leading-5",
          )}
        >
          Only {sharedApprovalOwnerLabel} can approve medium/high-risk changes on this shared
          thread.
        </div>
      ) : null}
      {recommendationSummary ? (
        <div
          className={cn(
            workbookInsetClass(),
            "mt-2 border-transparent px-2 py-2 text-[11px] leading-5 text-[var(--wb-text-muted)]",
          )}
        >
          {recommendationSummary}
          {props.currentUserSharedRecommendation
            ? ` You recommended ${props.currentUserSharedRecommendation === "approved" ? "approval" : "rejection"}.`
            : ""}
        </div>
      ) : null}
      {sharedReviewOwnerLabel && props.sharedReviewStatus ? (
        <div
          className={cn(
            workbookAlertClass({
              tone:
                props.sharedReviewStatus === "rejected" ? "danger" : "warning",
            }),
            "mt-2 border-[var(--wb-border-strong)] text-[11px] leading-5",
          )}
        >
          {props.sharedReviewStatus === "pending"
            ? `Awaiting ${sharedReviewOwnerLabel}'s approval before this shared bundle can be applied.`
            : props.sharedReviewStatus === "approved"
              ? `Approved by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`
              : `Rejected by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`}
        </div>
      ) : null}
      {props.canFinalizeSharedBundle && props.sharedReviewStatus !== null ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <button
            className={workbookButtonClass({ tone: "neutral" })}
            data-testid="workbook-agent-review-reject"
            type="button"
            onClick={() => {
              props.onReview("rejected");
            }}
          >
            Reject
          </button>
          <button
            className={workbookButtonClass({ tone: "accent" })}
            data-testid="workbook-agent-review-approve"
            type="button"
            onClick={() => {
              props.onReview("approved");
            }}
          >
            Approve
          </button>
        </div>
      ) : null}
      {props.canRecommendSharedBundle && props.sharedReviewStatus === "pending" ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <button
            className={workbookButtonClass({ tone: "neutral" })}
            data-testid="workbook-agent-review-reject"
            type="button"
            onClick={() => {
              props.onReview("rejected");
            }}
          >
            Recommend reject
          </button>
          <button
            className={workbookButtonClass({ tone: "accent" })}
            data-testid="workbook-agent-review-approve"
            type="button"
            onClick={() => {
              props.onReview("approved");
            }}
          >
            Recommend approve
          </button>
        </div>
      ) : null}
      {props.preview?.cellDiffs?.length ? (
        <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
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
                    <span key={kind} className={workbookPillClass({ tone: "neutral" })}>
                      {renderPreviewChangeKind(kind)}
                    </span>
                  ))}
                </div>
                <div className="text-[var(--wb-text-subtle)]">
                  {(diff.beforeFormula ?? String(diff.beforeInput ?? "")) || "(empty)"}
                </div>
                <div className="text-[var(--wb-text)]">
                  {(diff.afterFormula ?? String(diff.afterInput ?? "")) || "(empty)"}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <button
          className={workbookButtonClass({ tone: "neutral" })}
          type="button"
          onClick={props.onDismiss}
        >
          Clear
        </button>
        <button
          className={workbookButtonClass({ tone: "accent", weight: "strong" })}
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

export function WorkbookAgentPanel(props: {
  readonly activeThreadId: string | null;
  readonly currentContext: WorkbookAgentUiContext | null;
  readonly snapshot: WorkbookAgentSessionSnapshot | null;
  readonly pendingBundle: WorkbookAgentCommandBundle | null;
  readonly preview: WorkbookAgentPreviewSummary | null;
  readonly sharedApprovalOwnerUserId: string | null;
  readonly sharedReviewOwnerUserId: string | null;
  readonly sharedReviewStatus: "pending" | "approved" | "rejected" | null;
  readonly sharedReviewDecidedByUserId: string | null;
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[];
  readonly currentUserSharedRecommendation: "approved" | "rejected" | null;
  readonly canFinalizeSharedBundle: boolean;
  readonly canRecommendSharedBundle: boolean;
  readonly selectedCommandIndexes: readonly number[];
  readonly executionRecords: readonly WorkbookAgentExecutionRecord[];
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[];
  readonly cancellingWorkflowRunId: string | null;
  readonly isStartingWorkflow: boolean;
  readonly threadScope: WorkbookAgentThreadScope;
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[];
  readonly draft: string;
  readonly isLoading: boolean;
  readonly isApplyingBundle: boolean;
  readonly onApplyPendingBundle: () => void;
  readonly onDraftChange: (value: string) => void;
  readonly onDismissPendingBundle: () => void;
  readonly onReviewPendingBundle: (decision: "approved" | "rejected") => void;
  readonly onInterrupt: () => void;
  readonly onSelectAllPendingCommands: () => void;
  readonly onSelectThreadScope: (scope: WorkbookAgentThreadScope) => void;
  readonly onSelectThread: (threadId: string) => void;
  readonly onStartNewThread: () => void;
  readonly onTogglePendingCommand: (commandIndex: number) => void;
  readonly onCancelWorkflowRun: (runId: string) => void;
  readonly onReplayExecutionRecord: (recordId: string) => void;
  readonly onStartWorkflow: (
    template: Exclude<
      WorkbookAgentWorkflowRun["workflowTemplate"],
      "searchWorkbookQuery" | "createSheet" | "renameCurrentSheet"
    >,
  ) => void;
  readonly onStartNamedWorkflow: (
    template: Extract<WorkbookAgentWorkflowRun["workflowTemplate"], "createSheet" | "renameCurrentSheet">,
    name: string,
  ) => void;
  readonly onStartSearchWorkflow: (query: string) => void;
  readonly onStartStructuralWorkflow: (
    template: Extract<WorkbookAgentWorkflowRun["workflowTemplate"], "hideCurrentRow" | "hideCurrentColumn">,
  ) => void;
  readonly onSubmit: () => void;
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const node = scrollRef.current;
    if (!node) {
      return;
    }
    node.scrollTop = node.scrollHeight;
  }, [props.snapshot?.entries.length, props.snapshot?.status]);

  const isRunning = props.snapshot?.status === "inProgress";

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
    >
      <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2.5">
        <div className="min-w-0 text-[12px] font-medium text-[var(--wb-text)]">
          {contextLabel(props.snapshot?.context ?? props.currentContext)}
        </div>
        <ThreadComposerControls
          threadScope={props.threadScope}
          onSelectThreadScope={props.onSelectThreadScope}
          onStartNewThread={props.onStartNewThread}
        />
        <ThreadSummaryStrip
          activeThreadId={props.activeThreadId}
          threadSummaries={props.threadSummaries}
          onSelectThread={props.onSelectThread}
        />
        <WorkflowActionStrip
          disabled={props.isLoading || isRunning}
          isStartingWorkflow={props.isStartingWorkflow}
          onStartWorkflow={props.onStartWorkflow}
          onStartNamedWorkflow={props.onStartNamedWorkflow}
          onStartSearchWorkflow={props.onStartSearchWorkflow}
          onStartStructuralWorkflow={props.onStartStructuralWorkflow}
        />
      </div>
      <div
        ref={scrollRef}
        className="min-h-0 flex-1 overflow-y-auto bg-[var(--wb-app-bg)] px-2.5 py-2.5"
      >
        {props.pendingBundle ? (
          <div className="mb-3">
            <PendingBundleCard
              bundle={props.pendingBundle}
              preview={props.preview}
              sharedApprovalOwnerUserId={props.sharedApprovalOwnerUserId}
              sharedReviewOwnerUserId={props.sharedReviewOwnerUserId}
              sharedReviewStatus={props.sharedReviewStatus}
              sharedReviewDecidedByUserId={props.sharedReviewDecidedByUserId}
              sharedReviewRecommendations={props.sharedReviewRecommendations}
              currentUserSharedRecommendation={props.currentUserSharedRecommendation}
              canFinalizeSharedBundle={props.canFinalizeSharedBundle}
              canRecommendSharedBundle={props.canRecommendSharedBundle}
              selectedCommandIndexes={props.selectedCommandIndexes}
              isApplyingBundle={props.isApplyingBundle}
              onApply={props.onApplyPendingBundle}
              onDismiss={props.onDismissPendingBundle}
              onReview={props.onReviewPendingBundle}
              onSelectAll={props.onSelectAllPendingCommands}
              onToggleCommand={props.onTogglePendingCommand}
            />
          </div>
        ) : null}
        {props.isLoading ? null : props.snapshot && props.snapshot.entries.length > 0 ? (
          <div className="flex flex-col gap-1.5">
            {props.snapshot.entries.map((entry) => (
              <WorkbookAgentEntryRow key={entry.id} entry={entry} />
            ))}
            {props.workflowRuns.length > 0 ? (
              <div className="pt-2">
                <div className="mb-2 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
                  Workflows
                </div>
                <div className="flex flex-col gap-2">
                  {props.workflowRuns.slice(0, 5).map((run) => (
                    <WorkflowRunRow
                      key={run.runId}
                      isCancelling={props.cancellingWorkflowRunId === run.runId}
                      run={run}
                      onCancel={() => {
                        props.onCancelWorkflowRun(run.runId);
                      }}
                    />
                  ))}
                </div>
              </div>
            ) : null}
            {props.executionRecords.length > 0 ? (
              <div className="pt-2">
                <div className="mb-2 text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
                  Executions
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
        ) : null}
      </div>
      <form
        className="border-t border-[var(--wb-border)] bg-[var(--wb-surface)] px-2.5 py-2.5"
        onSubmit={(event) => {
          event.preventDefault();
          props.onSubmit();
        }}
      >
        <label className="sr-only" htmlFor="workbook-agent-input">
          Ask the workbook assistant
        </label>
        <div className="relative">
          <textarea
            id="workbook-agent-input"
            className="min-h-32 w-full resize-none rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3 pr-14 text-[13px] leading-5 text-[var(--wb-text)] placeholder:text-[var(--wb-text-subtle)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
            data-testid="workbook-agent-input"
            placeholder="Message"
            value={props.draft}
            onChange={(event) => {
              props.onDraftChange(event.target.value);
            }}
            onKeyDown={(event) => {
              if (isRunning) {
                return;
              }
              if (event.key !== "Enter" || event.shiftKey || event.nativeEvent.isComposing) {
                return;
              }
              event.preventDefault();
              props.onSubmit();
            }}
          />
          <button
            aria-label={isRunning ? "Stop" : "Send message"}
            className="absolute right-3 bottom-3 inline-flex h-9 w-9 items-center justify-center rounded-full border border-[#0f0f10] bg-[#0f0f10] text-[var(--color-mauve-200)] transition-colors hover:bg-[#18181b] hover:text-[var(--color-mauve-100)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:border-[var(--wb-border)] disabled:bg-[var(--wb-surface-muted)] disabled:text-[var(--wb-text-subtle)] disabled:opacity-70 disabled:shadow-none"
            data-testid="workbook-agent-send"
            disabled={!isRunning && (props.draft.trim().length === 0 || props.isLoading)}
            type="button"
            onClick={() => {
              if (isRunning) {
                props.onInterrupt();
                return;
              }
              props.onSubmit();
            }}
          >
            {isRunning ? <StopIcon /> : <SendArrowIcon />}
          </button>
        </div>
      </form>
    </div>
  );
}

function SendArrowIcon() {
  return (
    <svg aria-hidden="true" className="h-5 w-5" fill="none" viewBox="0 0 16 16">
      <path
        d="M8 12V4M8 4L4.75 7.25M8 4l3.25 3.25"
        stroke="currentColor"
        strokeLinecap="round"
        strokeLinejoin="round"
        strokeWidth="1.6"
      />
    </svg>
  );
}

function StopIcon() {
  return <div aria-hidden="true" className="h-3.5 w-3.5 rounded-[3px] bg-current" />;
}
