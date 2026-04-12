import { Fragment, type CSSProperties, useEffect, useRef } from "react";
import { Button } from "@base-ui/react/button";
import { ScrollArea } from "@base-ui/react/scroll-area";
import { ArrowUp, Square } from "lucide-react";
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
  WorkbookAgentThreadSnapshot,
  WorkbookAgentTimelineCitation,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
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
import { WorkbookAgentDisclosureRow } from "./workbook-agent-panel-disclosure-row.js";
import {
  agentPanelBodyMutedTextClass,
  agentPanelBodyTextClass,
  agentPanelComposerFrameClass,
  agentPanelComposerSendButtonClass,
  agentPanelComposerTextareaClass,
  agentPanelEyebrowTextClass,
  agentPanelFooterClass,
  agentPanelLabelTextClass,
  agentPanelMetaTextClass,
  agentPanelScrollAreaContentClass,
  agentPanelScrollAreaRootClass,
  agentPanelScrollAreaScrollbarClass,
  agentPanelScrollAreaThumbClass,
  agentPanelScrollAreaViewportClass,
  agentPanelTimelineListClass,
  agentPanelThreadButtonClass,
  agentPanelThreadListClass,
} from "./workbook-agent-panel-primitives.js";
import { WorkbookAgentMarkdown } from "./workbook-agent-markdown.js";
import { formatWorkbookCollaboratorLabel } from "./workbook-presence-model.js";
import {
  AssistantProgressRow,
  PreviewRangeList,
  WorkflowRunRow,
} from "./workbook-agent-panel-history.js";

const toolStatusPillClass = cva(
  "inline-flex h-5 items-center rounded-full border px-2 text-[10px] leading-none font-semibold uppercase tracking-[0.04em]",
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

const agentPanelThemeStyle: CSSProperties & Record<`--${string}`, string> = {
  "--wb-app-bg": "var(--color-mauve-50)",
  "--wb-surface": "white",
  "--wb-surface-subtle": "var(--color-mauve-50)",
  "--wb-surface-muted": "var(--color-mauve-100)",
  "--wb-border": "var(--color-mauve-200)",
  "--wb-border-strong": "var(--color-mauve-300)",
  "--wb-grid-border": "var(--color-mauve-100)",
  "--wb-text": "var(--color-mauve-950)",
  "--wb-text-muted": "var(--color-mauve-700)",
  "--wb-text-subtle": "var(--color-mauve-500)",
  "--wb-accent": "var(--color-mauve-900)",
  "--wb-accent-soft": "var(--color-mauve-100)",
  "--wb-accent-ring": "var(--color-mauve-400)",
  "--wb-hover": "var(--color-mauve-100)",
  "--wb-shadow-sm": "0 1px 2px rgba(15, 23, 42, 0.04)",
};

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
  const visibleThreadSummaries = props.threadSummaries.filter(
    (threadSummary) => threadSummary.threadId !== props.activeThreadId,
  );
  if (visibleThreadSummaries.length === 0) {
    return null;
  }

  return (
    <div className={agentPanelThreadListClass()}>
      {visibleThreadSummaries.map((threadSummary) => {
        const latestActivity = summarizeThreadActivity(threadSummary.latestEntryText);
        return (
          <Button
            key={threadSummary.threadId}
            aria-label={`Open ${threadSummary.scope} thread ${threadSummary.threadId}`}
            aria-pressed={false}
            className={agentPanelThreadButtonClass({ active: false })}
            data-testid={`workbook-agent-thread-${threadSummary.threadId}`}
            type="button"
            onClick={() => {
              props.onSelectThread(threadSummary.threadId);
            }}
          >
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-2">
                <span className={cn(agentPanelLabelTextClass(), "font-semibold")}>
                  {threadSummary.scope === "shared" ? "Shared" : "Private"}
                </span>
                <span className={agentPanelMetaTextClass()}>
                  {threadSummary.scope === "shared"
                    ? formatWorkbookCollaboratorLabel(threadSummary.ownerUserId)
                    : "Just you"}
                </span>
                <span className={agentPanelMetaTextClass()}>
                  {formatThreadEntryCount(threadSummary.entryCount)}
                </span>
              </div>
              {latestActivity ? (
                <div className={cn(agentPanelMetaTextClass(), "mt-0.5 truncate")}>
                  {latestActivity}
                </div>
              ) : null}
            </div>
            {threadSummary.reviewQueueItemCount > 0 ? (
              <span className={workbookPillClass({ tone: "accent", weight: "strong" })}>
                Review
              </span>
            ) : null}
          </Button>
        );
      })}
    </div>
  );
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry["toolStatus"] }) {
  if (props.status === "completed") {
    return null;
  }
  const label = props.status === "failed" ? "Failed" : "Running";
  return (
    <span
      className={toolStatusPillClass({
        status: props.status === "failed" ? "failed" : "running",
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

function renderToolDisplayName(toolName: string | null): string {
  const normalizedToolName = toolName ? normalizeWorkbookAgentToolName(toolName) : null;
  if (!normalizedToolName) {
    return "Tool call";
  }
  return normalizedToolName
    .split("_")
    .map((segment) =>
      segment.length === 0 ? segment : `${segment[0]!.toUpperCase()}${segment.slice(1)}`,
    )
    .join(" ");
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

function renderMarkdownPlainText(markdown: string): string {
  return markdown
    .replaceAll(/```[\s\S]*?```/g, " ")
    .replaceAll(/`([^`]+)`/g, "$1")
    .replaceAll(/\[([^\]]+)\]\(([^)]+)\)/g, "$1")
    .replaceAll(/!\[([^\]]*)\]\(([^)]+)\)/g, "$1")
    .replaceAll(/^#{1,6}\s+/gm, "")
    .replaceAll(/^>\s?/gm, "")
    .replaceAll(/[*_~]+/g, " ")
    .replaceAll(/\s+/g, " ");
}

function summarizeDisclosureText(text: string | null): string | null {
  if (!text) {
    return null;
  }
  const normalized = renderMarkdownPlainText(text).trim().replaceAll(/\s+/g, " ");
  if (normalized.length === 0) {
    return null;
  }
  return normalized.length <= 88 ? normalized : `${normalized.slice(0, 85)}...`;
}

function summarizePlainText(text: string | null, maxLength = 88): string | null {
  if (!text) {
    return null;
  }
  const normalized = text.trim().replaceAll(/\s+/g, " ");
  if (normalized.length === 0) {
    return null;
  }
  return normalized.length <= maxLength ? normalized : `${normalized.slice(0, maxLength - 3)}...`;
}

function summarizeToolEntry(entry: WorkbookAgentTimelineEntry): string | null {
  const parsed = safeParseToolOutput(entry.outputText);
  if (isRecord(parsed)) {
    if (typeof parsed["summary"] === "string") {
      return summarizePlainText(parsed["summary"], 96);
    }
    const workflowRun = isRecord(parsed["workflowRun"]) ? parsed["workflowRun"] : null;
    if (typeof workflowRun?.["summary"] === "string") {
      return summarizePlainText(workflowRun["summary"], 96);
    }
    const selection = isRecord(parsed["selection"]) ? parsed["selection"] : null;
    if (typeof selection?.["sheetName"] === "string" && typeof selection["address"] === "string") {
      const selectionRange = isRecord(selection["range"]) ? selection["range"] : null;
      const startAddress =
        typeof selectionRange?.["startAddress"] === "string"
          ? selectionRange["startAddress"]
          : selection["address"];
      const endAddress =
        typeof selectionRange?.["endAddress"] === "string"
          ? selectionRange["endAddress"]
          : selection["address"];
      return `${selection["sheetName"]}!${startAddress}${startAddress === endAddress ? "" : `:${endAddress}`}`;
    }
    const range = isRecord(parsed["range"]) ? parsed["range"] : null;
    if (
      typeof range?.["sheetName"] === "string" &&
      typeof range["startAddress"] === "string" &&
      typeof range["endAddress"] === "string"
    ) {
      const startAddress = range["startAddress"];
      const endAddress = range["endAddress"];
      return `${range["sheetName"]}!${startAddress}${startAddress === endAddress ? "" : `:${endAddress}`}`;
    }
    if (typeof parsed["sheetCount"] === "number") {
      return `${String(parsed["sheetCount"])} ${parsed["sheetCount"] === 1 ? "sheet" : "sheets"}`;
    }
    if (typeof parsed["changeCount"] === "number") {
      return `${String(parsed["changeCount"])} ${parsed["changeCount"] === 1 ? "change" : "changes"}`;
    }
  }
  const outputText = entry.outputText?.trim() ?? "";
  if (outputText.length > 0 && !outputText.startsWith("{") && !outputText.startsWith("[")) {
    return summarizePlainText(outputText, 96);
  }
  const argumentsText = entry.argumentsText?.trim() ?? "";
  if (
    argumentsText.length > 0 &&
    !argumentsText.startsWith("{") &&
    !argumentsText.startsWith("[")
  ) {
    return summarizePlainText(argumentsText, 96);
  }
  return null;
}

function isAppliedExecutionSystemEntry(entry: WorkbookAgentTimelineEntry): boolean {
  return (
    entry.kind === "system" &&
    (entry.text?.startsWith("Applied workbook change set at revision r") === true ||
      entry.text?.startsWith("Applied automatically workbook change set at revision r") === true ||
      entry.text?.startsWith("Applied automatically selected workbook change set at revision r") ===
        true ||
      entry.text?.startsWith("Applied selected workbook change set at revision r") === true)
  );
}

function TextDisclosureEntryRow(props: {
  readonly entry: WorkbookAgentTimelineEntry;
  readonly label: "Thought" | "Plan";
}) {
  const bodyText =
    props.entry.kind === "reasoning" || props.entry.kind === "plan" ? props.entry.text : null;
  if (!bodyText?.trim().length) {
    return null;
  }
  const summary = summarizeDisclosureText(bodyText);
  const disclosureKey = props.entry.kind;

  return (
    <WorkbookAgentDisclosureRow
      id={props.entry.id}
      label={props.label}
      labelClassName="font-semibold"
      panelTestId={`workbook-agent-${disclosureKey}-panel-${props.entry.id}`}
      summary={summary ?? `No ${props.label.toLowerCase()} details available.`}
      triggerLabel={{
        expanded: `Collapse ${props.label.toLowerCase()}`,
        collapsed: `Expand ${props.label.toLowerCase()}`,
      }}
      triggerTestId={`workbook-agent-${disclosureKey}-toggle-${props.entry.id}`}
    >
      <div className={agentPanelBodyMutedTextClass()}>
        <WorkbookAgentMarkdown markdown={bodyText} />
      </div>
      <TimelineCitationList citations={props.entry.citations} />
    </WorkbookAgentDisclosureRow>
  );
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
        <div className={cn(agentPanelMetaTextClass(), "flex items-start justify-between gap-3")}>
          <div>
            {readNumber(summary?.["issueCount"], issues.length)} issues ·{" "}
            {readNumber(summary?.["scannedFormulaCells"])} formulas
          </div>
          <div className={cn(agentPanelEyebrowTextClass(), "text-right")}>
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
                  <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>
                    {readString(issue["sheetName"])}!{readString(issue["address"])}
                  </div>
                  <div className={cn(agentPanelMetaTextClass(), "mt-1 break-all")}>
                    {readString(issue["formula"])}
                  </div>
                </div>
                <div className={cn(agentPanelEyebrowTextClass(), "text-right")}>
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
        <div className={cn(agentPanelMetaTextClass(), "flex items-start justify-between gap-3")}>
          <div className="truncate">“{readString(parsed["query"])}”</div>
          <div className={agentPanelEyebrowTextClass()}>
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
                  <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>
                    {readString(match["kind"]) === "sheet"
                      ? `Sheet ${readString(match["sheetName"])}`
                      : `${readString(match["sheetName"])}!${readString(match["address"])}`}
                  </div>
                  <div className={cn(agentPanelMetaTextClass(), "mt-1 break-all")}>
                    {readString(match["snippet"])}
                  </div>
                </div>
                <div className={agentPanelEyebrowTextClass()}>
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
        <div className={cn(agentPanelMetaTextClass(), "flex items-start justify-between gap-3")}>
          <div>
            {readString(root?.["sheetName"])}!{readString(root?.["address"])}
          </div>
          <div className={agentPanelEyebrowTextClass()}>
            {readString(parsed["direction"], "both")} · {readNumber(parsed["depth"])} hops
          </div>
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {layers.map((layer) => (
            <div
              key={`trace-layer-${readNumber(layer["depth"])}`}
              className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2"
            >
              <div className={agentPanelEyebrowTextClass()}>Hop {readNumber(layer["depth"])}</div>
              <div className="mt-2 grid gap-2 md:grid-cols-2">
                <div>
                  <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>Precedents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer["precedents"]) && layer["precedents"].length > 0 ? (
                      layer["precedents"]
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`precedent:${readString(node["sheetName"])}:${readString(node["address"])}`}
                            className={agentPanelMetaTextClass()}
                          >
                            {readString(node["sheetName"])}!{readString(node["address"])}{" "}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node["formula"]) || readString(node["valueText"])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className={agentPanelMetaTextClass()}>None</div>
                    )}
                  </div>
                </div>
                <div>
                  <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>Dependents</div>
                  <div className="mt-1 flex flex-col gap-1">
                    {Array.isArray(layer["dependents"]) && layer["dependents"].length > 0 ? (
                      layer["dependents"]
                        .flatMap((node) => (isRecord(node) ? [node] : []))
                        .map((node) => (
                          <div
                            key={`dependent:${readString(node["sheetName"])}:${readString(node["address"])}`}
                            className={agentPanelMetaTextClass()}
                          >
                            {readString(node["sheetName"])}!{readString(node["address"])}{" "}
                            <span className="text-[var(--wb-text-muted)]">
                              {readString(node["formula"]) || readString(node["valueText"])}
                            </span>
                          </div>
                        ))
                    ) : (
                      <div className={agentPanelMetaTextClass()}>None</div>
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
      <div className="flex justify-end px-3 py-3">
        <div
          className={cn(
            agentPanelBodyTextClass(),
            "max-w-[82%] break-words rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2 shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
          )}
        >
          <WorkbookAgentMarkdown markdown={entry.text ?? ""} />
          <TimelineCitationList citations={entry.citations} />
        </div>
      </div>
    );
  }

  if (entry.kind === "assistant") {
    if (entry.phase === "progress") {
      return null;
    }
    if (!entry.text?.trim().length) {
      return null;
    }
    return (
      <div className={cn(agentPanelBodyTextClass(), "px-3 py-3")}>
        <WorkbookAgentMarkdown markdown={entry.text} />
        <TimelineCitationList citations={entry.citations} />
      </div>
    );
  }

  if (entry.kind === "reasoning") {
    return <TextDisclosureEntryRow entry={entry} label="Thought" />;
  }

  if (entry.kind === "plan") {
    return <TextDisclosureEntryRow entry={entry} label="Plan" />;
  }

  if (entry.kind === "tool") {
    const displayName = renderToolDisplayName(entry.toolName);
    const summary = summarizeToolEntry(entry);
    const hasStructuredOutput = supportsStructuredToolOutput(entry.toolName);
    const parsedOutput = safeParseToolOutput(entry.outputText);
    const hasDetails =
      (entry.argumentsText?.trim().length ?? 0) > 0 || (entry.outputText?.trim().length ?? 0) > 0;
    if (!hasDetails) {
      return (
        <div className="px-3 py-2.5">
          <div className="flex items-center justify-between gap-3">
            <div className="min-w-0 flex flex-1 items-center gap-2.5">
              <div className={cn(agentPanelLabelTextClass(), "min-w-0")}>{displayName}</div>
              {summary ? (
                <div
                  className={cn(
                    agentPanelMetaTextClass(),
                    "min-w-0 flex-1 whitespace-normal break-words",
                  )}
                >
                  {summary}
                </div>
              ) : null}
            </div>
            <ToolStatusPill status={entry.toolStatus} />
          </div>
          <TimelineCitationList citations={entry.citations} />
        </div>
      );
    }
    return (
      <WorkbookAgentDisclosureRow
        badge={<ToolStatusPill status={entry.toolStatus} />}
        id={entry.id}
        label={displayName}
        panelTestId={`workbook-agent-tool-panel-${entry.id}`}
        summary={summary}
        triggerLabel={{
          expanded: `Collapse ${displayName}`,
          collapsed: `Expand ${displayName}`,
        }}
        triggerTestId={`workbook-agent-tool-toggle-${entry.id}`}
      >
        {entry.argumentsText?.trim().length ? (
          <div>
            <div className={agentPanelEyebrowTextClass()}>Arguments</div>
            <pre
              className={cn(
                agentPanelMetaTextClass(),
                "mt-1 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2",
              )}
            >
              {entry.argumentsText}
            </pre>
          </div>
        ) : null}
        {entry.outputText?.trim().length ? (
          <div className={entry.argumentsText?.trim().length ? "mt-2" : undefined}>
            <div className={agentPanelEyebrowTextClass()}>Output</div>
            {hasStructuredOutput && parsedOutput !== null ? (
              <StructuredToolOutput toolName={entry.toolName} outputText={entry.outputText} />
            ) : (
              <pre
                className={cn(
                  agentPanelMetaTextClass(),
                  "mt-1 overflow-x-auto rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2",
                )}
              >
                {entry.outputText}
              </pre>
            )}
          </div>
        ) : null}
        <TimelineCitationList citations={entry.citations} />
      </WorkbookAgentDisclosureRow>
    );
  }

  if (isAppliedExecutionSystemEntry(entry)) {
    return null;
  }

  return (
    <div className="px-3 py-2.5">
      <div className={agentPanelMetaTextClass()}>{entry.text}</div>
      <TimelineCitationList citations={entry.citations} />
    </div>
  );
}

function TimelineCitationList(props: {
  readonly citations: readonly WorkbookAgentTimelineCitation[];
}) {
  const segments = summarizeTimelineCitations(props.citations);
  if (segments.length === 0) {
    return null;
  }
  return (
    <div className={cn(agentPanelMetaTextClass(), "mt-1 break-words")}>
      {segments.map((segment, index) => (
        <Fragment key={segment}>
          {index > 0 ? <span aria-hidden="true"> · </span> : null}
          <span>{segment}</span>
        </Fragment>
      ))}
    </div>
  );
}

function summarizeTimelineCitations(
  citations: readonly WorkbookAgentTimelineCitation[],
): readonly string[] {
  const seen = new Set<string>();
  const segments: string[] = [];
  for (const citation of citations) {
    if (citation.kind === "revision") {
      continue;
    }
    const address =
      citation.startAddress === citation.endAddress
        ? `${citation.sheetName}!${citation.startAddress}`
        : `${citation.sheetName}!${citation.startAddress}:${citation.endAddress}`;
    const segment = `${citation.role === "target" ? "Target" : "Source"} ${address}`;
    if (seen.has(segment)) {
      continue;
    }
    seen.add(segment);
    segments.push(segment);
  }
  return segments;
}

function ReviewItemCard(props: {
  readonly reviewBundle: WorkbookAgentCommandBundle;
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
  readonly isApplyingReviewItem: boolean;
  readonly onApply: () => void;
  readonly onDismiss: () => void;
  readonly onReview: (decision: "approved" | "rejected") => void;
  readonly onSelectAll: () => void;
  readonly onToggleCommand: (commandIndex: number) => void;
}) {
  const selectedCount = props.selectedCommandIndexes.length;
  const hasFullSelection = selectedCount === props.reviewBundle.commands.length;
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
    !props.isApplyingReviewItem &&
    selectedCount > 0 &&
    sharedApprovalOwnerLabel === null &&
    (props.sharedReviewStatus === null || props.sharedReviewStatus === "approved");
  const applyLabel =
    selectedCount > 0 && !hasFullSelection
      ? props.isApplyingReviewItem
        ? "Applying…"
        : "Apply"
      : props.sharedReviewStatus === "pending"
        ? "Owner review"
        : props.sharedReviewStatus === "rejected"
          ? "Returned"
          : props.isApplyingReviewItem
            ? "Applying…"
            : "Apply";
  return (
    <div
      className={cn(
        workbookSurfaceClass({ emphasis: "raised" }),
        "border-[var(--wb-border-strong)] px-3 py-3",
      )}
    >
      <div className={cn(agentPanelLabelTextClass(), "font-semibold")}>
        {props.reviewBundle.summary}
      </div>
      <div className={cn(workbookInsetClass(), "mt-3 px-2 py-2")}>
        <div className="flex items-center justify-between gap-3">
          <div className={agentPanelEyebrowTextClass()}>
            {String(selectedCount)}/{String(props.reviewBundle.commands.length)}
          </div>
          {!hasFullSelection ? (
            <Button
              className="text-[12px] leading-5 font-medium text-[var(--wb-accent)] transition-colors hover:brightness-[0.95] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
              type="button"
              onClick={props.onSelectAll}
            >
              All
            </Button>
          ) : null}
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {props.reviewBundle.commands.map((command, index) => {
            const checked = props.selectedCommandIndexes.includes(index);
            const commandLabel = describeWorkbookAgentCommand(command);
            return (
              <div
                key={`${props.reviewBundle.id}:${JSON.stringify(command)}`}
                className={cn(
                  "flex items-start gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2 transition-colors",
                  checked
                    ? "border-[var(--wb-accent-ring)] bg-[var(--wb-surface)]"
                    : "border-[var(--wb-border)] bg-[var(--wb-surface)]",
                )}
              >
                <input
                  aria-label={`Toggle workbook review item change ${String(index + 1)}: ${commandLabel}`}
                  checked={checked}
                  className="mt-0.5 h-4 w-4 rounded border-[var(--wb-border)] text-[var(--wb-accent)] focus:ring-[var(--wb-accent-ring)]"
                  data-testid={`workbook-agent-review-command-toggle-${String(index)}`}
                  type="checkbox"
                  onChange={() => {
                    props.onToggleCommand(index);
                  }}
                />
                <div className="min-w-0">
                  <div className={agentPanelLabelTextClass()}>{commandLabel}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
      <PreviewRangeList ranges={props.preview?.ranges ?? props.reviewBundle.affectedRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div
          className={cn(
            workbookInsetClass(),
            agentPanelMetaTextClass(),
            "mt-2 border-transparent px-2 py-2",
          )}
        >
          {props.preview.structuralChanges.join(" · ")}
        </div>
      ) : null}
      {sharedApprovalOwnerLabel ? (
        <div
          className={cn(
            workbookAlertClass({ tone: "warning" }),
            agentPanelMetaTextClass(),
            "mt-2 border-[var(--wb-border-strong)]",
          )}
        >
          Owner review routes medium/high-risk changes to {sharedApprovalOwnerLabel} on this shared
          thread.
        </div>
      ) : null}
      {recommendationSummary ? (
        <div
          className={cn(
            workbookInsetClass(),
            agentPanelMetaTextClass(),
            "mt-2 border-transparent px-2 py-2",
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
              tone: props.sharedReviewStatus === "rejected" ? "danger" : "warning",
            }),
            agentPanelMetaTextClass(),
            "mt-2 border-[var(--wb-border-strong)]",
          )}
        >
          {props.sharedReviewStatus === "pending"
            ? `Owner review is in progress with ${sharedReviewOwnerLabel}.`
            : props.sharedReviewStatus === "approved"
              ? `Approved by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`
              : `Returned by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`}
        </div>
      ) : null}
      {props.canFinalizeSharedBundle && props.sharedReviewStatus !== null ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <Button
            className={workbookButtonClass({ tone: "neutral" })}
            data-testid="workbook-agent-review-item-reject"
            type="button"
            onClick={() => {
              props.onReview("rejected");
            }}
          >
            Reject
          </Button>
          <Button
            className={workbookButtonClass({ tone: "accent" })}
            data-testid="workbook-agent-review-item-approve"
            type="button"
            onClick={() => {
              props.onReview("approved");
            }}
          >
            Approve
          </Button>
        </div>
      ) : null}
      {props.canRecommendSharedBundle && props.sharedReviewStatus === "pending" ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <Button
            className={workbookButtonClass({ tone: "neutral" })}
            data-testid="workbook-agent-review-item-reject"
            type="button"
            onClick={() => {
              props.onReview("rejected");
            }}
          >
            Recommend reject
          </Button>
          <Button
            className={workbookButtonClass({ tone: "accent" })}
            data-testid="workbook-agent-review-item-approve"
            type="button"
            onClick={() => {
              props.onReview("approved");
            }}
          >
            Recommend approve
          </Button>
        </div>
      ) : null}
      {props.preview?.cellDiffs?.length ? (
        <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
          <div className="max-h-44 overflow-y-auto">
            {props.preview.cellDiffs.map((diff) => (
              <div
                key={`${diff.sheetName}:${diff.address}`}
                className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-x-2 border-t border-[var(--wb-border)] px-2 py-2 first:border-t-0"
              >
                <div className={cn(agentPanelLabelTextClass(), "col-span-2")}>
                  {diff.sheetName}!{diff.address}
                </div>
                <div className="col-span-2 mt-1 flex flex-wrap gap-1">
                  {diff.changeKinds.map((kind) => (
                    <span key={kind} className={workbookPillClass({ tone: "neutral" })}>
                      {renderPreviewChangeKind(kind)}
                    </span>
                  ))}
                </div>
                <div className={agentPanelMetaTextClass()}>
                  {(diff.beforeFormula ?? String(diff.beforeInput ?? "")) || "(empty)"}
                </div>
                <div className={agentPanelLabelTextClass()}>
                  {(diff.afterFormula ?? String(diff.afterInput ?? "")) || "(empty)"}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <Button
          className={workbookButtonClass({ tone: "neutral" })}
          type="button"
          onClick={props.onDismiss}
        >
          Clear
        </Button>
        <Button
          className={workbookButtonClass({ tone: "accent", weight: "strong" })}
          data-testid="workbook-agent-apply-review-item"
          disabled={!canApply}
          type="button"
          onClick={props.onApply}
        >
          {applyLabel}
        </Button>
      </div>
    </div>
  );
}

function ExecutionRecordRow(props: {
  readonly record: WorkbookAgentExecutionRecord;
  readonly onReplay: () => void;
}) {
  return (
    <div
      className={cn(
        workbookSurfaceClass({ emphasis: "raised" }),
        "border-[var(--wb-border)] px-3 py-2.5",
      )}
    >
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <span className={workbookPillClass({ tone: "accent", weight: "strong" })}>
              {props.record.appliedBy === "auto" ? "Applied automatically" : "Applied"}
            </span>
            <span className={agentPanelMetaTextClass()}>
              Revision r{String(props.record.appliedRevision)}
            </span>
          </div>
          <div className={cn(agentPanelLabelTextClass(), "mt-1.5")}>{props.record.summary}</div>
        </div>
        <Button
          className={workbookButtonClass({ tone: "neutral" })}
          type="button"
          onClick={props.onReplay}
        >
          Run again
        </Button>
      </div>
    </div>
  );
}

export function WorkbookAgentPanel(props: {
  readonly activeThreadId: string | null;
  readonly optimisticEntries?: readonly WorkbookAgentTimelineEntry[];
  readonly snapshot: WorkbookAgentThreadSnapshot | null;
  readonly activeResponseTurnId: string | null;
  readonly showAssistantProgress: boolean;
  readonly reviewBundle: WorkbookAgentCommandBundle | null;
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
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[];
  readonly draft: string;
  readonly isLoading: boolean;
  readonly isApplyingReviewItem: boolean;
  readonly onApplyReviewItem: () => void;
  readonly onDraftChange: (value: string) => void;
  readonly onDismissReviewItem: () => void;
  readonly onReviewReviewItem: (decision: "approved" | "rejected") => void;
  readonly onInterrupt: () => void;
  readonly onSelectAllReviewCommands: () => void;
  readonly onSelectThread: (threadId: string) => void;
  readonly onToggleReviewCommand: (commandIndex: number) => void;
  readonly onCancelWorkflowRun: (runId: string) => void;
  readonly onReplayExecutionRecord: (recordId: string) => void;
  readonly onSubmit: () => void;
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null);

  const optimisticEntries = props.optimisticEntries ?? [];

  useEffect(() => {
    const node = scrollRef.current;
    if (!node) {
      return;
    }
    node.scrollTop = node.scrollHeight;
  }, [optimisticEntries.length, props.snapshot?.entries.length, props.snapshot?.status]);

  const isRunning = props.snapshot?.status === "inProgress";
  const visibleEntries = [...optimisticEntries, ...(props.snapshot?.entries ?? [])];
  const progressAnchorIndex =
    props.showAssistantProgress && props.activeResponseTurnId
      ? visibleEntries.findLastIndex(
          (entry) =>
            entry.turnId === props.activeResponseTurnId && !isAppliedExecutionSystemEntry(entry),
        )
      : -1;

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col bg-[var(--wb-app-bg)]"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
      style={agentPanelThemeStyle}
    >
      <ScrollArea.Root className={agentPanelScrollAreaRootClass()}>
        <ScrollArea.Viewport
          ref={scrollRef}
          className={agentPanelScrollAreaViewportClass()}
          data-testid="workbook-agent-panel-scroll-viewport"
        >
          <ScrollArea.Content className={agentPanelScrollAreaContentClass()}>
            <div className="bg-[var(--wb-app-bg)] px-2.5 py-2.5">
              <ThreadSummaryStrip
                activeThreadId={props.activeThreadId}
                threadSummaries={props.threadSummaries}
                onSelectThread={props.onSelectThread}
              />
              {props.isLoading ? null : visibleEntries.length > 0 ? (
                <div className="flex flex-col gap-2">
                  <div className={agentPanelTimelineListClass()}>
                    {visibleEntries.map((entry, index) => (
                      <Fragment key={entry.id}>
                        <WorkbookAgentEntryRow entry={entry} />
                        {props.showAssistantProgress && progressAnchorIndex === index ? (
                          <AssistantProgressRow />
                        ) : null}
                      </Fragment>
                    ))}
                    {props.showAssistantProgress && progressAnchorIndex < 0 ? (
                      <AssistantProgressRow />
                    ) : null}
                  </div>
                  {props.executionRecords.length > 0 ? (
                    <div className="pt-1">
                      <div className={cn(agentPanelEyebrowTextClass(), "mb-2")}>Recent changes</div>
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
                  {props.workflowRuns.length > 0 ? (
                    <div className="pt-1">
                      <div className={cn(agentPanelEyebrowTextClass(), "mb-2")}>Workflows</div>
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
                </div>
              ) : null}
            </div>
          </ScrollArea.Content>
        </ScrollArea.Viewport>
        <ScrollArea.Scrollbar
          className={agentPanelScrollAreaScrollbarClass()}
          keepMounted
          orientation="vertical"
        >
          <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
        </ScrollArea.Scrollbar>
      </ScrollArea.Root>
      <div className={agentPanelFooterClass()}>
        {props.reviewBundle ? (
          <div className="mb-3">
            <ReviewItemCard
              reviewBundle={props.reviewBundle}
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
              isApplyingReviewItem={props.isApplyingReviewItem}
              onApply={props.onApplyReviewItem}
              onDismiss={props.onDismissReviewItem}
              onReview={props.onReviewReviewItem}
              onSelectAll={props.onSelectAllReviewCommands}
              onToggleCommand={props.onToggleReviewCommand}
            />
          </div>
        ) : null}
        <form
          onSubmit={(event) => {
            event.preventDefault();
            props.onSubmit();
          }}
        >
          <label className="sr-only" htmlFor="workbook-agent-input">
            Ask the workbook assistant
          </label>
          <div className={agentPanelComposerFrameClass()}>
            <textarea
              id="workbook-agent-input"
              className={agentPanelComposerTextareaClass()}
              data-testid="workbook-agent-input"
              placeholder="Ask the workbook assistant"
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
            <Button
              aria-label={isRunning ? "Stop" : "Send message"}
              className={agentPanelComposerSendButtonClass()}
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
            </Button>
          </div>
        </form>
      </div>
    </div>
  );
}

function SendArrowIcon() {
  return <ArrowUp aria-hidden="true" className="size-5" strokeWidth={1.9} />;
}

function StopIcon() {
  return <Square aria-hidden="true" className="size-4 fill-current" strokeWidth={0} />;
}
