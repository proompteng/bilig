import { useState } from "react";
import type { WorkbookAgentWorkflowTemplate } from "@bilig/contracts";
import { cn } from "./cn.js";
import {
  workbookButtonClass,
  workbookInsetClass,
  workbookPillClass,
} from "./workbook-shell-chrome.js";

interface WorkflowActionDefinition {
  readonly template: Exclude<
    WorkbookAgentWorkflowTemplate,
    "searchWorkbookQuery" | "createSheet" | "renameCurrentSheet"
  >;
  readonly label: string;
  readonly summary: string;
}

const WORKFLOW_ACTIONS: readonly WorkflowActionDefinition[] = [
  {
    template: "summarizeWorkbook",
    label: "Summarize workbook",
    summary: "Read workbook structure, sheets, and layout metadata.",
  },
  {
    template: "summarizeCurrentSheet",
    label: "Summarize current sheet",
    summary: "Read the active sheet layout, tables, filters, freeze panes, and hidden metadata.",
  },
  {
    template: "describeRecentChanges",
    label: "Recent changes",
    summary: "Read the latest durable workbook revisions and summarize them.",
  },
  {
    template: "findFormulaIssues",
    label: "Sheet formula issues",
    summary: "Scan formulas on the active sheet for errors, cycles, and JS-only fallback formulas.",
  },
  {
    template: "traceSelectionDependencies",
    label: "Trace selection links",
    summary: "Trace precedents and dependents from the current selection context.",
  },
  {
    template: "explainSelectionCell",
    label: "Explain current cell",
    summary: "Inspect the selected cell, formula state, and direct workbook links.",
  },
];

export function WorkflowActionStrip(props: {
  readonly disabled: boolean;
  readonly isStartingWorkflow: boolean;
  readonly onStartWorkflow: (
    template: Exclude<
      WorkbookAgentWorkflowTemplate,
      "searchWorkbookQuery" | "createSheet" | "renameCurrentSheet"
    >,
  ) => void;
  readonly onStartNamedWorkflow: (
    template: Extract<WorkbookAgentWorkflowTemplate, "createSheet" | "renameCurrentSheet">,
    name: string,
  ) => void;
  readonly onStartSearchWorkflow: (query: string) => void;
  readonly onStartStructuralWorkflow: (
    template: Extract<WorkbookAgentWorkflowTemplate, "hideCurrentRow" | "hideCurrentColumn">,
  ) => void;
}) {
  const [searchQuery, setSearchQuery] = useState("");
  const [sheetName, setSheetName] = useState("");

  return (
    <div className={cn(workbookInsetClass(), "mt-2 px-2 py-2")}>
      <div className="flex items-center justify-between gap-2">
        <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          Quick workflows
        </div>
        <span className={workbookPillClass({ tone: "neutral" })}>Saved reports + previews</span>
      </div>
      <div className="mt-2 grid gap-2">
        {WORKFLOW_ACTIONS.map((action) => (
          <button
            key={action.template}
            className={cn(
              workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" }),
              "h-auto items-start justify-start px-3 py-2 text-left",
            )}
            data-testid={`workbook-agent-workflow-start-${action.template}`}
            disabled={props.disabled || props.isStartingWorkflow}
            type="button"
            onClick={() => {
              props.onStartWorkflow(action.template);
            }}
          >
            <span className="flex flex-col gap-0.5">
              <span className="text-[11px] font-semibold text-[var(--wb-text)]">
                {props.isStartingWorkflow ? "Starting…" : action.label}
              </span>
              <span className="text-[11px] leading-4 text-[var(--wb-text-subtle)]">
                {action.summary}
              </span>
            </span>
          </button>
        ))}
        <div className="grid gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
          <div className="text-[11px] font-semibold text-[var(--wb-text)]">
            Structural previews
          </div>
          <div className="text-[11px] leading-4 text-[var(--wb-text-subtle)]">
            Stage durable sheet and axis changes through the same preview/apply path as direct
            edits.
          </div>
          <div className="flex items-center gap-2">
            <input
              className="min-w-0 flex-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-2.5 py-1.5 text-[11px] text-[var(--wb-text)] outline-none transition-colors placeholder:text-[var(--wb-text-muted)] focus:border-[var(--wb-accent)] focus:ring-2 focus:ring-[var(--wb-accent-ring)]"
              data-testid="workbook-agent-structural-sheet-name-input"
              disabled={props.disabled || props.isStartingWorkflow}
              placeholder="Sheet name"
              type="text"
              value={sheetName}
              onChange={(event) => {
                setSheetName(event.target.value);
              }}
              onKeyDown={(event) => {
                if (
                  event.key === "Enter" &&
                  sheetName.trim().length > 0 &&
                  !(props.disabled || props.isStartingWorkflow)
                ) {
                  event.preventDefault();
                  props.onStartNamedWorkflow("createSheet", sheetName.trim());
                  setSheetName("");
                }
              }}
            />
            <button
              className={workbookButtonClass({ size: "sm", tone: "accent", weight: "strong" })}
              data-testid="workbook-agent-workflow-start-createSheet"
              disabled={props.disabled || props.isStartingWorkflow || sheetName.trim().length === 0}
              type="button"
              onClick={() => {
                props.onStartNamedWorkflow("createSheet", sheetName.trim());
                setSheetName("");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Create sheet"}
            </button>
            <button
              className={workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" })}
              data-testid="workbook-agent-workflow-start-renameCurrentSheet"
              disabled={props.disabled || props.isStartingWorkflow || sheetName.trim().length === 0}
              type="button"
              onClick={() => {
                props.onStartNamedWorkflow("renameCurrentSheet", sheetName.trim());
                setSheetName("");
              }}
            >
              Rename current
            </button>
          </div>
          <div className="flex items-center gap-2">
            <button
              className={workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" })}
              data-testid="workbook-agent-workflow-start-hideCurrentRow"
              disabled={props.disabled || props.isStartingWorkflow}
              type="button"
              onClick={() => {
                props.onStartStructuralWorkflow("hideCurrentRow");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Hide current row"}
            </button>
            <button
              className={workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" })}
              data-testid="workbook-agent-workflow-start-hideCurrentColumn"
              disabled={props.disabled || props.isStartingWorkflow}
              type="button"
              onClick={() => {
                props.onStartStructuralWorkflow("hideCurrentColumn");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Hide current column"}
            </button>
          </div>
        </div>
        <div className="grid gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
          <div className="text-[11px] font-semibold text-[var(--wb-text)]">Search workbook</div>
          <div className="text-[11px] leading-4 text-[var(--wb-text-subtle)]">
            Run a durable workbook search from the rail and keep the report in the thread.
          </div>
          <div className="flex items-center gap-2">
            <input
              className="min-w-0 flex-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-app-bg)] px-2.5 py-1.5 text-[11px] text-[var(--wb-text)] outline-none transition-colors placeholder:text-[var(--wb-text-muted)] focus:border-[var(--wb-accent)] focus:ring-2 focus:ring-[var(--wb-accent-ring)]"
              data-testid="workbook-agent-workflow-search-input"
              disabled={props.disabled || props.isStartingWorkflow}
              placeholder="Search for a concept, value, or formula"
              type="text"
              value={searchQuery}
              onChange={(event) => {
                setSearchQuery(event.target.value);
              }}
              onKeyDown={(event) => {
                if (
                  event.key === "Enter" &&
                  searchQuery.trim().length > 0 &&
                  !(props.disabled || props.isStartingWorkflow)
                ) {
                  event.preventDefault();
                  props.onStartSearchWorkflow(searchQuery.trim());
                  setSearchQuery("");
                }
              }}
            />
            <button
              className={workbookButtonClass({ size: "sm", tone: "accent", weight: "strong" })}
              data-testid="workbook-agent-workflow-start-searchWorkbookQuery"
              disabled={
                props.disabled || props.isStartingWorkflow || searchQuery.trim().length === 0
              }
              type="button"
              onClick={() => {
                props.onStartSearchWorkflow(searchQuery.trim());
                setSearchQuery("");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Search"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
