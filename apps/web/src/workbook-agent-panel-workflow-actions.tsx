import type { WorkbookAgentWorkflowTemplate } from "@bilig/contracts";
import { cn } from "./cn.js";
import {
  workbookButtonClass,
  workbookInsetClass,
  workbookPillClass,
} from "./workbook-shell-chrome.js";

interface WorkflowActionDefinition {
  readonly template: WorkbookAgentWorkflowTemplate;
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
    template: "describeRecentChanges",
    label: "Recent changes",
    summary: "Read the latest durable workbook revisions and summarize them.",
  },
  {
    template: "findFormulaIssues",
    label: "Formula issues",
    summary: "Scan formulas for errors, cycles, and JS-only fallback formulas.",
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
  readonly onStartWorkflow: (template: WorkbookAgentWorkflowTemplate) => void;
}) {
  return (
    <div className={cn(workbookInsetClass(), "mt-2 px-2 py-2")}>
      <div className="flex items-center justify-between gap-2">
        <div className="text-[10px] font-semibold uppercase tracking-[0.04em] text-[var(--wb-text-subtle)]">
          Quick workflows
        </div>
        <span className={workbookPillClass({ tone: "neutral" })}>Read-only</span>
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
      </div>
    </div>
  );
}
