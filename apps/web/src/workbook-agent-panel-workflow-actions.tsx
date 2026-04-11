import { useState } from "react";
import { cn } from "./cn.js";
import { workbookButtonClass, workbookInsetClass } from "./workbook-shell-chrome.js";
import {
  agentPanelActionButtonClass,
  agentPanelActionGridClass,
  agentPanelFieldClass,
  agentPanelSectionClass,
  agentPanelSectionHeaderClass,
  agentPanelSectionHintClass,
  agentPanelSectionTitleClass,
  agentPanelToggleButtonClass,
} from "./workbook-agent-panel-primitives.js";

type WorkflowActionTemplate =
  | "summarizeWorkbook"
  | "summarizeCurrentSheet"
  | "describeRecentChanges"
  | "findFormulaIssues"
  | "highlightFormulaIssues"
  | "repairFormulaIssues"
  | "highlightCurrentSheetOutliers"
  | "styleCurrentSheetHeaders"
  | "normalizeCurrentSheetHeaders"
  | "normalizeCurrentSheetNumberFormats"
  | "normalizeCurrentSheetWhitespace"
  | "fillCurrentSheetFormulasDown"
  | "traceSelectionDependencies"
  | "explainSelectionCell"
  | "createCurrentSheetRollup"
  | "createCurrentSheetReviewTab";

interface WorkflowActionDefinition {
  readonly template: WorkflowActionTemplate;
  readonly label: string;
}

const WORKFLOW_ACTIONS: readonly WorkflowActionDefinition[] = [
  {
    template: "summarizeWorkbook",
    label: "Workbook summary",
  },
  {
    template: "summarizeCurrentSheet",
    label: "Sheet summary",
  },
  {
    template: "describeRecentChanges",
    label: "Recent changes",
  },
  {
    template: "findFormulaIssues",
    label: "Scan formulas",
  },
  {
    template: "highlightFormulaIssues",
    label: "Highlight formulas",
  },
  {
    template: "repairFormulaIssues",
    label: "Repair formulas",
  },
  {
    template: "highlightCurrentSheetOutliers",
    label: "Highlight outliers",
  },
  {
    template: "styleCurrentSheetHeaders",
    label: "Style headers",
  },
  {
    template: "normalizeCurrentSheetHeaders",
    label: "Normalize headers",
  },
  {
    template: "normalizeCurrentSheetNumberFormats",
    label: "Normalize formats",
  },
  {
    template: "normalizeCurrentSheetWhitespace",
    label: "Normalize text",
  },
  {
    template: "fillCurrentSheetFormulasDown",
    label: "Fill formulas",
  },
  {
    template: "createCurrentSheetRollup",
    label: "Rollup sheet",
  },
  {
    template: "createCurrentSheetReviewTab",
    label: "Create review tab",
  },
  {
    template: "traceSelectionDependencies",
    label: "Trace links",
  },
  {
    template: "explainSelectionCell",
    label: "Explain cell",
  },
];

export function WorkflowActionStrip(props: {
  readonly disabled: boolean;
  readonly isStartingWorkflow: boolean;
  readonly onStartWorkflow: (template: WorkflowActionTemplate) => void;
  readonly onStartNamedWorkflow: (
    template: "createSheet" | "renameCurrentSheet",
    name: string,
  ) => void;
  readonly onStartSearchWorkflow: (query: string) => void;
  readonly onStartStructuralWorkflow: (
    template: "hideCurrentRow" | "hideCurrentColumn" | "unhideCurrentRow" | "unhideCurrentColumn",
  ) => void;
}) {
  const [searchQuery, setSearchQuery] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [isExpanded, setIsExpanded] = useState(false);
  const compactButtonClass = cn(
    workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" }),
    agentPanelActionButtonClass({ emphasis: "subtle" }),
  );
  const primaryCompactButtonClass = cn(
    workbookButtonClass({ size: "sm", tone: "accent", weight: "strong" }),
    agentPanelActionButtonClass({ emphasis: "strong" }),
  );

  return (
    <div className={cn(workbookInsetClass(), "mt-2 px-2 py-2")}>
      <div className={agentPanelSectionHeaderClass()}>
        <div className="min-w-0">
          <div className={agentPanelSectionTitleClass()}>Workflows</div>
          <div className={agentPanelSectionHintClass()}>Optional shortcuts</div>
        </div>
        <button
          aria-expanded={isExpanded}
          className={cn(
            workbookButtonClass({ size: "sm", tone: "neutral", weight: "strong" }),
            agentPanelToggleButtonClass(),
          )}
          data-testid="workbook-agent-workflow-toggle"
          disabled={props.disabled || props.isStartingWorkflow}
          type="button"
          onClick={() => {
            setIsExpanded((value) => !value);
          }}
        >
          {isExpanded ? "Hide" : "Show"}
        </button>
      </div>
      <div className={cn("mt-2 grid gap-2", !isExpanded && "hidden")} hidden={!isExpanded}>
        <div className={agentPanelActionGridClass()}>
          {WORKFLOW_ACTIONS.map((action) => (
            <button
              key={action.template}
              className={compactButtonClass}
              data-testid={`workbook-agent-workflow-start-${action.template}`}
              disabled={props.disabled || props.isStartingWorkflow}
              type="button"
              onClick={() => {
                props.onStartWorkflow(action.template);
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : action.label}
            </button>
          ))}
        </div>
        <div className={cn(agentPanelSectionClass(), "grid gap-2")}>
          <div className={agentPanelSectionTitleClass()}>Structure</div>
          <input
            className={agentPanelFieldClass()}
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
          <div className={agentPanelActionGridClass()}>
            <button
              className={primaryCompactButtonClass}
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
              className={compactButtonClass}
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
          <div className={agentPanelActionGridClass()}>
            <button
              className={compactButtonClass}
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
              className={compactButtonClass}
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
          <div className={agentPanelActionGridClass()}>
            <button
              className={compactButtonClass}
              data-testid="workbook-agent-workflow-start-unhideCurrentRow"
              disabled={props.disabled || props.isStartingWorkflow}
              type="button"
              onClick={() => {
                props.onStartStructuralWorkflow("unhideCurrentRow");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Unhide current row"}
            </button>
            <button
              className={compactButtonClass}
              data-testid="workbook-agent-workflow-start-unhideCurrentColumn"
              disabled={props.disabled || props.isStartingWorkflow}
              type="button"
              onClick={() => {
                props.onStartStructuralWorkflow("unhideCurrentColumn");
              }}
            >
              {props.isStartingWorkflow ? "Starting…" : "Unhide current column"}
            </button>
          </div>
        </div>
        <div className={cn(agentPanelSectionClass(), "grid gap-2")}>
          <div className={agentPanelSectionTitleClass()}>Search workbook</div>
          <div className="grid grid-cols-[minmax(0,1fr)_auto] items-center gap-2">
            <input
              className={agentPanelFieldClass()}
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
              className={cn(primaryCompactButtonClass, "min-w-[5rem]")}
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
