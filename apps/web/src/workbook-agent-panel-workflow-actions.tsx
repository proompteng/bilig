import { useCallback, useState } from "react";
import { Button } from "@base-ui/react/button";
import { ChevronDown } from "lucide-react";
import { Popover } from "@base-ui/react/popover";
import { cn } from "./cn.js";
import { workbookButtonClass } from "./workbook-shell-chrome.js";
import {
  agentPanelActionButtonClass,
  agentPanelActionGridClass,
  agentPanelFieldClass,
  agentPanelSectionClass,
  agentPanelSectionHeaderClass,
  agentPanelSectionHintClass,
  agentPanelSectionTitleClass,
  agentPanelToolsPanelClass,
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
  const [isOpen, setIsOpen] = useState(false);
  const compactButtonClass = cn(
    workbookButtonClass({ size: "sm", tone: "neutral", weight: "regular" }),
    agentPanelActionButtonClass({ emphasis: "subtle" }),
  );
  const primaryCompactButtonClass = cn(
    workbookButtonClass({ size: "sm", tone: "accent", weight: "strong" }),
    agentPanelActionButtonClass({ emphasis: "strong" }),
  );
  const closeTools = useCallback(() => {
    setIsOpen(false);
  }, []);

  const runWorkflow = useCallback(
    (template: WorkflowActionTemplate) => {
      props.onStartWorkflow(template);
      closeTools();
    },
    [closeTools, props],
  );

  const runNamedWorkflow = useCallback(
    (template: "createSheet" | "renameCurrentSheet") => {
      const trimmedName = sheetName.trim();
      if (trimmedName.length === 0) {
        return;
      }
      props.onStartNamedWorkflow(template, trimmedName);
      setSheetName("");
      closeTools();
    },
    [closeTools, props, sheetName],
  );

  const runStructuralWorkflow = useCallback(
    (
      template: "hideCurrentRow" | "hideCurrentColumn" | "unhideCurrentRow" | "unhideCurrentColumn",
    ) => {
      props.onStartStructuralWorkflow(template);
      closeTools();
    },
    [closeTools, props],
  );

  const runSearchWorkflow = useCallback(() => {
    const trimmedQuery = searchQuery.trim();
    if (trimmedQuery.length === 0) {
      return;
    }
    props.onStartSearchWorkflow(trimmedQuery);
    setSearchQuery("");
    closeTools();
  }, [closeTools, props, searchQuery]);

  return (
    <Popover.Root
      modal={false}
      open={isOpen}
      onOpenChange={(nextOpen) => {
        setIsOpen(nextOpen);
      }}
    >
      <div className="flex items-center justify-end">
        <Popover.Trigger
          aria-expanded={isOpen}
          className={cn(
            workbookButtonClass({ size: "sm", tone: "neutral", weight: "strong" }),
            "min-w-[5rem] gap-1.5",
          )}
          data-testid="workbook-agent-workflow-toggle"
          disabled={props.disabled || props.isStartingWorkflow}
          type="button"
        >
          <span>Tools</span>
          <ChevronDown
            className={cn(
              "h-3.5 w-3.5 shrink-0 stroke-[1.75] transition-transform",
              isOpen && "rotate-180",
            )}
          />
        </Popover.Trigger>
      </div>
      <Popover.Portal keepMounted>
        <Popover.Positioner align="end" className="z-[1000]" side="top" sideOffset={8}>
          <Popover.Popup
            aria-label="Assistant tools"
            className={cn(agentPanelToolsPanelClass(), "mt-0 w-[22rem] max-w-[calc(100vw-2rem)]")}
          >
            <div className={agentPanelSectionHeaderClass()}>
              <div className="min-w-0">
                <div className={agentPanelSectionTitleClass()}>Assistant tools</div>
                <div className={agentPanelSectionHintClass()}>Run a structured helper</div>
              </div>
            </div>
            <div className="mt-2 grid gap-2">
              <div className={agentPanelActionGridClass()}>
                {WORKFLOW_ACTIONS.map((action) => (
                  <Button
                    key={action.template}
                    className={compactButtonClass}
                    data-testid={`workbook-agent-workflow-start-${action.template}`}
                    disabled={props.disabled || props.isStartingWorkflow}
                    type="button"
                    onClick={() => {
                      runWorkflow(action.template);
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : action.label}
                  </Button>
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
                      runNamedWorkflow("createSheet");
                    }
                  }}
                />
                <div className={agentPanelActionGridClass()}>
                  <Button
                    className={primaryCompactButtonClass}
                    data-testid="workbook-agent-workflow-start-createSheet"
                    disabled={
                      props.disabled || props.isStartingWorkflow || sheetName.trim().length === 0
                    }
                    type="button"
                    onClick={() => {
                      runNamedWorkflow("createSheet");
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Create sheet"}
                  </Button>
                  <Button
                    className={compactButtonClass}
                    data-testid="workbook-agent-workflow-start-renameCurrentSheet"
                    disabled={
                      props.disabled || props.isStartingWorkflow || sheetName.trim().length === 0
                    }
                    type="button"
                    onClick={() => {
                      runNamedWorkflow("renameCurrentSheet");
                    }}
                  >
                    Rename current
                  </Button>
                </div>
                <div className={agentPanelActionGridClass()}>
                  <Button
                    className={compactButtonClass}
                    data-testid="workbook-agent-workflow-start-hideCurrentRow"
                    disabled={props.disabled || props.isStartingWorkflow}
                    type="button"
                    onClick={() => {
                      runStructuralWorkflow("hideCurrentRow");
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Hide current row"}
                  </Button>
                  <Button
                    className={compactButtonClass}
                    data-testid="workbook-agent-workflow-start-hideCurrentColumn"
                    disabled={props.disabled || props.isStartingWorkflow}
                    type="button"
                    onClick={() => {
                      runStructuralWorkflow("hideCurrentColumn");
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Hide current column"}
                  </Button>
                </div>
                <div className={agentPanelActionGridClass()}>
                  <Button
                    className={compactButtonClass}
                    data-testid="workbook-agent-workflow-start-unhideCurrentRow"
                    disabled={props.disabled || props.isStartingWorkflow}
                    type="button"
                    onClick={() => {
                      runStructuralWorkflow("unhideCurrentRow");
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Unhide current row"}
                  </Button>
                  <Button
                    className={compactButtonClass}
                    data-testid="workbook-agent-workflow-start-unhideCurrentColumn"
                    disabled={props.disabled || props.isStartingWorkflow}
                    type="button"
                    onClick={() => {
                      runStructuralWorkflow("unhideCurrentColumn");
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Unhide current column"}
                  </Button>
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
                        runSearchWorkflow();
                      }
                    }}
                  />
                  <Button
                    className={cn(primaryCompactButtonClass, "min-w-[5rem]")}
                    data-testid="workbook-agent-workflow-start-searchWorkbookQuery"
                    disabled={
                      props.disabled || props.isStartingWorkflow || searchQuery.trim().length === 0
                    }
                    type="button"
                    onClick={() => {
                      runSearchWorkflow();
                    }}
                  >
                    {props.isStartingWorkflow ? "Starting…" : "Search"}
                  </Button>
                </div>
              </div>
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  );
}
