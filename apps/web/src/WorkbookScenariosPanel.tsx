import {
  formatWorkbookScenarioTimestamp,
  type WorkbookScenarioEntry,
} from "./workbook-scenarios-model.js";

function WorkbookScenarioRow(props: {
  readonly scenario: WorkbookScenarioEntry;
  readonly isDeleting: boolean;
  readonly onOpen: (scenario: WorkbookScenarioEntry) => void;
  readonly onDelete: (scenario: WorkbookScenarioEntry) => void;
}) {
  const { scenario } = props;
  return (
    <div
      className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3 shadow-[var(--wb-shadow-sm)]"
      data-testid="workbook-scenario-row"
    >
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            <span className="inline-flex rounded-full bg-[var(--wb-accent-soft)] px-2 py-0.5 text-[10px] font-semibold uppercase tracking-[0.08em] text-[var(--wb-accent)]">
              Scratchpad
            </span>
            <span className="truncate text-[13px] font-semibold text-[var(--wb-text)]">
              {scenario.name}
            </span>
          </div>
          <div className="mt-1 text-[11px] text-[var(--wb-text-subtle)]">
            Branched from r{scenario.baseRevision}
          </div>
        </div>
      </div>
      <div className="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
        {scenario.targetLabel ? (
          <span>{scenario.targetLabel}</span>
        ) : (
          <span>Workbook-wide branch</span>
        )}
        <span>•</span>
        <span>{formatWorkbookScenarioTimestamp(scenario.updatedAt)}</span>
      </div>
      <div className="mt-3 flex items-center gap-2">
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-accent-ring)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={() => {
            props.onOpen(scenario);
          }}
        >
          Open
        </button>
        <button
          className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[#f0c2c2] bg-[#fff7f7] px-3 text-[12px] font-medium text-[#991b1b] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[#e58e8e] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#f1b5b5] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
          disabled={!scenario.canDelete || props.isDeleting}
          type="button"
          onClick={() => {
            props.onDelete(scenario);
          }}
        >
          {props.isDeleting ? "Deleting..." : "Delete"}
        </button>
      </div>
    </div>
  );
}

export function WorkbookScenariosPanel(props: {
  readonly isOpen: boolean;
  readonly scenarios: readonly WorkbookScenarioEntry[];
  readonly draftName: string;
  readonly isCreating: boolean;
  readonly deletingDocumentIds: readonly string[];
  readonly onDraftNameChange: (value: string) => void;
  readonly onCreate: () => void;
  readonly onClose: () => void;
  readonly onOpen: (scenario: WorkbookScenarioEntry) => void;
  readonly onDelete: (scenario: WorkbookScenarioEntry) => void;
}) {
  if (!props.isOpen) {
    return null;
  }

  return (
    <aside
      aria-label="Workbook scratchpads"
      className="absolute inset-y-3 right-[66rem] z-20 flex w-[20rem] flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[0_20px_48px_rgba(15,23,42,0.16)]"
      data-testid="workbook-scenarios-panel"
      id="workbook-scenarios-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Scratchpads</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            Branch the authoritative workbook into private what-if documents for heavy analysis.
          </p>
        </div>
        <button
          aria-label="Close workbook scratchpads"
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          ×
        </button>
      </div>
      <div className="border-b border-[var(--wb-border)] px-4 py-3">
        <label className="flex flex-col gap-2">
          <span className="text-[11px] font-medium text-[var(--wb-text-subtle)]">Name</span>
          <input
            className="h-9 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[13px] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] outline-none transition-colors focus:border-[var(--wb-accent-ring)] focus:ring-2 focus:ring-[var(--wb-accent-ring)] focus:ring-offset-1"
            data-testid="workbook-scenario-name"
            placeholder="Pricing downside"
            type="text"
            value={props.draftName}
            onChange={(event) => {
              props.onDraftNameChange(event.currentTarget.value);
            }}
          />
        </label>
        <div className="mt-3 flex items-center gap-2">
          <button
            className="ml-auto inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 text-[12px] font-semibold text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)] transition-colors hover:bg-[var(--wb-accent-soft)]/80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
            data-testid="workbook-scenario-create"
            disabled={props.isCreating}
            type="button"
            onClick={props.onCreate}
          >
            {props.isCreating ? "Creating..." : "Create branch"}
          </button>
        </div>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.scenarios.length === 0 ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            Create a scratchpad when you want to explore risky edits without touching the main
            workbook.
          </div>
        ) : (
          <div className="flex flex-col gap-2">
            {props.scenarios.map((scenario) => (
              <WorkbookScenarioRow
                key={scenario.documentId}
                scenario={scenario}
                isDeleting={props.deletingDocumentIds.includes(scenario.documentId)}
                onDelete={props.onDelete}
                onOpen={props.onOpen}
              />
            ))}
          </div>
        )}
      </div>
    </aside>
  );
}
