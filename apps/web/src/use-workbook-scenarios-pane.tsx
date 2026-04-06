import { useCallback, useMemo, useState } from "react";
import type { Viewport } from "@bilig/protocol";
import { WorkbookScenariosPanel } from "./WorkbookScenariosPanel.js";
import { navigateToWorkbook as defaultNavigateToWorkbook } from "./workbook-navigation.js";
import {
  createWorkbookScenarioRequest,
  deleteWorkbookScenarioRequest,
} from "./workbook-scenarios-client.js";
import {
  useWorkbookScenarioContext,
  useWorkbookScenarios,
  type ZeroWorkbookScenarioQuerySource,
} from "./use-workbook-scenarios.js";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import type { WorkbookScenarioEntry } from "./workbook-scenarios-model.js";

function defaultWorkbookScenarioName(selection: WorkerRuntimeSelection): string {
  return `${selection.sheetName} ${selection.address} scratchpad`;
}

export interface ZeroWorkbookScenarioSource extends ZeroWorkbookScenarioQuerySource {}

export function useWorkbookScenariosPane(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly selection: WorkerRuntimeSelection;
  readonly zero: ZeroWorkbookScenarioSource;
  readonly enabled: boolean;
  readonly getCurrentViewport: () => Viewport;
  readonly createScenario?: (input: {
    documentId: string;
    name: string;
    selection: WorkerRuntimeSelection;
    viewport: Viewport;
  }) => Promise<{ documentId: string }>;
  readonly deleteScenario?: (input: {
    documentId: string;
    scenarioDocumentId: string;
  }) => Promise<void>;
  readonly navigateToWorkbook?: (input: {
    documentId: string;
    sheetName?: string | null;
    address?: string | null;
  }) => void;
}) {
  const {
    currentUserId,
    documentId,
    enabled,
    getCurrentViewport,
    selection,
    zero,
    createScenario = ({ documentId: sourceDocumentId, name, selection: nextSelection, viewport }) =>
      createWorkbookScenarioRequest({
        documentId: sourceDocumentId,
        name,
        sheetName: nextSelection.sheetName,
        address: nextSelection.address,
        viewport,
      }),
    deleteScenario = deleteWorkbookScenarioRequest,
    navigateToWorkbook = defaultNavigateToWorkbook,
  } = input;
  const scenarios = useWorkbookScenarios({
    documentId,
    currentUserId,
    zero,
    enabled,
  });
  const scenarioContext = useWorkbookScenarioContext({
    documentId,
    zero,
    enabled,
  });
  const [isOpen, setIsOpen] = useState(false);
  const [draftName, setDraftName] = useState("");
  const [isCreating, setIsCreating] = useState(false);
  const [deletingDocumentIds, setDeletingDocumentIds] = useState<readonly string[]>([]);
  const scenarioCount = Math.min(scenarios.length, 99);

  const openScenario = useCallback(
    (scenario: Pick<WorkbookScenarioEntry, "documentId" | "sheetName" | "address">) => {
      navigateToWorkbook({
        documentId: scenario.documentId,
        sheetName: scenario.sheetName,
        address: scenario.address,
      });
    },
    [navigateToWorkbook],
  );

  const createScenarioBranch = useCallback(async () => {
    if (!enabled || isCreating) {
      return;
    }
    setIsCreating(true);
    try {
      const nextScenario = await createScenario({
        documentId,
        name: draftName.trim() || defaultWorkbookScenarioName(selection),
        selection,
        viewport: getCurrentViewport(),
      });
      setDraftName("");
      navigateToWorkbook({
        documentId: nextScenario.documentId,
        sheetName: selection.sheetName,
        address: selection.address,
      });
    } finally {
      setIsCreating(false);
    }
  }, [
    createScenario,
    documentId,
    draftName,
    enabled,
    getCurrentViewport,
    isCreating,
    navigateToWorkbook,
    selection,
  ]);

  const removeScenario = useCallback(
    async (scenario: WorkbookScenarioEntry) => {
      if (!enabled || deletingDocumentIds.includes(scenario.documentId)) {
        return;
      }
      setDeletingDocumentIds((current) => [...current, scenario.documentId]);
      try {
        await deleteScenario({
          documentId,
          scenarioDocumentId: scenario.documentId,
        });
      } finally {
        setDeletingDocumentIds((current) =>
          current.filter((documentIdEntry) => documentIdEntry !== scenario.documentId),
        );
      }
    },
    [deleteScenario, deletingDocumentIds, documentId, enabled],
  );

  const scenariosToggle = useMemo(
    () => (
      <button
        aria-controls="workbook-scenarios-panel"
        aria-expanded={isOpen}
        aria-label={`Show workbook scratchpads (${scenarios.length})`}
        className="inline-flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
        data-testid="workbook-scenarios-toggle"
        type="button"
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        <span>Scratchpads</span>
        <span className="inline-flex min-w-5 items-center justify-center rounded-full bg-[var(--wb-surface-subtle)] px-1.5 text-[11px] font-semibold text-[var(--wb-text)]">
          {scenarioCount}
        </span>
      </button>
    ),
    [isOpen, scenarioCount, scenarios.length],
  );

  const scenarioStatus = useMemo(() => {
    if (!scenarioContext) {
      return null;
    }
    return (
      <div className="inline-flex items-center gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-1.5 text-[12px] text-[var(--wb-accent)] shadow-[var(--wb-shadow-sm)]">
        <span className="font-semibold">{scenarioContext.name}</span>
        <span>Scratchpad from r{scenarioContext.baseRevision}</span>
        <button
          className="inline-flex h-6 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-accent-ring)] bg-[var(--wb-surface)] px-2 text-[11px] font-medium text-[var(--wb-accent)]"
          type="button"
          onClick={() => {
            navigateToWorkbook({
              documentId: scenarioContext.workbookId,
              sheetName: scenarioContext.sheetName,
              address: scenarioContext.address,
            });
          }}
        >
          Return to source
        </button>
      </div>
    );
  }, [navigateToWorkbook, scenarioContext]);

  const scenariosPanel = useMemo(
    () => (
      <WorkbookScenariosPanel
        deletingDocumentIds={deletingDocumentIds}
        draftName={draftName}
        isCreating={isCreating}
        isOpen={isOpen}
        scenarios={scenarios}
        onClose={() => {
          setIsOpen(false);
        }}
        onCreate={() => {
          void createScenarioBranch();
        }}
        onDelete={(scenario) => {
          void removeScenario(scenario);
        }}
        onDraftNameChange={setDraftName}
        onOpen={openScenario}
      />
    ),
    [
      createScenarioBranch,
      deletingDocumentIds,
      draftName,
      isCreating,
      isOpen,
      openScenario,
      removeScenario,
      scenarios,
    ],
  );

  return {
    scenarioStatus,
    scenariosPanel,
    scenariosToggle,
  };
}
