import { useMemo, type ReactNode } from "react";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";
import { WorkbookSideRailTabs } from "./WorkbookSideRailTabs.js";
import { useWorkbookChangesPane } from "./use-workbook-changes-pane.js";
import { useWorkbookAgentPane } from "./use-workbook-agent-pane.js";
import { useWorkbookPresence } from "./use-workbook-presence.js";

type WorkbookPanelsZeroSource = Parameters<typeof useWorkbookPresence>[0]["zero"] &
  Parameters<typeof useWorkbookChangesPane>[0]["zero"];

type WorkbookAgentContextGetter = Parameters<typeof useWorkbookAgentPane>[0]["getContext"];
type WorkbookAgentPreviewBundle = (
  bundle: WorkbookAgentCommandBundle,
) => ReturnType<Parameters<typeof useWorkbookAgentPane>[0]["previewBundle"]>;

export function useWorkbookAppPanels(input: {
  documentId: string;
  replicaId: string;
  selection: WorkerRuntimeSelection;
  sheetNames: readonly string[];
  zero: WorkbookPanelsZeroSource;
  runtimeReady: boolean;
  zeroConfigured: boolean;
  remoteSyncAvailable: boolean;
  toolbarHeaderStatus: ReactNode;
  selectAddress: (sheetName: string, address: string) => void;
  getAgentContext: WorkbookAgentContextGetter;
  previewAgentBundle: WorkbookAgentPreviewBundle;
}) {
  const {
    documentId,
    getAgentContext,
    previewAgentBundle,
    remoteSyncAvailable,
    replicaId,
    runtimeReady,
    selection,
    selectAddress,
    sheetNames,
    toolbarHeaderStatus,
    zero,
    zeroConfigured,
  } = input;

  const collaborators = useWorkbookPresence({
    documentId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const { changeCount, changesPanel } = useWorkbookChangesPane({
    documentId,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured,
    onJump: (sheetName, address) => {
      selectAddress(sheetName, address);
    },
  });

  const { agentPanel, agentError, clearAgentError, pendingCommandCount, previewRanges } =
    useWorkbookAgentPane({
      documentId,
      enabled: runtimeReady,
      getContext: getAgentContext,
      previewBundle: previewAgentBundle,
    });

  const headerStatus = useMemo(
    () => (
      <div className="flex flex-wrap items-center justify-end gap-1.5">
        {toolbarHeaderStatus}
        {collaborators.length > 0 ? (
          <WorkbookPresenceBar
            collaborators={collaborators}
            onJump={(sheetName, address) => {
              selectAddress(sheetName, address);
            }}
          />
        ) : null}
      </div>
    ),
    [collaborators, selectAddress, toolbarHeaderStatus],
  );

  const sideRail = useMemo(
    () => (
      <WorkbookSideRailTabs
        defaultValue="assistant"
        tabs={[
          {
            value: "assistant",
            label: "Assistant",
            count: pendingCommandCount > 0 ? pendingCommandCount : undefined,
            panel: agentPanel,
          },
          {
            value: "changes",
            label: "Changes",
            count: changeCount,
            panel: changesPanel,
          },
        ]}
      />
    ),
    [agentPanel, changeCount, changesPanel, pendingCommandCount],
  );

  return {
    agentError,
    agentPanel,
    changeCount,
    changesPanel,
    clearAgentError,
    headerStatus,
    pendingCommandCount,
    previewRanges,
    sideRail,
  };
}
