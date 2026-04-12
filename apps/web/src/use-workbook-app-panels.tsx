import { useEffect, useMemo, useRef, type ReactNode } from "react";
import { Button } from "@base-ui/react/button";
import { Tabs } from "@base-ui/react/tabs";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import {
  workbookHeaderActionButtonClass,
  workbookHeaderCountClass,
} from "./workbook-header-controls.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";
import {
  panelCountClass,
  panelIndicatorClass,
  panelListClass,
  panelContentClass,
  panelRootClass,
  panelTabClass,
  type WorkbookSidePanelTabDefinition,
} from "./WorkbookSidePanelTabs.js";
import { cn } from "./cn.js";
import { useWorkbookAgentPane } from "./use-workbook-agent-pane.js";
import { useWorkbookPresence } from "./use-workbook-presence.js";
import { useWorkbookShellLayout } from "./use-workbook-shell-layout.js";

type WorkbookPanelsZeroSource = Parameters<typeof useWorkbookPresence>[0]["zero"];

type WorkbookAgentContextGetter = Parameters<typeof useWorkbookAgentPane>[0]["getContext"];
type WorkbookAgentPreviewCommandBundle = (
  bundle: WorkbookAgentCommandBundle,
) => ReturnType<Parameters<typeof useWorkbookAgentPane>[0]["previewCommandBundle"]>;

export function useWorkbookAppPanels(input: {
  documentId: string;
  currentUserId: string;
  presenceClientId: string;
  replicaId: string;
  selection: WorkerRuntimeSelection;
  sheetNames: readonly string[];
  zero: WorkbookPanelsZeroSource;
  runtimeReady: boolean;
  zeroConfigured: boolean;
  remoteSyncAvailable: boolean;
  changeCount: number;
  changesPanel: ReactNode;
  selectAddress: (sheetName: string, address: string) => void;
  getAgentContext: WorkbookAgentContextGetter;
  previewAgentCommandBundle: WorkbookAgentPreviewCommandBundle;
}) {
  const {
    changeCount,
    changesPanel,
    currentUserId,
    documentId,
    getAgentContext,
    presenceClientId,
    previewAgentCommandBundle,
    remoteSyncAvailable,
    replicaId,
    runtimeReady,
    selection,
    selectAddress,
    sheetNames,
    zero,
    zeroConfigured,
  } = input;

  const collaborators = useWorkbookPresence({
    documentId,
    currentUserId,
    currentPresenceClientId: presenceClientId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const {
    agentPanel,
    agentError,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    startNewThread,
  } = useWorkbookAgentPane({
    currentUserId,
    documentId,
    enabled: runtimeReady,
    getContext: getAgentContext,
    previewCommandBundle: previewAgentCommandBundle,
    zero,
    zeroEnabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const sidePanelTabs = useMemo<readonly WorkbookSidePanelTabDefinition[]>(
    () => [
      {
        value: "assistant",
        label: "Assistant",
        count: pendingCommandCount > 0 ? pendingCommandCount : undefined,
        panel: agentPanel,
      },
      {
        value: "changes",
        label: "Changes",
        count: changeCount > 0 ? changeCount : undefined,
        panel: changesPanel,
      },
    ],
    [agentPanel, changeCount, changesPanel, pendingCommandCount],
  );
  const visibleSidePanelTabs = useMemo(
    () => sidePanelTabs.filter((tab) => tab.panel != null),
    [sidePanelTabs],
  );
  const {
    activeSidePanelTab,
    isSidePanelOpen,
    openSidePanel,
    setActiveSidePanelTab,
    setSidePanelWidth,
    sidePanelWidth,
    toggleSidePanel,
  } = useWorkbookShellLayout({
    documentId,
    persistenceKey: `${documentId}:${currentUserId}`,
    availableTabs: visibleSidePanelTabs.map((tab) => tab.value),
    defaultTab: null,
  });
  const sidePanelId = `workbook-side-panel-${documentId}`;
  const previousPendingCommandCountRef = useRef(pendingCommandCount);

  useEffect(() => {
    const hadPendingCommands = previousPendingCommandCountRef.current > 0;
    const hasPendingCommands = pendingCommandCount > 0;
    previousPendingCommandCountRef.current = pendingCommandCount;
    if (!hasPendingCommands || hadPendingCommands) {
      return;
    }
    if (!visibleSidePanelTabs.some((tab) => tab.value === "assistant")) {
      return;
    }
    openSidePanel("assistant");
  }, [openSidePanel, pendingCommandCount, visibleSidePanelTabs]);

  const sidePanelToggleControls = useMemo(
    () => (
      <div
        className="inline-flex items-center gap-1 rounded-md border border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] p-1 shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
        data-testid="workbook-side-panel-toggle-group"
      >
        {visibleSidePanelTabs.map((tab) => {
          const active = isSidePanelOpen && activeSidePanelTab === tab.value;
          return (
            <Button
              aria-controls={sidePanelId}
              aria-expanded={active}
              aria-pressed={active}
              className={workbookHeaderActionButtonClass({
                active,
                grouped: true,
              })}
              data-testid={`workbook-side-panel-toggle-${tab.value}`}
              key={tab.value}
              type="button"
              onClick={() => {
                toggleSidePanel(tab.value);
              }}
            >
              <span>{tab.label}</span>
              {typeof tab.count === "number" ? (
                <span className={workbookHeaderCountClass}>{String(Math.min(tab.count, 99))}</span>
              ) : null}
            </Button>
          );
        })}
      </div>
    ),
    [activeSidePanelTab, isSidePanelOpen, sidePanelId, toggleSidePanel, visibleSidePanelTabs],
  );

  const toolbarTrailingContent = useMemo(() => {
    if (visibleSidePanelTabs.length === 0 && collaborators.length === 0) {
      return null;
    }
    return (
      <>
        {visibleSidePanelTabs.length > 0 ? sidePanelToggleControls : null}
        {collaborators.length > 0 ? (
          <WorkbookPresenceBar
            collaborators={collaborators}
            onJump={(sheetName, address) => {
              selectAddress(sheetName, address);
            }}
          />
        ) : null}
      </>
    );
  }, [collaborators, selectAddress, sidePanelToggleControls, visibleSidePanelTabs.length]);

  const sidePanel = useMemo(
    () =>
      isSidePanelOpen &&
      activeSidePanelTab &&
      visibleSidePanelTabs.some((tab) => tab.value === activeSidePanelTab) ? (
        <Tabs.Root
          className={panelRootClass()}
          value={activeSidePanelTab}
          onValueChange={(nextValue) => {
            setActiveSidePanelTab(String(nextValue));
          }}
        >
          <Tabs.List aria-label="Workbook panels" className={panelListClass()}>
            <div className="flex min-w-0 flex-1 items-end gap-1">
              {visibleSidePanelTabs.map((tab) => (
                <Tabs.Tab
                  className={(state) => panelTabClass({ active: state.active })}
                  data-testid={`workbook-side-panel-tab-${tab.value}`}
                  key={tab.value}
                  value={tab.value}
                >
                  <span>{tab.label}</span>
                  {typeof tab.count === "number" ? (
                    <span
                      className={cn(
                        panelCountClass({
                          active: activeSidePanelTab === tab.value,
                        }),
                      )}
                    >
                      {String(Math.min(tab.count, 99))}
                    </span>
                  ) : null}
                </Tabs.Tab>
              ))}
            </div>
            <Button
              className={cn(
                workbookHeaderActionButtonClass({ active: false }),
                "ml-auto shrink-0 self-center",
              )}
              data-testid="workbook-agent-new-thread"
              type="button"
              onClick={startNewThread}
            >
              New thread
            </Button>
            <Tabs.Indicator className={panelIndicatorClass()} renderBeforeHydration />
          </Tabs.List>
          {visibleSidePanelTabs.map((tab) => (
            <Tabs.Panel
              className={panelContentClass()}
              data-testid={`workbook-side-panel-panel-${tab.value}`}
              keepMounted
              key={tab.value}
              value={tab.value}
            >
              {tab.panel}
            </Tabs.Panel>
          ))}
        </Tabs.Root>
      ) : null,
    [
      activeSidePanelTab,
      isSidePanelOpen,
      setActiveSidePanelTab,
      startNewThread,
      visibleSidePanelTabs,
    ],
  );

  return {
    agentError,
    agentPanel,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    sidePanelId,
    sidePanel,
    setSidePanelWidth,
    sidePanelWidth,
    toolbarTrailingContent,
  };
}
