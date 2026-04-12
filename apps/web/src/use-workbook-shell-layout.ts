import { useCallback, useEffect, useMemo, useState } from "react";

const STORAGE_KEY_PREFIX = "bilig:workbook-shell-layout:";

export const DEFAULT_WORKBOOK_SIDE_PANEL_WIDTH = 320;
const MIN_WORKBOOK_SIDE_PANEL_WIDTH = 280;
const MAX_WORKBOOK_SIDE_PANEL_WIDTH = 420;
const WORKBOOK_SIDE_PANEL_VIEWPORT_FRACTION = 0.42;

interface StoredWorkbookShellLayout {
  sidePanelOpen?: boolean;
  sidePanelTab?: string;
  sidePanelWidth?: number;
}

interface WorkbookShellLayoutState {
  isSidePanelOpen: boolean;
  activeSidePanelTab: string | null;
  sidePanelWidth: number;
}

function storageKey(scope: string): string {
  return `${STORAGE_KEY_PREFIX}${scope}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function clampWorkbookSidePanelWidth(width: number): number {
  const viewportWidth = typeof window === "undefined" ? null : window.innerWidth;
  const viewportAwareMax =
    viewportWidth && Number.isFinite(viewportWidth)
      ? Math.min(
          MAX_WORKBOOK_SIDE_PANEL_WIDTH,
          Math.max(
            MIN_WORKBOOK_SIDE_PANEL_WIDTH,
            Math.round(viewportWidth * WORKBOOK_SIDE_PANEL_VIEWPORT_FRACTION),
          ),
        )
      : MAX_WORKBOOK_SIDE_PANEL_WIDTH;
  return Math.min(viewportAwareMax, Math.max(MIN_WORKBOOK_SIDE_PANEL_WIDTH, Math.round(width)));
}

function normalizeStoredWorkbookShellLayout(
  value: unknown,
  availableTabs: readonly string[],
  defaultTab: string | null,
): WorkbookShellLayoutState {
  const activeSidePanelTab =
    isRecord(value) && typeof value["sidePanelTab"] === "string"
      ? value["sidePanelTab"]
      : defaultTab;
  const sidePanelWidth =
    isRecord(value) && typeof value["sidePanelWidth"] === "number"
      ? clampWorkbookSidePanelWidth(value["sidePanelWidth"])
      : DEFAULT_WORKBOOK_SIDE_PANEL_WIDTH;
  const isSidePanelOpen =
    isRecord(value) && value["sidePanelOpen"] === true && activeSidePanelTab !== null;
  const resolvedActiveTab =
    activeSidePanelTab && availableTabs.includes(activeSidePanelTab)
      ? activeSidePanelTab
      : defaultTab;
  return {
    isSidePanelOpen: isSidePanelOpen && resolvedActiveTab !== null,
    activeSidePanelTab: resolvedActiveTab,
    sidePanelWidth,
  };
}

function loadPersistedWorkbookShellLayout(
  scope: string,
  availableTabs: readonly string[],
  defaultTab: string | null,
): WorkbookShellLayoutState {
  try {
    const raw = window.localStorage.getItem(storageKey(scope));
    if (!raw) {
      return normalizeStoredWorkbookShellLayout(null, availableTabs, defaultTab);
    }
    const parsed = JSON.parse(raw) as unknown;
    return normalizeStoredWorkbookShellLayout(parsed, availableTabs, defaultTab);
  } catch {
    return normalizeStoredWorkbookShellLayout(null, availableTabs, defaultTab);
  }
}

function persistWorkbookShellLayout(scope: string, layout: WorkbookShellLayoutState): void {
  try {
    const stored: StoredWorkbookShellLayout = {
      sidePanelOpen: layout.isSidePanelOpen,
      sidePanelWidth: clampWorkbookSidePanelWidth(layout.sidePanelWidth),
    };
    if (layout.activeSidePanelTab !== null) {
      stored.sidePanelTab = layout.activeSidePanelTab;
    }
    window.localStorage.setItem(storageKey(scope), JSON.stringify(stored));
  } catch {
    // Ignore storage failures and keep the shell usable.
  }
}

export function useWorkbookShellLayout(input: {
  documentId: string;
  persistenceKey?: string;
  availableTabs: readonly string[];
  defaultTab?: string | null;
}) {
  const { availableTabs, defaultTab, documentId, persistenceKey } = input;
  const resolvedPersistenceKey = persistenceKey ?? documentId;
  const resolvedDefaultTab = useMemo(() => {
    if (availableTabs.length === 0) {
      return null;
    }
    if (defaultTab === null) {
      return null;
    }
    return defaultTab === undefined ? null : availableTabs.includes(defaultTab) ? defaultTab : null;
  }, [availableTabs, defaultTab]);
  const [layout, setLayout] = useState<WorkbookShellLayoutState>(() =>
    typeof window === "undefined"
      ? {
          isSidePanelOpen: false,
          activeSidePanelTab: resolvedDefaultTab,
          sidePanelWidth: DEFAULT_WORKBOOK_SIDE_PANEL_WIDTH,
        }
      : loadPersistedWorkbookShellLayout(resolvedPersistenceKey, availableTabs, resolvedDefaultTab),
  );

  useEffect(() => {
    setLayout((current) => {
      const activeSidePanelTab =
        current.activeSidePanelTab && availableTabs.includes(current.activeSidePanelTab)
          ? current.activeSidePanelTab
          : resolvedDefaultTab;
      const nextLayout = {
        isSidePanelOpen: current.isSidePanelOpen && activeSidePanelTab !== null,
        activeSidePanelTab,
        sidePanelWidth: clampWorkbookSidePanelWidth(current.sidePanelWidth),
      };
      if (
        current.isSidePanelOpen === nextLayout.isSidePanelOpen &&
        current.activeSidePanelTab === nextLayout.activeSidePanelTab &&
        current.sidePanelWidth === nextLayout.sidePanelWidth
      ) {
        return current;
      }
      return nextLayout;
    });
  }, [availableTabs, resolvedDefaultTab]);

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }
    persistWorkbookShellLayout(resolvedPersistenceKey, layout);
  }, [layout, resolvedPersistenceKey]);

  const setActiveSidePanelTab = useCallback(
    (nextTab: string) => {
      setLayout((current) => ({
        ...current,
        activeSidePanelTab: availableTabs.includes(nextTab) ? nextTab : current.activeSidePanelTab,
      }));
    },
    [availableTabs],
  );

  const openSidePanel = useCallback(
    (nextTab?: string) => {
      setLayout((current) => {
        const activeSidePanelTab =
          nextTab && availableTabs.includes(nextTab)
            ? nextTab
            : (current.activeSidePanelTab ?? resolvedDefaultTab);
        return {
          ...current,
          isSidePanelOpen: activeSidePanelTab !== null,
          activeSidePanelTab,
        };
      });
    },
    [availableTabs, resolvedDefaultTab],
  );

  const closeSidePanel = useCallback(() => {
    setLayout((current) => ({
      ...current,
      isSidePanelOpen: false,
    }));
  }, []);

  const toggleSidePanel = useCallback(
    (nextTab?: string) => {
      setLayout((current) => {
        const desiredTab =
          nextTab && availableTabs.includes(nextTab)
            ? nextTab
            : (current.activeSidePanelTab ?? resolvedDefaultTab);
        if (desiredTab === null) {
          return {
            ...current,
            isSidePanelOpen: false,
            activeSidePanelTab: null,
          };
        }
        if (current.isSidePanelOpen && current.activeSidePanelTab === desiredTab) {
          return {
            ...current,
            isSidePanelOpen: false,
            activeSidePanelTab: desiredTab,
          };
        }
        return {
          ...current,
          isSidePanelOpen: true,
          activeSidePanelTab: desiredTab,
        };
      });
    },
    [availableTabs, resolvedDefaultTab],
  );

  const setSidePanelWidth = useCallback((nextWidth: number) => {
    setLayout((current) => ({
      ...current,
      sidePanelWidth: clampWorkbookSidePanelWidth(nextWidth),
    }));
  }, []);

  return {
    activeSidePanelTab: layout.activeSidePanelTab,
    closeSidePanel,
    isSidePanelOpen: layout.isSidePanelOpen,
    openSidePanel,
    setActiveSidePanelTab,
    setSidePanelWidth,
    sidePanelWidth: layout.sidePanelWidth,
    toggleSidePanel,
  };
}
