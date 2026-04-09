import { useCallback, useEffect, useMemo, useState } from "react";

const STORAGE_KEY_PREFIX = "bilig:workbook-shell-layout:";

export const DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH = 344;
export const MIN_WORKBOOK_SIDE_RAIL_WIDTH = 304;
export const MAX_WORKBOOK_SIDE_RAIL_WIDTH = 520;

interface StoredWorkbookShellLayout {
  sideRailOpen?: boolean;
  sideRailTab?: string;
  sideRailWidth?: number;
}

interface WorkbookShellLayoutState {
  isSideRailOpen: boolean;
  activeSideRailTab: string | null;
  sideRailWidth: number;
}

function storageKey(documentId: string): string {
  return `${STORAGE_KEY_PREFIX}${documentId}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function clampWorkbookSideRailWidth(width: number): number {
  return Math.min(
    MAX_WORKBOOK_SIDE_RAIL_WIDTH,
    Math.max(MIN_WORKBOOK_SIDE_RAIL_WIDTH, Math.round(width)),
  );
}

function normalizeStoredWorkbookShellLayout(
  value: unknown,
  availableTabs: readonly string[],
  defaultTab: string | null,
): WorkbookShellLayoutState {
  const activeSideRailTab =
    isRecord(value) && typeof value["sideRailTab"] === "string" ? value["sideRailTab"] : defaultTab;
  const sideRailWidth =
    isRecord(value) && typeof value["sideRailWidth"] === "number"
      ? clampWorkbookSideRailWidth(value["sideRailWidth"])
      : DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH;
  const isSideRailOpen =
    isRecord(value) && value["sideRailOpen"] === true && activeSideRailTab !== null;
  const resolvedActiveTab =
    activeSideRailTab && availableTabs.includes(activeSideRailTab) ? activeSideRailTab : defaultTab;
  return {
    isSideRailOpen: isSideRailOpen && resolvedActiveTab !== null,
    activeSideRailTab: resolvedActiveTab,
    sideRailWidth,
  };
}

export function loadPersistedWorkbookShellLayout(
  documentId: string,
  availableTabs: readonly string[],
  defaultTab: string | null,
): WorkbookShellLayoutState {
  try {
    const raw = window.localStorage.getItem(storageKey(documentId));
    if (!raw) {
      return normalizeStoredWorkbookShellLayout(null, availableTabs, defaultTab);
    }
    const parsed = JSON.parse(raw) as unknown;
    return normalizeStoredWorkbookShellLayout(parsed, availableTabs, defaultTab);
  } catch {
    return normalizeStoredWorkbookShellLayout(null, availableTabs, defaultTab);
  }
}

export function persistWorkbookShellLayout(
  documentId: string,
  layout: WorkbookShellLayoutState,
): void {
  try {
    const stored: StoredWorkbookShellLayout = {
      sideRailOpen: layout.isSideRailOpen,
      sideRailWidth: clampWorkbookSideRailWidth(layout.sideRailWidth),
    };
    if (layout.activeSideRailTab !== null) {
      stored.sideRailTab = layout.activeSideRailTab;
    }
    window.localStorage.setItem(storageKey(documentId), JSON.stringify(stored));
  } catch {
    // Ignore storage failures and keep the shell usable.
  }
}

export function useWorkbookShellLayout(input: {
  documentId: string;
  availableTabs: readonly string[];
  defaultTab?: string;
}) {
  const { documentId, availableTabs, defaultTab } = input;
  const resolvedDefaultTab = useMemo(() => {
    if (availableTabs.length === 0) {
      return null;
    }
    return defaultTab && availableTabs.includes(defaultTab) ? defaultTab : availableTabs[0]!;
  }, [availableTabs, defaultTab]);
  const [layout, setLayout] = useState<WorkbookShellLayoutState>(() =>
    typeof window === "undefined"
      ? {
          isSideRailOpen: false,
          activeSideRailTab: resolvedDefaultTab,
          sideRailWidth: DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH,
        }
      : loadPersistedWorkbookShellLayout(documentId, availableTabs, resolvedDefaultTab),
  );

  useEffect(() => {
    setLayout((current) => {
      const activeSideRailTab =
        current.activeSideRailTab && availableTabs.includes(current.activeSideRailTab)
          ? current.activeSideRailTab
          : resolvedDefaultTab;
      const nextLayout = {
        isSideRailOpen: current.isSideRailOpen && activeSideRailTab !== null,
        activeSideRailTab,
        sideRailWidth: clampWorkbookSideRailWidth(current.sideRailWidth),
      };
      if (
        current.isSideRailOpen === nextLayout.isSideRailOpen &&
        current.activeSideRailTab === nextLayout.activeSideRailTab &&
        current.sideRailWidth === nextLayout.sideRailWidth
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
    persistWorkbookShellLayout(documentId, layout);
  }, [documentId, layout]);

  const setActiveSideRailTab = useCallback(
    (nextTab: string) => {
      setLayout((current) => ({
        ...current,
        activeSideRailTab: availableTabs.includes(nextTab) ? nextTab : current.activeSideRailTab,
      }));
    },
    [availableTabs],
  );

  const openSideRail = useCallback(
    (nextTab?: string) => {
      setLayout((current) => {
        const activeSideRailTab =
          nextTab && availableTabs.includes(nextTab)
            ? nextTab
            : (current.activeSideRailTab ?? resolvedDefaultTab);
        return {
          ...current,
          isSideRailOpen: activeSideRailTab !== null,
          activeSideRailTab,
        };
      });
    },
    [availableTabs, resolvedDefaultTab],
  );

  const closeSideRail = useCallback(() => {
    setLayout((current) => ({
      ...current,
      isSideRailOpen: false,
    }));
  }, []);

  const toggleSideRail = useCallback(
    (nextTab?: string) => {
      setLayout((current) => {
        const desiredTab =
          nextTab && availableTabs.includes(nextTab)
            ? nextTab
            : (current.activeSideRailTab ?? resolvedDefaultTab);
        if (desiredTab === null) {
          return {
            ...current,
            isSideRailOpen: false,
            activeSideRailTab: null,
          };
        }
        if (current.isSideRailOpen && current.activeSideRailTab === desiredTab) {
          return {
            ...current,
            isSideRailOpen: false,
            activeSideRailTab: desiredTab,
          };
        }
        return {
          ...current,
          isSideRailOpen: true,
          activeSideRailTab: desiredTab,
        };
      });
    },
    [availableTabs, resolvedDefaultTab],
  );

  const setSideRailWidth = useCallback((nextWidth: number) => {
    setLayout((current) => ({
      ...current,
      sideRailWidth: clampWorkbookSideRailWidth(nextWidth),
    }));
  }, []);

  return {
    activeSideRailTab: layout.activeSideRailTab,
    closeSideRail,
    isSideRailOpen: layout.isSideRailOpen,
    openSideRail,
    setActiveSideRailTab,
    setSideRailWidth,
    sideRailWidth: layout.sideRailWidth,
    toggleSideRail,
  };
}
