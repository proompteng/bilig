import { useEffect, useMemo, useState } from "react";
import { queries } from "@bilig/zero-sync";
import {
  normalizeWorkbookScenarioEntry,
  normalizeWorkbookScenarioRows,
  selectWorkbookScenarioEntries,
  type WorkbookScenarioEntry,
} from "./workbook-scenarios-model.js";

interface ZeroLiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

export interface ZeroWorkbookScenarioQuerySource {
  materialize(query: unknown): unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isZeroLiveView<T>(value: unknown): value is ZeroLiveView<T> {
  return (
    isRecord(value) &&
    "data" in value &&
    typeof value["addListener"] === "function" &&
    typeof value["destroy"] === "function"
  );
}

export function useWorkbookScenarios(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly zero: ZeroWorkbookScenarioQuerySource;
  readonly enabled: boolean;
}): readonly WorkbookScenarioEntry[] {
  const { currentUserId, documentId, enabled, zero } = input;
  const [rows, setRows] = useState(
    [] as readonly ReturnType<typeof normalizeWorkbookScenarioRows>[number][],
  );

  useEffect(() => {
    if (!enabled) {
      setRows([]);
      return;
    }
    const view = zero.materialize(queries.workbookScenario.byWorkbook({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook scenarios query returned an invalid live view");
    }
    const publishRows = (value: unknown) => {
      setRows(normalizeWorkbookScenarioRows(value));
    };
    publishRows(view.data);
    const cleanup = view.addListener((value) => {
      publishRows(value);
    });
    return () => {
      cleanup();
      view.destroy();
    };
  }, [documentId, enabled, zero]);

  return useMemo(
    () =>
      selectWorkbookScenarioEntries({
        rows,
        currentUserId,
      }),
    [currentUserId, rows],
  );
}

export function useWorkbookScenarioContext(input: {
  readonly documentId: string;
  readonly zero: ZeroWorkbookScenarioQuerySource;
  readonly enabled: boolean;
}): WorkbookScenarioEntry | null {
  const { documentId, enabled, zero } = input;
  const [entry, setEntry] = useState<WorkbookScenarioEntry | null>(null);

  useEffect(() => {
    if (!enabled) {
      setEntry(null);
      return;
    }
    const view = zero.materialize(queries.workbookScenario.byDocument({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook scenario context query returned an invalid live view");
    }
    const publishEntry = (value: unknown) => {
      setEntry(normalizeWorkbookScenarioEntry(value));
    };
    publishEntry(view.data);
    const cleanup = view.addListener((value) => {
      publishEntry(value);
    });
    return () => {
      cleanup();
      view.destroy();
    };
  }, [documentId, enabled, zero]);

  return entry;
}
