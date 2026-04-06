import { useEffect, useMemo, useState } from "react";
import { queries } from "@bilig/zero-sync";
import {
  normalizeWorkbookVersionRows,
  selectWorkbookVersionEntries,
  type WorkbookVersionEntry,
} from "./workbook-versions-model.js";

interface ZeroLiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

export interface ZeroWorkbookVersionQuerySource {
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

export function useWorkbookVersions(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly zero: ZeroWorkbookVersionQuerySource;
  readonly enabled: boolean;
}): readonly WorkbookVersionEntry[] {
  const { currentUserId, documentId, enabled, zero } = input;
  const [rows, setRows] = useState(
    [] as readonly ReturnType<typeof normalizeWorkbookVersionRows>[number][],
  );

  useEffect(() => {
    if (!enabled) {
      setRows([]);
      return;
    }
    const view = zero.materialize(queries.workbookVersion.byWorkbook({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook versions query returned an invalid live view");
    }
    const publishRows = (value: unknown) => {
      setRows(normalizeWorkbookVersionRows(value));
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
      selectWorkbookVersionEntries({
        rows,
        currentUserId,
      }),
    [currentUserId, rows],
  );
}
