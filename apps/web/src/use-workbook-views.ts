import { useEffect, useMemo, useState } from "react";
import { queries } from "@bilig/zero-sync";
import {
  normalizeWorkbookSheetViewRows,
  selectWorkbookSheetViewEntries,
  type WorkbookSheetViewEntry,
} from "./workbook-views-model.js";

interface ZeroLiveView<T> {
  readonly data: T;
  addListener(listener: (value: T) => void): () => void;
  destroy(): void;
}

export interface ZeroWorkbookSheetViewQuerySource {
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

export function useWorkbookViews(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly sheetNames: readonly string[];
  readonly zero: ZeroWorkbookSheetViewQuerySource;
  readonly enabled: boolean;
}): readonly WorkbookSheetViewEntry[] {
  const { currentUserId, documentId, enabled, sheetNames, zero } = input;
  const [rows, setRows] = useState(
    [] as readonly ReturnType<typeof normalizeWorkbookSheetViewRows>[number][],
  );

  useEffect(() => {
    if (!enabled) {
      setRows([]);
      return;
    }
    const view = zero.materialize(queries.sheetView.byWorkbook({ documentId }));
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error("Zero workbook views query returned an invalid live view");
    }
    const publishRows = (value: unknown) => {
      setRows(normalizeWorkbookSheetViewRows(value));
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
      selectWorkbookSheetViewEntries({
        rows,
        currentUserId,
        knownSheetNames: sheetNames,
      }),
    [currentUserId, rows, sheetNames],
  );
}
