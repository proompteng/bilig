import { useCallback, type MutableRefObject } from "react";
import type { CommitOp } from "@bilig/core";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import { createNextSheetName, normalizeSheetNameKey } from "./worker-workbook-app-model.js";

export function useWorkbookSheetActions(input: {
  sheetNames: readonly string[];
  selectionRef: MutableRefObject<WorkerRuntimeSelection>;
  invokeMutation: (method: "renderCommit", ops: CommitOp[]) => Promise<void>;
  selectAddress: (sheetName: string, address: string) => void;
  reportRuntimeError: (error: unknown) => void;
}) {
  const { invokeMutation, reportRuntimeError, selectAddress, selectionRef, sheetNames } = input;

  const createSheet = useCallback(() => {
    const nextSheetName = createNextSheetName(sheetNames);
    void invokeMutation("renderCommit", [
      {
        kind: "upsertSheet",
        name: nextSheetName,
        order: sheetNames.length,
      } satisfies CommitOp,
    ])
      .then(() => selectAddress(nextSheetName, "A1"))
      .catch(reportRuntimeError);
  }, [invokeMutation, reportRuntimeError, selectAddress, sheetNames]);

  const renameSheet = useCallback(
    (currentName: string, nextName: string) => {
      const trimmedName = nextName.trim();
      if (trimmedName.length === 0 || trimmedName === currentName) {
        return;
      }
      const currentKey = normalizeSheetNameKey(currentName);
      const nextKey = normalizeSheetNameKey(trimmedName);
      if (
        sheetNames.some(
          (name) =>
            normalizeSheetNameKey(name) === nextKey && normalizeSheetNameKey(name) !== currentKey,
        )
      ) {
        return;
      }

      void invokeMutation("renderCommit", [
        {
          kind: "renameSheet",
          oldName: currentName,
          newName: trimmedName,
        } satisfies CommitOp,
      ])
        .then(() => {
          if (selectionRef.current.sheetName === currentName) {
            selectAddress(trimmedName, selectionRef.current.address);
          }
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, selectAddress, selectionRef, sheetNames],
  );

  const deleteSheet = useCallback(
    (targetName: string) => {
      if (sheetNames.length <= 1 || !sheetNames.includes(targetName)) {
        return;
      }

      const fallbackSheetName =
        selectionRef.current.sheetName === targetName
          ? (sheetNames.find((name) => name !== targetName) ?? null)
          : null;

      void invokeMutation("renderCommit", [
        { kind: "deleteSheet", name: targetName } satisfies CommitOp,
      ])
        .then(() => {
          if (fallbackSheetName) {
            selectAddress(fallbackSheetName, "A1");
          }
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, selectAddress, selectionRef, sheetNames],
  );

  return {
    createSheet,
    deleteSheet,
    renameSheet,
  };
}
