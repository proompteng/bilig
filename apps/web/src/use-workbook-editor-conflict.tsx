import {
  useCallback,
  useEffect,
  useMemo,
  type Dispatch,
  type MutableRefObject,
  type SetStateAction,
} from "react";
import type { CellSnapshot } from "@bilig/protocol";
import type { EditSelectionBehavior } from "@bilig/grid";
import { WorkbookEditorConflictBanner } from "./WorkbookEditorConflictBanner.js";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import {
  parseEditorInput,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
  type EditingMode,
  type ParsedEditorInput,
  type WorkbookEditorConflict,
} from "./worker-workbook-app-model.js";

export function useWorkbookEditorConflict(input: {
  editingMode: EditingMode;
  editorValue: string;
  editorConflict: WorkbookEditorConflict | null;
  setEditorConflict: Dispatch<SetStateAction<WorkbookEditorConflict | null>>;
  selectedCell: CellSnapshot;
  selection: WorkerRuntimeSelection;
  editorValueRef: MutableRefObject<string>;
  editorTargetRef: MutableRefObject<WorkerRuntimeSelection>;
  editorBaseSnapshotRef: MutableRefObject<CellSnapshot>;
  editingModeRef: MutableRefObject<EditingMode>;
  cloneLiveSelectedCell: (nextSelection?: WorkerRuntimeSelection) => CellSnapshot;
  completeEditNavigation: (targetSelection: WorkerRuntimeSelection) => WorkerRuntimeSelection;
  finishEditingWithAuthoritative: (targetSelection: WorkerRuntimeSelection) => void;
  resetEditorConflictTracking: (nextSelection?: WorkerRuntimeSelection) => void;
  applyParsedInput: (
    sheetName: string,
    address: string,
    parsed: ParsedEditorInput,
  ) => Promise<void>;
  reportRuntimeError: (error: unknown) => void;
  setEditorSelectionBehavior: Dispatch<SetStateAction<EditSelectionBehavior>>;
  setEditingMode: Dispatch<SetStateAction<EditingMode>>;
}) {
  const {
    applyParsedInput,
    cloneLiveSelectedCell,
    completeEditNavigation,
    editingMode,
    editingModeRef,
    editorConflict,
    editorBaseSnapshotRef,
    editorTargetRef,
    editorValue,
    editorValueRef,
    finishEditingWithAuthoritative,
    resetEditorConflictTracking,
    reportRuntimeError,
    selectedCell,
    selection,
    setEditorConflict,
    setEditingMode,
    setEditorSelectionBehavior,
  } = input;

  useEffect(() => {
    if (editingMode === "idle") {
      return;
    }
    const targetSelection = editorTargetRef.current;
    if (
      selection.sheetName !== targetSelection.sheetName ||
      selection.address !== targetSelection.address
    ) {
      return;
    }
    const authoritativeSnapshot = cloneLiveSelectedCell(targetSelection);
    const baseSnapshot = editorBaseSnapshotRef.current;
    const parsedDraft = parseEditorInput(editorValueRef.current);

    if (
      sameCellContent(baseSnapshot, authoritativeSnapshot) ||
      parsedEditorInputMatchesSnapshot(parsedDraft, authoritativeSnapshot)
    ) {
      setEditorConflict((current) => (current === null ? current : null));
      return;
    }

    setEditorConflict((current) => {
      const nextPhase =
        current?.sheetName === targetSelection.sheetName &&
        current?.address === targetSelection.address &&
        current.phase === "compare"
          ? "compare"
          : "badge";
      if (
        current?.sheetName === targetSelection.sheetName &&
        current?.address === targetSelection.address &&
        current.phase === nextPhase &&
        sameCellContent(current.baseSnapshot, baseSnapshot) &&
        sameCellContent(current.authoritativeSnapshot, authoritativeSnapshot)
      ) {
        return current;
      }
      return {
        sheetName: targetSelection.sheetName,
        address: targetSelection.address,
        phase: nextPhase,
        baseSnapshot,
        authoritativeSnapshot,
      };
    });
  }, [
    cloneLiveSelectedCell,
    editingMode,
    editorBaseSnapshotRef,
    editorTargetRef,
    editorValueRef,
    selectedCell,
    selection.address,
    selection.sheetName,
    setEditorConflict,
  ]);

  const reviewEditorConflict = useCallback(() => {
    const targetSelection = editorTargetRef.current;
    setEditorConflict((current) => {
      if (
        current?.sheetName !== targetSelection.sheetName ||
        current?.address !== targetSelection.address
      ) {
        return current;
      }
      return {
        ...current,
        phase: "compare",
        authoritativeSnapshot: cloneLiveSelectedCell(targetSelection),
      };
    });
  }, [cloneLiveSelectedCell, editorTargetRef, setEditorConflict]);

  const applyEditorConflictDraft = useCallback(() => {
    const targetSelection = editorTargetRef.current;
    const parsed = parseEditorInput(editorValueRef.current);
    const nextSelection = completeEditNavigation(targetSelection);
    setEditorSelectionBehavior("select-all");
    editingModeRef.current = "idle";
    setEditingMode("idle");
    resetEditorConflictTracking(nextSelection);
    void applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed).catch(
      reportRuntimeError,
    );
  }, [
    applyParsedInput,
    completeEditNavigation,
    editingModeRef,
    editorTargetRef,
    editorValueRef,
    reportRuntimeError,
    resetEditorConflictTracking,
    setEditingMode,
    setEditorSelectionBehavior,
  ]);

  const keepEditorConflictAuthoritative = useCallback(() => {
    finishEditingWithAuthoritative(editorTargetRef.current);
  }, [editorTargetRef, finishEditingWithAuthoritative]);

  const keepEditorConflictDraftLocal = useCallback(() => {
    setEditorConflict((current) => (current ? { ...current, phase: "badge" } : current));
  }, [setEditorConflict]);

  return useMemo(() => {
    if (editorConflict === null) {
      return null;
    }
    return (
      <WorkbookEditorConflictBanner
        conflict={editorConflict}
        localDraft={editorValue}
        onApplyMine={applyEditorConflictDraft}
        onKeepAuthoritative={keepEditorConflictAuthoritative}
        onKeepDraftLocal={keepEditorConflictDraftLocal}
        onReview={reviewEditorConflict}
      />
    );
  }, [
    applyEditorConflictDraft,
    editorConflict,
    editorValue,
    keepEditorConflictAuthoritative,
    keepEditorConflictDraftLocal,
    reviewEditorConflict,
  ]);
}
