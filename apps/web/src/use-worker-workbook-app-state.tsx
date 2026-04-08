import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { useActorRef, useSelector } from "@xstate/react";
import {
  isWorkbookAgentCommandBundle,
  isWorkbookAgentPreviewSummary,
  type WorkbookAgentCommandBundle,
} from "@bilig/agent-api";
import type { CommitOp } from "@bilig/core";
import type { EditMovement, EditSelectionBehavior } from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, LiteralInput, Viewport } from "@bilig/protocol";
import { createWorkerRuntimeMachine } from "./runtime-machine.js";
import { resolveRuntimeConfig } from "./runtime-config.js";
import type { ZeroClient } from "./runtime-session.js";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import { loadPersistedSelection, persistSelection } from "./selection-persistence.js";
import { ProjectedViewportStore } from "./projected-viewport-store.js";
import { WorkbookEditorConflictBanner } from "./WorkbookEditorConflictBanner.js";
import {
  type EditingMode,
  type ParsedEditorInput,
  type WorkbookEditorConflict,
  type ZeroConnectionState,
  canAttemptRemoteSync,
  clampSelectionMovement,
  createNextSheetName,
  emptyCellSnapshot,
  normalizeSheetNameKey,
  parseEditorInput,
  parsedEditorInputEquals,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  parseSelectionRangeLabel,
  sameCellContent,
  toEditorValue,
  toResolvedValue,
} from "./worker-workbook-app-model.js";
import { useWorkbookSync } from "./use-workbook-sync.js";
import { useWorkbookToolbar } from "./use-workbook-toolbar.js";
import { useWorkbookPresence } from "./use-workbook-presence.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";
import { WorkbookSideRailTabs } from "./WorkbookSideRailTabs.js";
import { useWorkbookChangesPane } from "./use-workbook-changes-pane.js";
import { useWorkbookAgentPane } from "./use-workbook-agent-pane.js";

const workerRuntimeMachine = createWorkerRuntimeMachine();

interface LocalOnlyZeroSource {
  materialize(query: unknown): {
    data: unknown;
    addListener(listener: (value: unknown) => void): () => void;
    destroy(): void;
  };
  mutate(mutation: unknown): unknown;
}

function selectionViewport(selection: WorkerRuntimeSelection): Viewport {
  const parsed = parseCellAddress(selection.address, selection.sheetName);
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  };
}

export function useWorkerWorkbookAppState(input: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>;
  connectionState: ZeroConnectionState;
  zero?: ZeroClient;
}) {
  const { runtimeConfig, connectionState, zero } = input;
  const documentId = runtimeConfig.documentId;
  const zeroConfigured = Boolean(zero);
  const zeroSource = useMemo<LocalOnlyZeroSource>(
    () =>
      zero ??
      ({
        materialize(_query: unknown) {
          return {
            data: [],
            addListener(_listener: (value: unknown) => void) {
              return () => undefined;
            },
            destroy() {},
          };
        },
        mutate(_mutation: unknown) {
          return {};
        },
      } satisfies LocalOnlyZeroSource),
    [zero],
  );
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, []);
  const initialSelection = useMemo(() => loadPersistedSelection(documentId), [documentId]);
  const runtimeActorRef = useActorRef(workerRuntimeMachine, {
    input: {
      documentId,
      replicaId,
      persistState: runtimeConfig.persistState,
      connectionStateName: connectionState.name,
      ...(zero ? { zero } : {}),
      initialSelection,
    },
  });
  const runtimeController = useSelector(runtimeActorRef, (snapshot) => snapshot.context.controller);
  const workerHandle = useSelector(runtimeActorRef, (snapshot) => snapshot.context.handle);
  const runtimeState = useSelector(runtimeActorRef, (snapshot) => snapshot.context.runtimeState);
  const selection = useSelector(runtimeActorRef, (snapshot) => snapshot.context.selection);
  const runtimeError = useSelector(runtimeActorRef, (snapshot) => snapshot.context.error);
  const runtimeReady = Boolean(workerHandle);
  const workbookReady = runtimeReady;
  const emptySelectedCell = useMemo(
    () => emptyCellSnapshot(selection.sheetName, selection.address),
    [selection.address, selection.sheetName],
  );
  const [selectionLabel, setSelectionLabel] = useState("A1");
  const [zeroHealthReady, setZeroHealthReady] = useState(false);
  const [editorValue, setEditorValue] = useState("");
  const [editorSelectionBehavior, setEditorSelectionBehavior] =
    useState<EditSelectionBehavior>("select-all");
  const [editingMode, setEditingMode] = useState<EditingMode>("idle");
  const [editorConflict, setEditorConflict] = useState<WorkbookEditorConflict | null>(null);
  const selectionRef = useRef(selection);
  const workerHandleRef = useRef(workerHandle);
  const editorValueRef = useRef(editorValue);
  const editingModeRef = useRef(editingMode);
  const editorTargetRef = useRef(selection);
  const editorBaseSnapshotRef = useRef<CellSnapshot>(emptySelectedCell);
  const zeroRef = useRef<LocalOnlyZeroSource>(zeroSource);
  const connectionStateRef = useRef(connectionState.name);
  const visibleViewportRef = useRef<Viewport>(selectionViewport(selection));

  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);

  useEffect(() => {
    persistSelection(documentId, selection);
  }, [documentId, selection]);

  useEffect(() => {
    workerHandleRef.current = workerHandle;
  }, [workerHandle]);

  useEffect(() => {
    editorValueRef.current = editorValue;
  }, [editorValue]);

  useEffect(() => {
    editingModeRef.current = editingMode;
  }, [editingMode]);

  useEffect(() => {
    zeroRef.current = zeroSource;
  }, [zeroSource]);

  useEffect(() => {
    connectionStateRef.current = connectionState.name;
  }, [connectionState.name]);

  useEffect(() => {
    runtimeActorRef.send({
      type: "connection.changed",
      connectionStateName: connectionState.name,
    });
  }, [connectionState.name, runtimeActorRef]);

  useEffect(() => {
    if (!runtimeReady) {
      setZeroHealthReady(false);
      return;
    }
    if (
      connectionState.name === "disconnected" ||
      connectionState.name === "needs-auth" ||
      connectionState.name === "error" ||
      connectionState.name === "closed"
    ) {
      setZeroHealthReady(false);
      return;
    }

    let cancelled = false;
    const probe = async (): Promise<void> => {
      try {
        const response = await fetch("/zero/keepalive", { cache: "no-store" });
        if (response.ok) {
          if (!cancelled) {
            setZeroHealthReady(true);
          }
          return;
        }
      } catch {}
      if (!cancelled) {
        window.setTimeout(() => {
          void probe();
        }, 250);
      }
    };

    setZeroHealthReady(false);
    void probe();
    return () => {
      cancelled = true;
    };
  }, [connectionState.name, runtimeReady]);

  const writesAllowed = runtimeReady;
  const remoteSyncAvailable = canAttemptRemoteSync(connectionState.name);

  const columnWidths = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}),
      [workerHandle],
    ),
    () => workerHandle?.viewportStore.getColumnWidths(selection.sheetName),
    () => workerHandle?.viewportStore.getColumnWidths(selection.sheetName),
  );

  const selectedCell = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}),
      [workerHandle],
    ),
    () =>
      workerHandle?.viewportStore.peekCell(selection.sheetName, selection.address) ??
      emptySelectedCell,
    () => emptySelectedCell,
  );

  const reportRuntimeError = useCallback(
    (error: unknown) => {
      runtimeActorRef.send({
        type: "session.error",
        message: error instanceof Error ? error.message : String(error),
      });
    },
    [runtimeActorRef],
  );
  const clearRuntimeError = useCallback(() => {
    runtimeActorRef.send({ type: "error.clear" });
  }, [runtimeActorRef]);
  const { invokeMutation, invokeColumnWidthMutation } = useWorkbookSync({
    documentId,
    connectionStateName: connectionState.name,
    connectionStateRef,
    runtimeController,
    workerHandleRef,
    zeroRef,
    reportRuntimeError,
  });

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current;
      if (!active) {
        return selectedCell;
      }
      return active.viewportStore.getCell(nextSelection.sheetName, nextSelection.address);
    },
    [selectedCell],
  );

  const cloneLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => structuredClone(getLiveSelectedCell(nextSelection)),
    [getLiveSelectedCell],
  );

  const resetEditorConflictTracking = useCallback(
    (nextSelection = selectionRef.current) => {
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextSelection);
      setEditorConflict(null);
    },
    [cloneLiveSelectedCell],
  );

  const completeEditNavigation = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      if (!movement) {
        selectionRef.current = targetSelection;
        editorTargetRef.current = targetSelection;
        return targetSelection;
      }
      const nextAddress = clampSelectionMovement(
        targetSelection.address,
        targetSelection.sheetName,
        movement,
      );
      const nextSelection = { sheetName: targetSelection.sheetName, address: nextAddress };
      selectionRef.current = nextSelection;
      editorTargetRef.current = nextSelection;
      runtimeActorRef.send({ type: "selection.changed", selection: nextSelection });
      return nextSelection;
    },
    [runtimeActorRef],
  );

  const finishEditingWithAuthoritative = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      const nextSelection = completeEditNavigation(targetSelection, movement);
      const nextEditorValue = toEditorValue(cloneLiveSelectedCell(nextSelection));
      editorValueRef.current = nextEditorValue;
      setEditorValue(nextEditorValue);
      setEditorSelectionBehavior("select-all");
      editingModeRef.current = "idle";
      setEditingMode("idle");
      resetEditorConflictTracking(nextSelection);
    },
    [cloneLiveSelectedCell, completeEditNavigation, resetEditorConflictTracking],
  );

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = "select-all",
      mode: Exclude<EditingMode, "idle"> = "cell",
    ) => {
      if (!writesAllowed) {
        return;
      }
      const nextEditorValue = seed ?? toEditorValue(getLiveSelectedCell());
      const nextTarget = selectionRef.current;
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextTarget);
      setEditorConflict(null);
      editorValueRef.current = nextEditorValue;
      setEditorValue(nextEditorValue);
      setEditorSelectionBehavior(selectionBehavior);
      editorTargetRef.current = nextTarget;
      editingModeRef.current = mode;
      setEditingMode(mode);
    },
    [cloneLiveSelectedCell, getLiveSelectedCell, writesAllowed],
  );

  const applyParsedInput = useCallback(
    async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      if (parsed.kind === "formula") {
        await invokeMutation("setCellFormula", sheetName, address, parsed.formula);
        return;
      }
      if (parsed.kind === "clear") {
        await invokeMutation("clearCell", sheetName, address);
        return;
      }
      await invokeMutation("setCellValue", sheetName, address, parsed.value);
    },
    [invokeMutation],
  );

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      if (!writesAllowed) {
        return;
      }
      const targetSelection =
        editingModeRef.current === "idle" ? selectionRef.current : editorTargetRef.current;
      const nextValue =
        editingModeRef.current === "idle"
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : editorValueRef.current;
      const parsed = parseEditorInput(nextValue);
      const baseSnapshot = editorBaseSnapshotRef.current;
      const authoritativeSnapshot = cloneLiveSelectedCell(targetSelection);
      const draftMatchesAuthoritative = parsedEditorInputMatchesSnapshot(
        parsed,
        authoritativeSnapshot,
      );
      const draftMatchesBase = parsedEditorInputEquals(
        parsed,
        parsedEditorInputFromSnapshot(baseSnapshot),
      );

      if (!sameCellContent(baseSnapshot, authoritativeSnapshot) && !draftMatchesAuthoritative) {
        if (draftMatchesBase) {
          finishEditingWithAuthoritative(targetSelection, movement);
          return;
        }
        setEditorConflict({
          sheetName: targetSelection.sheetName,
          address: targetSelection.address,
          phase: "compare",
          baseSnapshot,
          authoritativeSnapshot,
        });
        return;
      }

      if (draftMatchesAuthoritative) {
        finishEditingWithAuthoritative(targetSelection, movement);
        return;
      }

      const nextSelection = completeEditNavigation(targetSelection, movement);
      setEditorSelectionBehavior("select-all");
      editingModeRef.current = "idle";
      setEditingMode("idle");
      resetEditorConflictTracking(nextSelection);
      void applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed).catch(
        reportRuntimeError,
      );
    },
    [
      applyParsedInput,
      cloneLiveSelectedCell,
      completeEditNavigation,
      finishEditingWithAuthoritative,
      getLiveSelectedCell,
      reportRuntimeError,
      resetEditorConflictTracking,
      writesAllowed,
    ],
  );

  const cancelEditor = useCallback(() => {
    const nextEditorValue = toEditorValue(getLiveSelectedCell());
    editorValueRef.current = nextEditorValue;
    setEditorValue(nextEditorValue);
    setEditorSelectionBehavior("select-all");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
    resetEditorConflictTracking();
  }, [getLiveSelectedCell, resetEditorConflictTracking]);

  const clearSelectedRange = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    const targetRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);
    editorValueRef.current = "";
    setEditorValue("");
    editorTargetRef.current = selectionRef.current;
    editingModeRef.current = "idle";
    setEditingMode("idle");
    resetEditorConflictTracking();
    void invokeMutation("clearRange", targetRange).catch((error: unknown) => {
      reportRuntimeError(error);
    });
  }, [
    invokeMutation,
    reportRuntimeError,
    resetEditorConflictTracking,
    selection.sheetName,
    selectionLabel,
    writesAllowed,
  ]);

  const clearSelectedCell = useCallback(() => {
    if (!writesAllowed) {
      return;
    }
    clearSelectedRange();
  }, [clearSelectedRange, writesAllowed]);

  const toggleBooleanCell = useCallback(
    (sheetName: string, address: string, nextValue: boolean) => {
      if (!writesAllowed) {
        return;
      }
      void applyParsedInput(sheetName, address, { kind: "value", value: nextValue }).catch(
        reportRuntimeError,
      );
    },
    [applyParsedInput, reportRuntimeError, writesAllowed],
  );

  const pasteIntoSelection = useCallback(
    (sheetName: string, startAddr: string, values: readonly (readonly string[])[]) => {
      const start = parseCellAddress(startAddr, sheetName);
      const ops: {
        kind: "upsertCell" | "deleteCell";
        sheetName: string;
        addr: string;
        formula?: string;
        value?: LiteralInput;
      }[] = [];
      values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((cellValue, colOffset) => {
          const address = formatAddress(start.row + rowOffset, start.col + colOffset);
          const parsed = parseEditorInput(cellValue);
          if (parsed.kind === "formula") {
            ops.push({
              kind: "upsertCell",
              sheetName,
              addr: address,
              formula: parsed.formula,
            });
            return;
          }
          if (parsed.kind === "clear") {
            ops.push({ kind: "deleteCell", sheetName, addr: address });
            return;
          }
          ops.push({
            kind: "upsertCell",
            sheetName,
            addr: address,
            value: parsed.value,
          });
        });
      });
      if (ops.length === 0) {
        return;
      }
      void invokeMutation("renderCommit", ops).catch(reportRuntimeError);
      setEditorSelectionBehavior("select-all");
      editorTargetRef.current = selectionRef.current;
      editingModeRef.current = "idle";
      setEditingMode("idle");
      resetEditorConflictTracking();
    },
    [invokeMutation, reportRuntimeError, resetEditorConflictTracking],
  );

  const fillSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("fillRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          resetEditorConflictTracking();
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, resetEditorConflictTracking],
  );

  const copySelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("copyRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          resetEditorConflictTracking();
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, resetEditorConflictTracking],
  );

  const moveSelectionRange = useCallback(
    (
      sourceStartAddr: string,
      sourceEndAddr: string,
      targetStartAddr: string,
      targetEndAddr: string,
    ) => {
      const targetSelection = selectionRef.current;
      const source = {
        sheetName: targetSelection.sheetName,
        startAddress: sourceStartAddr,
        endAddress: sourceEndAddr,
      };
      const target = {
        sheetName: targetSelection.sheetName,
        startAddress: targetStartAddr,
        endAddress: targetEndAddr,
      };
      void invokeMutation("moveRange", source, target)
        .then(() => {
          editorTargetRef.current = selectionRef.current;
          editingModeRef.current = "idle";
          setEditingMode("idle");
          resetEditorConflictTracking();
          return undefined;
        })
        .catch(reportRuntimeError);
    },
    [invokeMutation, reportRuntimeError, resetEditorConflictTracking],
  );

  const selectAddress = useCallback(
    (sheetName: string, address: string) => {
      const previousSelection = selectionRef.current;
      if (
        editingModeRef.current === "idle" &&
        previousSelection.sheetName === sheetName &&
        previousSelection.address === address
      ) {
        return;
      }
      if (editingModeRef.current !== "idle") {
        editorTargetRef.current = { sheetName, address };
        editingModeRef.current = "idle";
        setEditingMode("idle");
      }
      const nextSelection = { sheetName, address };
      if (previousSelection.sheetName !== sheetName) {
        visibleViewportRef.current = selectionViewport(nextSelection);
      }
      selectionRef.current = nextSelection;
      editorTargetRef.current = nextSelection;
      resetEditorConflictTracking(nextSelection);
      runtimeActorRef.send({ type: "selection.changed", selection: nextSelection });
    },
    [resetEditorConflictTracking, runtimeActorRef],
  );

  const handleEditorChange = useCallback(
    (next: string) => {
      if (editingModeRef.current === "idle") {
        editorBaseSnapshotRef.current = cloneLiveSelectedCell(selectionRef.current);
        setEditorConflict(null);
      }
      editorValueRef.current = next;
      setEditorValue(next);
      setEditingMode((current) => {
        const nextMode = current === "idle" ? "cell" : current;
        editingModeRef.current = nextMode;
        return nextMode;
      });
    },
    [cloneLiveSelectedCell],
  );

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
  }, [cloneLiveSelectedCell, editingMode, selectedCell, selection.address, selection.sheetName]);

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
  }, [cloneLiveSelectedCell]);

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
  }, [applyParsedInput, completeEditNavigation, reportRuntimeError, resetEditorConflictTracking]);

  const keepEditorConflictAuthoritative = useCallback(() => {
    finishEditingWithAuthoritative(editorTargetRef.current);
  }, [finishEditingWithAuthoritative]);

  const keepEditorConflictDraftLocal = useCallback(() => {
    setEditorConflict((current) => (current ? { ...current, phase: "badge" } : current));
  }, []);

  const isEditing = editingMode !== "idle";
  const isEditingCell = editingMode === "cell";
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell);
  const resolvedValue = toResolvedValue(selectedCell);
  const editorConflictBanner = useMemo(() => {
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
    editorValue,
    editorConflict,
    keepEditorConflictAuthoritative,
    keepEditorConflictDraftLocal,
    reviewEditorConflict,
  ]);
  const handleVisibleViewportChange = useCallback((viewport: Viewport) => {
    visibleViewportRef.current = viewport;
  }, []);
  const sheetNames = useMemo(
    () => [...(runtimeState?.sheetNames ?? [selection.sheetName])],
    [runtimeState?.sheetNames, selection.sheetName],
  );
  const collaborators = useWorkbookPresence({
    documentId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero: zeroSource,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });
  const { changeCount, changesPanel } = useWorkbookChangesPane({
    documentId,
    sheetNames,
    zero: zeroSource,
    enabled: runtimeReady && zeroConfigured,
    onJump: (sheetName, address) => {
      selectAddress(sheetName, address);
    },
  });
  const { agentPanel, agentError, clearAgentError, pendingCommandCount, previewRanges } =
    useWorkbookAgentPane({
      documentId,
      enabled: runtimeReady,
      getContext: () => ({
        selection: selectionRef.current,
        viewport: visibleViewportRef.current,
      }),
      previewBundle: async (bundle: WorkbookAgentCommandBundle) => {
        if (!runtimeController || !isWorkbookAgentCommandBundle(bundle)) {
          throw new Error("Workbook runtime is not ready for agent preview");
        }
        const value = await runtimeController.invoke("previewAgentCommandBundle", bundle);
        if (!isWorkbookAgentPreviewSummary(value)) {
          throw new Error("Worker returned an invalid workbook agent preview");
        }
        return value;
      },
    });
  const selectedStyle = workerHandle?.viewportStore.getCellStyle(selectedCell.styleId);
  const selectionRange = parseSelectionRangeLabel(selectionLabel, selection.sheetName);

  const subscribeViewport = useCallback(
    (
      sheetName: string,
      viewport: Parameters<ProjectedViewportStore["subscribeViewport"]>[1],
      listener: Parameters<ProjectedViewportStore["subscribeViewport"]>[2],
    ) => {
      if (!runtimeController) {
        return () => {};
      }
      return runtimeController.subscribeViewport(sheetName, viewport, listener);
    },
    [runtimeController],
  );

  const {
    headerStatus: toolbarHeaderStatus,
    ribbon,
    selectionStatus,
    statusModeLabel,
  } = useWorkbookToolbar({
    connectionStateName: connectionState.name,
    runtimeReady,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    invokeMutation,
    selectionRange,
    selection,
    selectionLabel,
    selectedCell,
    selectedStyle,
    writesAllowed,
  });

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
    [invokeMutation, reportRuntimeError, selectAddress, sheetNames],
  );

  const headerStatus = useMemo(() => {
    return (
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
    );
  }, [collaborators, selectAddress, toolbarHeaderStatus]);

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
    clearAgentError,
    clearRuntimeError,
    agentPanel,
    beginEditing,
    cancelEditor,
    clearSelectedCell,
    columnWidths,
    commitEditor,
    copySelectionRange,
    createSheet,
    changesPanel,
    editorConflictBanner,
    editorSelectionBehavior,
    fillSelectionRange,
    handleEditorChange,
    headerStatus,
    handleVisibleViewportChange,
    invokeColumnWidthMutation,
    isEditing,
    isEditingCell,
    moveSelectionRange,
    pasteIntoSelection,
    previewRanges,
    remoteSyncAvailable,
    renameSheet,
    reportRuntimeError,
    resolvedValue,
    ribbon,
    runtimeError,
    runtimeReady,
    selectAddress,
    selectedCell,
    selection,
    selectionStatus,
    setSelectionLabel,
    sheetNames,
    sideRail,
    statusModeLabel,
    subscribeViewport,
    toggleBooleanCell,
    visibleEditorValue,
    workbookReady,
    workerHandle,
    writesAllowed,
    zeroConfigured,
  };
}
